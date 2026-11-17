// A Word document's text (word::text): paragraphs as lines, list numbers as Word shows them, tables
// as rows of cells. The XML is matched by local name: documents differ in their prefixes.
#include <sc/office/office.h>

#include <sc/archive/sc_archive.h>
#include <sc/web/xml.h>

#include <map>
#include <optional>
#include <stdexcept>

namespace sc::office {
    namespace {
        // An attribute by its local name, whatever its prefix ("w:val").
        std::optional<std::string> attribute_of(const web::element &element, const std::string_view local) {
            for (const auto &each: element.attributes()) {
                const std::string_view name = each.name;
                const auto colon = name.find(':');
                if ((colon == std::string_view::npos ? name : name.substr(colon + 1)) == local) return std::string{each.value};
            }
            return std::nullopt;
        }

        std::optional<web::element> child_of(const std::optional<web::element> &element, const std::string_view local) {
            if (!element) return std::nullopt;
            for (const auto &child: element->children())
                if (child.is(local)) return child;
            return std::nullopt;
        }

        // A child's w:val ("<w:numId w:val="3"/>").
        std::optional<std::string> value_of(const std::optional<web::element> &element, const std::string_view local) {
            const auto child = child_of(element, local);
            return child ? attribute_of(*child, "val") : std::nullopt;
        }

        int number_of(const std::optional<std::string> &text, const int fallback) {
            if (!text) return fallback;
            try {
                return std::stoi(*text);
            } catch (const std::exception &) {
                return fallback;
            }
        }

        // Every element of a local name, in document order.
        void all_named(const web::element &element, const std::string_view local, std::vector<web::element> &out) {
            for (const auto &child: element.children()) {
                if (child.is(local)) out.push_back(child);
                all_named(child, local, out);
            }
        }

        std::string letters(const int number) {
            if (number < 1) return std::to_string(number);
            return std::string(static_cast<std::size_t>((number - 1) / 26 + 1), static_cast<char>('a' + (number - 1) % 26));
        }

        std::string roman(int number) {
            static constexpr std::pair<const char *, int> numerals[]{{"M", 1000}, {"CM", 900}, {"D", 500}, {"CD", 400}, {"C", 100},
                                                                     {"XC", 90}, {"L", 50}, {"XL", 40}, {"X", 10}, {"IX", 9},
                                                                     {"V", 5}, {"IV", 4}, {"I", 1}};
            std::string out;
            for (const auto &[letters_of, value]: numerals)
                for (; number >= value; number -= value) out += letters_of;
            return out;
        }

        std::string formatted(const int number, const std::string &format) {
            if (format == "lowerLetter") return letters(number);
            if (format == "upperLetter") {
                auto out = letters(number);
                for (auto &c: out) c = static_cast<char>(std::toupper(static_cast<unsigned char>(c)));
                return out;
            }
            if (format == "lowerRoman") {
                auto out = roman(number);
                for (auto &c: out) c = static_cast<char>(std::tolower(static_cast<unsigned char>(c)));
                return out;
            }
            if (format == "upperRoman") return roman(number);
            if (format == "decimalZero") return number >= 0 && number < 10 ? "0" + std::to_string(number) : std::to_string(number);
            return std::to_string(number);
        }

        // The text of a document's parts, with the numbering its lists and styles give.
        class reader {
        public:
            reader(const std::optional<web::document> &numbering, const std::optional<web::document> &styles) {
                if (numbering) read_numbering(numbering->root());
                if (styles) read_styles(styles->root());
            }

            void walk(const web::element &element) {
                for (const auto &child: element.children()) {
                    const auto name = child.name();
                    // Left out: the copies kept for older readers, deleted text, field codes, properties
                    if (name == "Fallback" || name == "del" || name == "delText" || name == "instrText" || name == "pPr" || name == "rPr" ||
                        name == "sdtPr" || name == "sdtEndPr")
                        continue;
                    if (name == "t") {
                        text_ += child.text();
                    } else if (name == "tab" || name == "ptab") {
                        text_ += '\t';
                    } else if (name == "br" || name == "cr") {
                        text_ += '\n';
                    } else if (name == "noBreakHyphen") {
                        text_ += '-';
                    } else if (name == "p") {
                        text_ += number(child);
                        walk(child);
                        text_ += '\n';
                    } else if (name == "tc") {
                        walk(child);
                        while (!text_.empty() && text_.back() == '\n') text_.pop_back();
                        text_ += " | ";
                    } else if (name == "tr") {
                        walk(child);
                        if (text_.ends_with(" | ")) text_.resize(text_.size() - 3);
                        text_ += '\n';
                    } else {
                        walk(child);
                    }
                }
            }

            // Spacing tidied: no space before a line end, at most one blank line, none at either end.
            [[nodiscard]] std::string result() const {
                std::string out;
                out.reserve(text_.size());
                for (const char c: text_) {
                    if (c == '\n') {
                        while (!out.empty() && (out.back() == ' ' || out.back() == '\t')) out.pop_back();
                        if (out.ends_with("\n\n")) continue;
                    }
                    out += c;
                }
                const auto first = out.find_first_not_of(" \t\n");
                if (first == std::string::npos) return {};
                const auto last = out.find_last_not_of(" \t\n");
                return out.substr(first, last - first + 1);
            }

            void separate() {
                if (!text_.empty()) text_ += "\n\n";
            }

        private:
            struct level {
                int start = 1;
                std::string format = "decimal", text;
                bool legal = false;
                std::string style; // the paragraph style linked to this level (w:pStyle)
            };
            struct list {
                std::string abstract;
                std::map<int, int> starts; // ilvl -> startOverride
            };
            struct style {
                std::optional<std::string> num_id, ilvl, based_on;
            };

            void read_numbering(const web::element &root) {
                std::vector<web::element> abstracts, lists;
                all_named(root, "abstractNum", abstracts);
                all_named(root, "num", lists);
                for (const auto &abstract: abstracts) {
                    const auto id = attribute_of(abstract, "abstractNumId").value_or("");
                    for (const auto &each: abstract.children()) {
                        if (!each.is("lvl")) continue;
                        const std::optional<web::element> lvl{each};
                        level made;
                        made.start = number_of(value_of(lvl, "start"), 1);
                        made.format = value_of(lvl, "numFmt").value_or("decimal");
                        made.text = value_of(lvl, "lvlText").value_or("");
                        made.legal = child_of(lvl, "isLgl").has_value();
                        made.style = value_of(lvl, "pStyle").value_or("");
                        levels_[id][number_of(attribute_of(each, "ilvl"), 0)] = std::move(made);
                    }
                }
                for (const auto &each: lists) {
                    const std::optional<web::element> num{each};
                    list made;
                    made.abstract = value_of(num, "abstractNumId").value_or("");
                    for (const auto &override_: each.children()) {
                        if (!override_.is("lvlOverride")) continue;
                        if (const auto start = value_of(std::optional{override_}, "startOverride"))
                            made.starts[number_of(attribute_of(override_, "ilvl"), 0)] = number_of(start, 1);
                    }
                    lists_[attribute_of(each, "numId").value_or("")] = std::move(made);
                }
            }

            void read_styles(const web::element &root) {
                std::vector<web::element> all;
                all_named(root, "style", all);
                for (const auto &each: all) {
                    const std::optional<web::element> element{each};
                    const auto numbering = child_of(child_of(element, "pPr"), "numPr");
                    styles_[attribute_of(each, "styleId").value_or("")] = {value_of(numbering, "numId"), value_of(numbering, "ilvl"),
                                                                            value_of(element, "basedOn")};
                }
            }

            // A paragraph's list number as Word shows it ("4.1. "), counted on; empty outside a list,
            // and for bullets.
            std::string number(const web::element &paragraph) {
                const auto properties = child_of(std::optional{paragraph}, "pPr");
                const auto direct = child_of(properties, "numPr");
                auto list_id = value_of(direct, "numId");
                auto ilvl = value_of(direct, "ilvl");
                // Through its style, and the styles that one is based on
                auto style_id = value_of(properties, "pStyle");
                std::optional<std::string> linked;
                for (int seen = 0; (!list_id || !ilvl) && style_id && styles_.contains(*style_id) && seen < 10; ++seen) {
                    const auto &each = styles_.at(*style_id);
                    if (!list_id) list_id = each.num_id;
                    if (!ilvl) ilvl = each.ilvl;
                    if (!linked) linked = style_id;
                    style_id = each.based_on;
                }
                if (!list_id || *list_id == "0" || !lists_.contains(*list_id)) return {};
                const auto &which = lists_.at(*list_id);
                const auto found = levels_.find(which.abstract);
                if (found == levels_.end()) return {};
                const auto &levels = found->second;
                int at = 0;
                if (ilvl) {
                    at = number_of(ilvl, 0);
                } else {
                    // A style's level: the list level linked to that style
                    for (const auto &[index, each]: levels)
                        if (linked && each.style == *linked) at = index;
                }
                if (!levels.contains(at)) return {};
                const auto start = [&](const int index) {
                    if (const auto override_ = which.starts.find(index); override_ != which.starts.end()) return override_->second;
                    if (const auto each = levels.find(index); each != levels.end()) return each->second.start;
                    return 1;
                };
                // Lists of one abstract number count together; a level restarts the levels below it
                auto &counter = counters_[which.abstract];
                counter[at] = counter.contains(at) ? counter[at] + 1 : start(at);
                for (auto it = counter.upper_bound(at); it != counter.end();) it = counter.erase(it);
                const auto &format = levels.at(at);
                if (format.format == "bullet" || format.format == "none") return {};
                std::string out;
                for (std::size_t i = 0; i < format.text.size(); ++i) {
                    const char c = format.text[i];
                    if (c == '%' && i + 1 < format.text.size() && format.text[i + 1] >= '1' && format.text[i + 1] <= '9') {
                        const int index = format.text[i + 1] - '1';
                        const int value = counter.contains(index) ? counter.at(index) : start(index);
                        const auto level_format = levels.contains(index) ? levels.at(index).format : std::string{"decimal"};
                        out += formatted(value, format.legal ? "decimal" : level_format);
                        ++i;
                    } else {
                        out += c;
                    }
                }
                if (out.find_first_not_of(" \t") == std::string::npos) return {};
                return out + " ";
            }

            std::map<std::string, std::map<int, level>> levels_; // abstractNumId -> ilvl -> level
            std::map<std::string, list> lists_;                  // numId -> list
            std::map<std::string, style> styles_;                // styleId -> its numbering, basedOn
            std::map<std::string, std::map<int, int>> counters_; // abstractNumId -> ilvl -> the current number
            std::string text_;
        };

        std::optional<web::document> part(const archive::reader &docx, const std::string &path) {
            if (!docx.contains(path)) return std::nullopt;
            try {
                return web::document::parse(docx.read(path));
            } catch (const std::exception &) {
                return std::nullopt; // a damaged part: read without it
            }
        }

        // word/_rels/document.xml.rels: the targets of the relationships of a type ("header", "footer").
        std::vector<std::string> targets_of(const archive::reader &docx, const std::string_view type) {
            std::vector<std::string> out;
            const auto rels = part(docx, "word/_rels/document.xml.rels");
            if (!rels) return out;
            for (const auto &relationship: rels->root().children("Relationship")) {
                if (!std::string_view{relationship["Type"]}.ends_with("/" + std::string{type})) continue;
                std::string target{relationship["Target"]};
                out.push_back(target.starts_with('/') ? target.substr(1) : "word/" + target);
            }
            return out;
        }

        std::string text_of(const archive::reader &docx, const word::text_options &options) {
            if (!docx.contains("word/document.xml")) throw std::runtime_error("Word: not a .docx (no word/document.xml)");
            const auto document = part(docx, "word/document.xml");
            if (!document) throw std::runtime_error("Word: word/document.xml can't be read");
            reader read{part(docx, "word/numbering.xml"), part(docx, "word/styles.xml")};
            read.walk(document->root());
            const auto add = [&](const std::string &path) {
                if (const auto extra = part(docx, path)) {
                    read.separate();
                    read.walk(extra->root());
                }
            };
            if (options.headers_footers) {
                for (const auto &path: targets_of(docx, "header")) add(path);
                for (const auto &path: targets_of(docx, "footer")) add(path);
            }
            if (options.notes) {
                add("word/footnotes.xml");
                add("word/endnotes.xml");
            }
            return read.result();
        }
    }

    std::string word::text(const std::filesystem::path &docx, const text_options &options) {
        return text_of(archive::open(docx), options);
    }

    std::string word::text_from(std::string docx, const text_options &options) {
        return text_of(archive::from_memory(std::move(docx)), options);
    }
}
