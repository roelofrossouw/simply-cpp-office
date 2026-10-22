#include "office.h"

#include <compression.h>
#include <core.h>
#include <xml.h>

#include <algorithm>
#include <map>
#include <set>
#include <stdexcept>

namespace sc::office {
    namespace {
        // An attribute by its local name, whatever its namespace prefix ("r:embed" is usual).
        std::string attribute_named(const web::element &element, const std::string_view local) {
            for (const auto &each: element.attributes()) {
                const std::string_view name = each.name;
                const auto colon = name.find(':');
                if ((colon == std::string_view::npos ? name : name.substr(colon + 1)) == local) return std::string{each.value};
            }
            return {};
        }

        // The relationship ids of the pictures placed in the text, in order: DrawingML's
        // <a:blip r:embed> and the older VML's <v:imagedata r:id>.
        void placed(const web::element &element, std::vector<std::string> &ids) {
            if (element.is("blip")) {
                if (auto id = attribute_named(element, "embed"); !id.empty()) ids.push_back(std::move(id));
            } else if (element.is("imagedata")) {
                if (auto id = attribute_named(element, "id"); !id.empty()) ids.push_back(std::move(id));
            }
            for (const auto &child: element.children()) placed(child, ids);
        }

        // word/_rels/document.xml.rels: each picture's relationship id and its path in the archive.
        std::map<std::string, std::string> picture_targets(const compression::zip_reader &docx) {
            std::map<std::string, std::string> targets;
            if (!docx.contains("word/_rels/document.xml.rels")) return targets;
            const auto rels = web::document::parse(docx.read("word/_rels/document.xml.rels"));
            for (const auto &relationship: rels.root().children("Relationship")) {
                if (relationship["TargetMode"] == "External" || !relationship["Type"].ends_with("/image")) continue;
                std::string target{relationship["Target"]};
                target = target.starts_with('/') ? target.substr(1) : "word/" + target; // relative to word/
                targets[std::string{relationship["Id"]}] = target;
            }
            return targets;
        }

        std::vector<picture> pictures_of(const compression::zip_reader &docx) {
            if (!docx.contains("word/document.xml")) throw std::runtime_error("Word: not a .docx (no word/document.xml)");
            const auto targets = picture_targets(docx);
            std::vector<std::string> ids;
            placed(web::document::parse(docx.read("word/document.xml")).root(), ids);

            std::vector<picture> out;
            std::set<std::string> taken;
            const auto take = [&](const std::string &path) {
                if (!taken.insert(path).second || !docx.contains(path)) return;
                out.push_back({std::filesystem::path{path}.filename().string(), docx.read(path)});
            };
            for (const auto &id: ids)
                if (const auto found = targets.find(id); found != targets.end()) take(found->second);
            for (const auto &entry: docx.entries()) // the rest: headers, footers, not placed
                if (entry.name.starts_with("word/media/") && !entry.folder) take(entry.name);
            return out;
        }
    }

    bool picture::raster() const {
        static const std::set<std::string> rasters{".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tif", ".tiff", ".webp"};
        return rasters.contains(lowercase(std::filesystem::path{name}.extension().string()));
    }

    std::vector<picture> word::images(const std::filesystem::path &docx) {
        return pictures_of(compression::zip_reader{docx});
    }

    std::vector<picture> word::images_in(std::string docx) {
        return pictures_of(compression::zip_reader::from_memory(std::move(docx)));
    }
}
