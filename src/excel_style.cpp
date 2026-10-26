#include "excel_style.h"

#include <array>
#include <stdexcept>
#include <string>
#include <utility>

namespace sc::office::excel {
    namespace {
        constexpr std::array<std::pair<underline, std::string_view>, 5> underlines{{
            {underline::none, "none"}, {underline::single, "single"}, {underline::doubled, "double"},
            {underline::single_accounting, "singleAccounting"}, {underline::double_accounting, "doubleAccounting"}}};

        constexpr std::array<std::pair<line, std::string_view>, 14> lines{{
            {line::none, "none"}, {line::thin, "thin"}, {line::medium, "medium"}, {line::thick, "thick"}, {line::hair, "hair"},
            {line::dotted, "dotted"}, {line::dashed, "dashed"}, {line::doubled, "double"}, {line::dash_dot, "dashDot"},
            {line::dash_dot_dot, "dashDotDot"}, {line::slant_dash_dot, "slantDashDot"}, {line::medium_dashed, "mediumDashed"},
            {line::medium_dash_dot, "mediumDashDot"}, {line::medium_dash_dot_dot, "mediumDashDotDot"}}};

        constexpr std::array<std::pair<horizontal, std::string_view>, 8> horizontals{{
            {horizontal::general, "general"}, {horizontal::left, "left"}, {horizontal::center, "center"}, {horizontal::right, "right"},
            {horizontal::fill, "fill"}, {horizontal::justify, "justify"}, {horizontal::center_across, "centerContinuous"},
            {horizontal::distributed, "distributed"}}};

        constexpr std::array<std::pair<vertical, std::string_view>, 5> verticals{{
            {vertical::bottom, "bottom"}, {vertical::top, "top"}, {vertical::center, "center"}, {vertical::justify, "justify"},
            {vertical::distributed, "distributed"}}};

        template<typename Kind, std::size_t N>
        std::string_view name_in(const std::array<std::pair<Kind, std::string_view>, N> &table, const Kind kind) {
            for (const auto &[each, name]: table)
                if (each == kind) return name;
            return "none";
        }

        template<typename Kind, std::size_t N>
        Kind named_in(const std::array<std::pair<Kind, std::string_view>, N> &table, const std::string_view name, const char *what) {
            for (const auto &[each, each_name]: table)
                if (each_name == name) return each;
            std::string known;
            for (const auto &[each, each_name]: table) known += (known.empty() ? "" : ", ") + std::string{each_name};
            throw std::invalid_argument("excel: \"" + std::string{name} + "\" isn't " + what + " (" + known + ")");
        }
    }

    std::string_view name(const underline kind) { return name_in(underlines, kind); }
    std::string_view name(const line kind) { return name_in(lines, kind); }
    underline underline_named(const std::string_view name) { return named_in(underlines, name, "an underline"); }
    line line_named(const std::string_view name) { return named_in(lines, name, "a border's line"); }
    std::string_view name(const horizontal kind) { return name_in(horizontals, kind); }
    std::string_view name(const vertical kind) { return name_in(verticals, kind); }
    horizontal horizontal_named(const std::string_view name) { return named_in(horizontals, name, "a horizontal alignment"); }
    vertical vertical_named(const std::string_view name) { return named_in(verticals, name, "a vertical alignment"); }
}
