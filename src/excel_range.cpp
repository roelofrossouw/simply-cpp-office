#include "excel_range.h"

#include <algorithm>
#include <stdexcept>

namespace sc::office::excel {
    range::range(const address one, const address other)
        : first{std::min(one.row, other.row), std::min(one.column, other.column)},
          last{std::max(one.row, other.row), std::max(one.column, other.column)} {}

    range::range(const std::string_view name) {
        const auto colon = name.find(':');
        try {
            *this = colon == std::string_view::npos ? range{address{name}, address{name}} : range{address{name.substr(0, colon)}, address{name.substr(colon + 1)}};
        } catch (const std::invalid_argument &) {
            throw std::invalid_argument("excel: \"" + std::string{name} + "\" isn't a range like A1:D3");
        }
    }

    std::string range::name() const { return first == last ? first.name() : first.name() + ":" + last.name(); }

    bool range::contains(const address where) const {
        return where.row >= first.row && where.row <= last.row && where.column >= first.column && where.column <= last.column;
    }

    bool range::overlaps(const range &other) const {
        return first.row <= other.last.row && other.first.row <= last.row && first.column <= other.last.column && other.first.column <= last.column;
    }
}
