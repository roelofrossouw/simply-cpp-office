#include "excel_sheet.h"

#include <cmath>
#include <stdexcept>

namespace sc::office::excel {
    bool sheet::contains(const address where) const {
        const auto found = rows_.find(where.row);
        return found != rows_.end() && found->second.contains(where.column);
    }

    const cell &sheet::at(const address where) const {
        if (const auto found = rows_.find(where.row); found != rows_.end() && found->second.contains(where.column))
            return found->second.at(where.column);
        throw std::out_of_range("excel: no cell " + where.name() + " in sheet \"" + name_ + "\"");
    }

    cell &sheet::at(const address where) { return const_cast<cell &>(std::as_const(*this).at(where)); }

    cell &sheet::operator[](const address where) {
        if (where.row >= max_rows || where.column >= max_columns) throw std::out_of_range("excel: cell beyond XFD1048576");
        return rows_.try_emplace(where.row, where.row).first->second[where.column];
    }

    std::size_t sheet::erase(const address where) {
        const auto found = rows_.find(where.row);
        if (found == rows_.end()) return 0;
        const auto removed = found->second.erase(where.column);
        if (found->second.empty()) rows_.erase(found);
        return removed;
    }

    std::optional<double> sheet::width(const std::size_t column) const {
        if (const auto found = widths_.find(column); found != widths_.end()) return found->second;
        return std::nullopt;
    }

    void sheet::width(const std::size_t column, const std::optional<double> characters) {
        if (column >= max_columns) throw std::out_of_range("excel: a column beyond XFD");
        if (!characters) {
            widths_.erase(column);
            return;
        }
        if (!std::isfinite(*characters) || *characters < 0 || *characters > 255)
            throw std::out_of_range("excel: a column's width is from 0 to 255 characters");
        widths_[column] = *characters;
    }
}
