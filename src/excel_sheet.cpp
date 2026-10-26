#include "excel_sheet.h"

#include <algorithm>
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

    std::optional<sc::distance> sheet::height(const std::size_t row) const {
        if (const auto found = heights_.find(row); found != heights_.end()) return found->second;
        return std::nullopt;
    }

    void sheet::height(const std::size_t row, const std::optional<sc::distance> points) {
        if (row >= max_rows) throw std::out_of_range("excel: a row beyond 1048576");
        if (!points) {
            heights_.erase(row);
            return;
        }
        if (!std::isfinite(points->points()) || points->points() < 0 || points->points() > 409)
            throw std::out_of_range("excel: a row's height is from 0 to 409 points");
        heights_[row] = *points;
    }

    void sheet::merge(const range &cells) {
        if (cells.last.row >= max_rows || cells.last.column >= max_columns) throw std::out_of_range("excel: a range beyond XFD1048576");
        if (cells.first == cells.last) throw std::invalid_argument("excel: merging " + cells.name() + " needs more than one cell");
        for (const auto &each: merged_)
            if (each.overlaps(cells)) throw std::invalid_argument("excel: " + cells.name() + " overlaps " + each.name() + ", merged already");
        merged_.push_back(cells);
    }

    std::size_t sheet::unmerge(const range &cells) { return std::erase(merged_, cells); }

    void sheet::hidden(const std::size_t column, const bool hide) {
        if (column >= max_columns) throw std::out_of_range("excel: a column beyond XFD");
        if (hide) hidden_.insert(column);
        else hidden_.erase(column);
    }

    void sheet::freeze(const address top_left) {
        if (top_left.row >= max_rows || top_left.column >= max_columns) throw std::out_of_range("excel: a cell beyond XFD1048576");
        if (top_left == address{0, 0}) frozen_.reset();
        else frozen_ = top_left;
    }
}
