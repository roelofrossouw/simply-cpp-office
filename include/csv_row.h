#pragma once

#include "csv_cell.h"
#include "excel_iterator.h"

#include <cstddef>
#include <map>

// CSV (sc-office): a row of a sheet, holding the cells that are set, as an Excel row does.
namespace sc::office::csv {
    // A row: the cells in it that are set, in column order.
    class row {
    public:
        explicit row(const std::size_t index) : index_(index) {}

        // Its place, from 0.
        [[nodiscard]] std::size_t index() const { return index_; }

        [[nodiscard]] bool contains(const std::size_t column) const { return cells_.contains(column); }
        // The cell in that column; std::out_of_range when there's none.
        [[nodiscard]] const cell &at(std::size_t column) const;
        [[nodiscard]] cell &at(std::size_t column);
        // The cell in that column, added (empty) when there's none, as std::map does. On a const
        // row it's at().
        cell &operator[](std::size_t column);
        [[nodiscard]] const cell &operator[](const std::size_t column) const { return at(column); }
        // Removes the cell in that column: 1, or 0 when there's none (as std::map does).
        std::size_t erase(const std::size_t column) { return cells_.erase(column); }

        [[nodiscard]] auto begin() { return excel::detail::values_iterator{cells_.begin()}; }
        [[nodiscard]] auto end() { return excel::detail::values_iterator{cells_.end()}; }
        [[nodiscard]] auto begin() const { return excel::detail::values_iterator{cells_.cbegin()}; }
        [[nodiscard]] auto end() const { return excel::detail::values_iterator{cells_.cend()}; }
        [[nodiscard]] std::size_t size() const { return cells_.size(); }
        [[nodiscard]] bool empty() const { return cells_.empty(); }

    private:
        std::size_t index_;
        std::map<std::size_t, cell> cells_;
    };
}
