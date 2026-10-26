#pragma once

#include "excel_address.h"
#include "excel_row.h"

#include <cstddef>
#include <map>
#include <string>
#include <utility>

// Excel (sc-office): a sheet, holding the rows that have cells.
namespace sc::office::excel {
    // A sheet: the rows that have cells, in order.
    class sheet {
    public:
        explicit sheet(std::string name) : name_(std::move(name)) {}

        [[nodiscard]] const std::string &name() const { return name_; }

        [[nodiscard]] bool contains(address where) const;
        // The cell; std::out_of_range when there's none.
        [[nodiscard]] const cell &at(address where) const;
        [[nodiscard]] cell &at(address where);
        // The cell, added (empty) when there's none, as std::map does. Throws std::out_of_range
        // beyond XFD1048576. On a const sheet it's at().
        cell &operator[](address where);
        [[nodiscard]] const cell &operator[](const address where) const { return at(where); }
        // Removes the cell, and its row when that was its last: 1, or 0 when there's none (as
        // std::map does).
        std::size_t erase(address where);

        [[nodiscard]] auto begin() { return detail::values_iterator{rows_.begin()}; }
        [[nodiscard]] auto end() { return detail::values_iterator{rows_.end()}; }
        [[nodiscard]] auto begin() const { return detail::values_iterator{rows_.cbegin()}; }
        [[nodiscard]] auto end() const { return detail::values_iterator{rows_.cend()}; }
        // The number of rows that have cells.
        [[nodiscard]] std::size_t size() const { return rows_.size(); }
        [[nodiscard]] bool empty() const { return rows_.empty(); }

    private:
        std::string name_;
        std::map<std::size_t, row> rows_;
    };
}
