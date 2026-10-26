#pragma once

#include "excel_address.h"
#include "excel_row.h"

#include <cstddef>
#include <map>
#include <optional>
#include <string>
#include <string_view>
#include <utility>

// Excel (sc-office): a sheet, holding the rows that have cells, and its columns' widths.
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

        // A column's width in characters (so many "0"s in the default font, as Excel shows it); none:
        // Excel's default (8.43). By number from 0, or letters: width("B").
        [[nodiscard]] std::optional<double> width(std::size_t column) const;
        [[nodiscard]] std::optional<double> width(const std::string_view column) const { return width(column_index(column)); }
        // Sets a column's width in characters, from 0 to 255; std::nullopt gives it Excel's default.
        // Throws std::out_of_range beyond column XFD or 255 characters.
        void width(std::size_t column, std::optional<double> characters);
        void width(const std::string_view column, const std::optional<double> characters) { width(column_index(column), characters); }
        // The columns given a width, by number.
        [[nodiscard]] const std::map<std::size_t, double> &widths() const { return widths_; }

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
        std::map<std::size_t, double> widths_;
    };
}
