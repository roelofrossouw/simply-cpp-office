#pragma once

#include "excel_address.h"
#include "excel_range.h"
#include "excel_row.h"

#include <distance.h>

#include <cstddef>
#include <map>
#include <optional>
#include <string>
#include <string_view>
#include <vector>
#include <utility>

// Excel (sc-office): a sheet, holding the rows that have cells, its columns' widths, its rows'
// heights and its merged cells.
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

        // A row's height, in points (a plain number is points: Excel's unit); none: Excel's default
        // (15 for Calibri 11, more when the text needs it). By number from 0.
        [[nodiscard]] std::optional<sc::distance> height(std::size_t row) const;
        // Sets a row's height, from 0 to 409 points (0 hides it); std::nullopt gives it Excel's
        // default. A row with a height and no cells is kept. Throws std::out_of_range beyond row
        // 1048576 or 409 points.
        void height(std::size_t row, std::optional<sc::distance> points);
        // The rows given a height, by number.
        [[nodiscard]] const std::map<std::size_t, sc::distance> &heights() const { return heights_; }

        // Merges the cells of a range into one, shown as its top left cell (the others keep what
        // they hold, as Excel keeps it hidden): merge("A1:D1"). Throws std::invalid_argument for one
        // cell, or one that overlaps a merge already there; std::out_of_range beyond XFD1048576.
        void merge(const range &cells);
        // Undoes the merge of exactly that range: 1, or 0 when there's none (as std::map's erase).
        std::size_t unmerge(const range &cells);
        // The merged ranges, in the order they were merged (or read).
        [[nodiscard]] const std::vector<range> &merged() const { return merged_; }

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
        std::map<std::size_t, sc::distance> heights_;
        std::vector<range> merged_;
    };
}
