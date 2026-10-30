#pragma once

#include "csv_row.h"
#include "excel_sheet.h"

#include <cstddef>
#include <map>
#include <string>
#include <utility>

// CSV (sc-office): a sheet of values, with the same access as an Excel sheet: by "B3" or {2, 1},
// rows and cells in order, cells added as std::map adds them.
namespace sc::office::csv {
    // A sheet: the rows that have cells, in order.
    class sheet {
    public:
        // A CSV file has one sheet and no name for it: read() names it after the file ("prices"
        // for prices.csv).
        explicit sheet(std::string name = {}) : name_(std::move(name)) {}
        // An Excel sheet's values and name; formulas give the value they last calculated, styles
        // and layout are left behind.
        explicit sheet(const excel::sheet &from);

        [[nodiscard]] const std::string &name() const { return name_; }

        [[nodiscard]] bool contains(address where) const;
        // The cell; std::out_of_range when there's none.
        [[nodiscard]] const cell &at(address where) const;
        [[nodiscard]] cell &at(address where);
        // The cell, added (empty) when there's none, as std::map does. On a const sheet it's at().
        cell &operator[](address where);
        [[nodiscard]] const cell &operator[](const address where) const { return at(where); }
        // Removes the cell, and its row when that was its last: 1, or 0 when there's none (as
        // std::map does).
        std::size_t erase(address where);

        [[nodiscard]] auto begin() { return excel::detail::values_iterator{rows_.begin()}; }
        [[nodiscard]] auto end() { return excel::detail::values_iterator{rows_.end()}; }
        [[nodiscard]] auto begin() const { return excel::detail::values_iterator{rows_.cbegin()}; }
        [[nodiscard]] auto end() const { return excel::detail::values_iterator{rows_.cend()}; }
        // The number of rows that have cells.
        [[nodiscard]] std::size_t size() const { return rows_.size(); }
        [[nodiscard]] bool empty() const { return rows_.empty(); }

        // The same values as an Excel sheet, named name (this sheet's, or "Sheet1" when it has
        // none): to write as an .xlsx, or style.
        [[nodiscard]] excel::sheet to_excel(std::string name = {}) const;

    private:
        std::string name_;
        std::map<std::size_t, row> rows_;
    };
}
