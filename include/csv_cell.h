#pragma once

#include "excel_address.h"
#include "excel_cell.h"

#include <cstddef>
#include <string>
#include <utility>
#include <variant>

// CSV (sc-office): a cell, with its place and value. The values and addresses are Excel's, so code
// that only uses data works the same on both.
namespace sc::office::csv {
    using value = excel::value;
    using address = excel::address;

    // A cell: where it is and its value.
    struct cell {
        std::size_t row = 0;    // from 0
        std::size_t column = 0; // from 0
        csv::value value;

        // Sets the value: sheet["A1"] = "Item".
        cell &operator=(csv::value what) {
            value = std::move(what);
            return *this;
        }

        // Its name: "B3".
        [[nodiscard]] std::string name() const { return address{row, column}.name(); }
        // Whether it has no value.
        [[nodiscard]] bool empty() const { return std::holds_alternative<std::monostate>(value); }
    };
}
