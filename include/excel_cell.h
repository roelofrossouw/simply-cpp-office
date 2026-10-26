#pragma once

#include "excel_address.h"

#include <cstddef>
#include <string>
#include <utility>
#include <variant>

// Excel (sc-office): a cell, with its place, value and formula.
namespace sc::office::excel {
    // A cell's value: empty, a number, text or true/false. Dates are numbers (days since
    // 1899-12-30, as Excel keeps them); an error ("#DIV/0!") reads as its text.
    using value = std::variant<std::monostate, double, std::string, bool>;

    // A cell: where it is, its value, and its formula if it has one.
    struct cell {
        std::size_t row = 0;    // from 0
        std::size_t column = 0; // from 0
        // With a formula, the value it last calculated (empty when it was never calculated).
        excel::value value;
        // "=C6" (with the =, as Excel shows it); empty: none.
        std::string formula;

        // Sets the value: sheet["A1"] = "Item".
        cell &operator=(excel::value what) {
            value = std::move(what);
            return *this;
        }

        // Its name: "B3".
        [[nodiscard]] std::string name() const { return address{row, column}.name(); }
        // Whether it has neither a value nor a formula.
        [[nodiscard]] bool empty() const { return std::holds_alternative<std::monostate>(value) && formula.empty(); }
    };
}
