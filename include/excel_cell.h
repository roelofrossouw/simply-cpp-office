#pragma once

#include "excel_address.h"
#include "excel_style.h"

#include <cstddef>
#include <string>
#include <utility>
#include <variant>

// Excel (sc-office): a cell, with its place, value, formula and style.
namespace sc::office::excel {
    // A cell's value: empty, a number, text or true/false. Dates are numbers (days since
    // 1899-12-30, as Excel keeps them); an error ("#DIV/0!") reads as its text.
    using value = std::variant<std::monostate, double, std::string, bool>;

    // A cell: where it is, its value, its formula if it has one, and how it looks.
    struct cell {
        std::size_t row = 0;    // from 0
        std::size_t column = 0; // from 0
        // With a formula, the value it last calculated (empty when it was never calculated).
        excel::value value;
        // "=C6" (with the =, as Excel shows it); empty: none.
        std::string formula;
        // Its font and background; a default style is Excel's plain one.
        excel::style style;

        // Sets the value, keeping the formula and style: sheet["A1"] = "Item".
        cell &operator=(excel::value what) {
            value = std::move(what);
            return *this;
        }

        // Its name: "B3".
        [[nodiscard]] std::string name() const { return address{row, column}.name(); }
        // Whether it has neither a value nor a formula (it may still have a style).
        [[nodiscard]] bool empty() const { return std::holds_alternative<std::monostate>(value) && formula.empty(); }
    };
}
