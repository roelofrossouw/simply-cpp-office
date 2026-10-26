#pragma once

#include <cstddef>
#include <string>
#include <string_view>

// Excel (sc-office): a cell's place in a sheet, by row and column or by Excel's name ("B3").
namespace sc::office::excel {
    // Excel's limits: XFD1048576 is the last cell.
    constexpr std::size_t max_rows = 1048576;
    constexpr std::size_t max_columns = 16384;

    // A cell's place, counted from 0: {0, 0} is A1, {2, 1} is B3. Or by its name: "B3".
    struct address {
        std::size_t row = 0;
        std::size_t column = 0;

        address() = default;
        address(const std::size_t row, const std::size_t column) : row(row), column(column) {}
        // "B3" (or "b3") as {2, 1}. Throws std::invalid_argument for anything else ("3B", "A0",
        // "XFE1").
        address(std::string_view name);
        address(const char *name) : address(std::string_view{name}) {}
        address(const std::string &name) : address(std::string_view{name}) {}

        // Its name: {2, 1} as "B3".
        [[nodiscard]] std::string name() const;

        bool operator==(const address &) const = default;
    };

    // A column's letters: 0 as "A", 26 as "AA".
    [[nodiscard]] std::string column_name(std::size_t column);
    // A column's number from its letters: "A" (or "a") as 0, "AA" as 26. Throws
    // std::invalid_argument for anything else ("A1", "XFE").
    [[nodiscard]] std::size_t column_index(std::string_view letters);
}
