#pragma once

#include "excel_address.h"

#include <string>
#include <string_view>

// Excel (sc-office): a block of cells, from its top left to its bottom right ("A1:D3").
namespace sc::office::excel {
    struct range {
        address first; // top left
        address last;  // bottom right

        range() = default;
        // The block between two corners, any two opposite ones: {"D3", "A1"} is A1:D3.
        range(address one, address other);
        // "A1:D3" (or "a1:d3"); "B2" alone is that one cell. Throws std::invalid_argument for
        // anything else.
        range(std::string_view name);
        range(const char *name) : range(std::string_view{name}) {}
        range(const std::string &name) : range(std::string_view{name}) {}

        // Its name: "A1:D3" ("B2" for one cell).
        [[nodiscard]] std::string name() const;
        [[nodiscard]] bool contains(address where) const;
        // Whether they have a cell in common.
        [[nodiscard]] bool overlaps(const range &other) const;
        // The number of cells in it.
        [[nodiscard]] std::size_t size() const { return (last.row - first.row + 1) * (last.column - first.column + 1); }

        bool operator==(const range &) const = default;
    };
}
