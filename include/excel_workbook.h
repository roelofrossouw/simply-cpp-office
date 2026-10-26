#pragma once

#include "excel_sheet.h"

#include <cstddef>
#include <deque>
#include <filesystem>
#include <string>
#include <string_view>

// Excel (sc-office): a workbook, holding its sheets in order.
namespace sc::office::excel {
    // In excel.h: they fill a workbook's sheets.
    class workbook;
    workbook read(const std::filesystem::path &file);
    workbook read_from(std::string bytes);

    // A workbook: its sheets in order, found by name ignoring case, as Excel does.
    class workbook {
    public:
        // Adds an empty sheet. Throws std::invalid_argument for a name Excel doesn't take: empty,
        // longer than 31 characters, with any of : \ / ? * [ ], starting or ending with ', or
        // already used (ignoring case).
        sheet &add(std::string name);

        [[nodiscard]] bool contains(std::string_view name) const;
        // The sheet of that name; std::out_of_range when there's none.
        [[nodiscard]] sheet &at(std::string_view name);
        [[nodiscard]] const sheet &at(std::string_view name) const;
        // The sheet of that name, added when there's none (as std::map does); a const workbook
        // has no sheet to add, so there it's at().
        sheet &operator[](std::string_view name);
        [[nodiscard]] const sheet &operator[](const std::string_view name) const { return at(name); }

        // The sheets, in the workbook's order. A sheet stays where it is as more are added
        // (they're kept in a deque).
        [[nodiscard]] auto begin() { return sheets_.begin(); }
        [[nodiscard]] auto end() { return sheets_.end(); }
        [[nodiscard]] auto begin() const { return sheets_.begin(); }
        [[nodiscard]] auto end() const { return sheets_.end(); }
        [[nodiscard]] std::size_t size() const { return sheets_.size(); }
        [[nodiscard]] bool empty() const { return sheets_.empty(); }
        [[nodiscard]] sheet &front() { return sheets_.front(); }
        [[nodiscard]] const sheet &front() const { return sheets_.front(); }

    private:
        [[nodiscard]] const sheet *find(std::string_view name) const;

        std::deque<sheet> sheets_;

        friend workbook read(const std::filesystem::path &file);
        friend workbook read_from(std::string bytes);
    };
}
