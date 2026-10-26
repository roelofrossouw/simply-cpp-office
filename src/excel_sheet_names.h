#pragma once

// Sheet names as Excel takes them: for workbook::add() and write(). Not installed.

#include <core.h>

#include <stdexcept>
#include <string>
#include <string_view>

namespace sc::office::excel::detail {
    // Whether two sheet names are the same to Excel: ASCII letters in either case.
    inline bool same_name(const std::string_view a, const std::string_view b) { return lowercase(a) == lowercase(b); }

    // Throws std::invalid_argument for a name Excel doesn't take.
    inline void check_name(const std::string_view name) {
        if (name.empty()) throw std::invalid_argument("excel: a sheet needs a name");
        // UTF-8 characters, not bytes.
        std::size_t characters = 0;
        for (const char c: name) characters += (static_cast<unsigned char>(c) & 0xC0) != 0x80;
        if (characters > 31) throw std::invalid_argument("excel: sheet name \"" + std::string{name} + "\" is longer than 31 characters");
        if (name.find_first_of(":\\/?*[]") != std::string_view::npos)
            throw std::invalid_argument("excel: sheet name \"" + std::string{name} + "\" has one of : \\ / ? * [ ]");
        if (name.front() == '\'' || name.back() == '\'')
            throw std::invalid_argument("excel: sheet name \"" + std::string{name} + "\" starts or ends with '");
    }
}
