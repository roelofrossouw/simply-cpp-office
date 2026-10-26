#include "excel_address.h"

#include <cctype>
#include <stdexcept>

namespace sc::office::excel {
    address::address(const std::string_view name) {
        const auto wrong = [&] { return std::invalid_argument("excel: \"" + std::string{name} + "\" isn't a cell name like B3"); };
        std::size_t at = 0, letters = 0, number = 0;
        for (; at < name.size() && std::isalpha(static_cast<unsigned char>(name[at])); ++at) {
            letters = letters * 26 + static_cast<std::size_t>(std::toupper(static_cast<unsigned char>(name[at])) - 'A' + 1);
            if (letters > max_columns) throw wrong();
        }
        if (at == 0 || at == name.size()) throw wrong();
        for (; at < name.size(); ++at) {
            if (!std::isdigit(static_cast<unsigned char>(name[at]))) throw wrong();
            number = number * 10 + static_cast<std::size_t>(name[at] - '0');
            if (number > max_rows) throw wrong();
        }
        if (number == 0) throw wrong();
        row = number - 1;
        column = letters - 1;
    }

    std::string address::name() const { return column_name(column) + std::to_string(row + 1); }

    std::string column_name(std::size_t column) {
        std::string letters;
        for (++column; column > 0; column = (column - 1) / 26) letters.insert(letters.begin(), static_cast<char>('A' + (column - 1) % 26));
        return letters;
    }
}
