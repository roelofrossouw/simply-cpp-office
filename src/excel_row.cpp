#include "excel_row.h"

#include <stdexcept>

namespace sc::office::excel {
    const cell &row::at(const std::size_t column) const {
        if (const auto found = cells_.find(column); found != cells_.end()) return found->second;
        throw std::out_of_range("excel: no cell " + address{index_, column}.name());
    }

    cell &row::at(const std::size_t column) { return const_cast<cell &>(std::as_const(*this).at(column)); }

    cell &row::operator[](const std::size_t column) {
        if (column >= max_columns) throw std::out_of_range("excel: no column beyond XFD");
        return cells_.try_emplace(column, cell{.row = index_, .column = column}).first->second;
    }
}
