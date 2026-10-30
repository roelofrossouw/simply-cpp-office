#include "csv_sheet.h"

#include <stdexcept>

namespace sc::office::csv {
    const cell &row::at(const std::size_t column) const {
        if (const auto found = cells_.find(column); found != cells_.end()) return found->second;
        throw std::out_of_range("csv: no cell " + address{index_, column}.name());
    }

    cell &row::at(const std::size_t column) { return const_cast<cell &>(std::as_const(*this).at(column)); }

    cell &row::operator[](const std::size_t column) {
        return cells_.try_emplace(column, cell{.row = index_, .column = column}).first->second;
    }

    sheet::sheet(const excel::sheet &from) : name_(from.name()) {
        for (const auto &each: from)
            for (const auto &item: each)
                if (!std::holds_alternative<std::monostate>(item.value)) (*this)[{item.row, item.column}] = item.value;
    }

    bool sheet::contains(const address where) const {
        const auto found = rows_.find(where.row);
        return found != rows_.end() && found->second.contains(where.column);
    }

    const cell &sheet::at(const address where) const {
        if (const auto found = rows_.find(where.row); found != rows_.end() && found->second.contains(where.column))
            return found->second.at(where.column);
        throw std::out_of_range("csv: no cell " + where.name() + (name_.empty() ? "" : " in \"" + name_ + "\""));
    }

    cell &sheet::at(const address where) { return const_cast<cell &>(std::as_const(*this).at(where)); }

    cell &sheet::operator[](const address where) { return rows_.try_emplace(where.row, where.row).first->second[where.column]; }

    std::size_t sheet::erase(const address where) {
        const auto found = rows_.find(where.row);
        if (found == rows_.end()) return 0;
        const auto removed = found->second.erase(where.column);
        if (found->second.empty()) rows_.erase(found);
        return removed;
    }

    excel::sheet sheet::to_excel(std::string name) const {
        excel::sheet out{!name.empty() ? std::move(name) : !name_.empty() ? name_ : std::string{"Sheet1"}};
        for (const auto &each: *this)
            for (const auto &item: each) out[{item.row, item.column}] = item.value;
        return out;
    }
}
