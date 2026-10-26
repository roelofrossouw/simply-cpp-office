#include "excel_workbook.h"
#include "excel_sheet_names.h"

#include <stdexcept>
#include <utility>

namespace sc::office::excel {
    sheet &workbook::add(std::string name) {
        detail::check_name(name);
        if (contains(name)) throw std::invalid_argument("excel: there's a sheet named \"" + name + "\" already");
        return sheets_.emplace_back(std::move(name));
    }

    const sheet *workbook::find(const std::string_view name) const {
        for (const auto &each: sheets_)
            if (detail::same_name(each.name(), name)) return &each;
        return nullptr;
    }

    bool workbook::contains(const std::string_view name) const { return find(name) != nullptr; }

    const sheet &workbook::at(const std::string_view name) const {
        if (const auto *found = find(name)) return *found;
        throw std::out_of_range("excel: no sheet named \"" + std::string{name} + "\"");
    }

    sheet &workbook::at(const std::string_view name) { return const_cast<sheet &>(std::as_const(*this).at(name)); }

    sheet &workbook::operator[](const std::string_view name) {
        if (const auto *found = find(name)) return const_cast<sheet &>(*found);
        return add(std::string{name});
    }
}
