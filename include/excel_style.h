#pragma once

#include <color.h>
#include <distance.h>

#include <optional>

// Excel (sc-office): how a cell looks: its font's size, colour and weight, and its background.
namespace sc::office::excel {
    // A cell's font. What isn't set is the workbook's default (Calibri, 11 points, automatic colour).
    struct font {
        // In points: 14, or sc::dist::point(14).
        std::optional<sc::distance> size;
        // sc::color{"#c00000"}, sc::color::Red ...; alpha is left out.
        std::optional<sc::color> color;
        bool bold = false;

        bool operator==(const font &) const = default;
    };

    // How a cell looks; a default style is Excel's plain one.
    //
    //     sheet["A1"].style.font.bold = true;
    //     sheet["A1"].style.background = sc::color{"#ffff00"};
    //     sheet["B1"].style = {.font = {.size = 14, .color = sc::color{"#1f4e79"}, .bold = true},
    //                          .background = sc::color{"#ddebf7"}};
    struct style {
        excel::font font;
        // A solid fill; none when it isn't set.
        std::optional<sc::color> background;

        bool operator==(const style &) const = default;
        // Whether it's Excel's plain style.
        [[nodiscard]] bool empty() const { return *this == style{}; }
    };
}
