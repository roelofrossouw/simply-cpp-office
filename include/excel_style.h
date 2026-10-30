#pragma once

#include <color.h>
#include <distance.h>

#include <optional>
#include <string>
#include <string_view>

// Excel (sc-office): how a cell looks: its font's family, size, colour, weight, slant,
// strikethrough and underline, its background, its borders, its alignment and its number format.
namespace sc::office::excel {
    // A font's underline, as Excel has them: the accounting ones run the cell's width.
    enum class underline { none, single, doubled, single_accounting, double_accounting };

    // A border's line, as Excel has them.
    enum class line {
        none, thin, medium, thick, hair, dotted, dashed, doubled, dash_dot, dash_dot_dot, slant_dash_dot,
        medium_dashed, medium_dash_dot, medium_dash_dot_dot
    };

    // How text sits across a cell. general: text left, numbers right; fill repeats the text
    // across the cell; center_across centres it across this and the empty cells to its right.
    enum class horizontal { general, left, center, right, fill, justify, center_across, distributed };

    // How text sits up and down a cell; bottom is Excel's default.
    enum class vertical { bottom, top, center, justify, distributed };

    // Excel's names for them: "single", "double", "singleAccounting"; "thin", "mediumDashDot";
    // "center", "centerContinuous"; "top". none is "none".
    [[nodiscard]] std::string_view name(underline kind);
    [[nodiscard]] std::string_view name(line kind);
    [[nodiscard]] std::string_view name(horizontal kind);
    [[nodiscard]] std::string_view name(vertical kind);
    // The underline, line or alignment Excel names so; std::invalid_argument for a name it doesn't
    // have.
    [[nodiscard]] underline underline_named(std::string_view name);
    [[nodiscard]] line line_named(std::string_view name);
    [[nodiscard]] horizontal horizontal_named(std::string_view name);
    [[nodiscard]] vertical vertical_named(std::string_view name);

    // A cell's font. What isn't set is the workbook's default (Calibri, 11 points, automatic colour).
    struct font {
        // Its family: "Arial", "Courier New"; empty: the workbook's (Calibri).
        std::string name;
        // In points: 14, or sc::dist::point(14).
        std::optional<sc::distance> size;
        // sc::color{"#c00000"}, sc::color::Red ...; alpha is left out.
        std::optional<sc::color> color;
        bool bold = false;
        bool italic = false;
        bool strike = false; // struck through
        excel::underline underline = underline::none;

        bool operator==(const font &) const = default;
    };

    // One side's border: its line, and its colour (automatic, black, when it isn't set).
    struct edge {
        excel::line line = line::none;
        std::optional<sc::color> color;

        bool operator==(const edge &) const = default;
    };

    // A cell's borders, side by side, and its diagonals: up from bottom left to top right, down from
    // top left to bottom right. Excel draws both diagonals alike, so when both are set they must be
    // the same edge (writing throws otherwise).
    struct borders {
        edge left, right, top, bottom;
        edge diagonal_up, diagonal_down;

        // The same edge on all four sides (no diagonals): excel::borders::all({excel::line::thin}).
        [[nodiscard]] static borders all(const edge &each) { return {each, each, each, each}; }

        bool operator==(const borders &) const = default;
    };

    // Where a cell's text sits, whether it wraps onto more lines, how far it's indented (in steps of
    // about three characters' width, up to 250), and how it's turned.
    struct alignment {
        excel::horizontal horizontal = horizontal::general;
        excel::vertical vertical = vertical::bottom;
        bool wrap = false;
        unsigned indent = 0;
        // Degrees counter-clockwise, from -90 (reading downwards) to 90 (upwards), as Excel's
        // Orientation shows them.
        int rotation = 0;
        // Letters stacked one under another (Excel's Vertical Text); rotation then counts for nothing.
        bool stacked = false;

        bool operator==(const alignment &) const = default;
    };

    // How a cell looks; a default style is Excel's plain one.
    //
    //     sheet["A1"].style.font.bold = true;
    //     sheet["A1"].style.background = sc::color{"#ffff00"};
    //     sheet["A1"].style.borders.bottom = {excel::line::doubled, sc::color{"#1f4e79"}};
    //     sheet["A1"].style.font.name = "Arial";
    //     sheet["B1"].style = {.font = {.size = 14, .color = sc::color{"#1f4e79"}, .bold = true},
    //                          .background = sc::color{"#ddebf7"},
    //                          .borders = excel::borders::all({excel::line::thin}),
    //                          .alignment = {.horizontal = excel::horizontal::center, .wrap = true}};
    //     sheet["C1"].style.number_format = "#,##0.00";
    struct style {
        excel::font font;
        // A solid fill; none when it isn't set.
        std::optional<sc::color> background;
        excel::borders borders;
        excel::alignment alignment;
        // Excel's format code: "0.00", "#,##0", "0%", "yyyy-mm-dd", "h:mm", "@" (text),
        // "R #,##0.00;[Red]-R #,##0.00". Empty: General. A date is a number, days since 1899-12-30,
        // shown as one by a date format.
        std::string number_format;

        bool operator==(const style &) const = default;
        // Whether it's Excel's plain style.
        [[nodiscard]] bool empty() const { return *this == style{}; }
    };
}
