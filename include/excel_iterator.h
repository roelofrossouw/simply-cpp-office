#pragma once

#include <cstddef>
#include <iterator>
#include <type_traits>
#include <utility>

// Excel (sc-office): rows and cells are kept in maps by their place; their iterators give the rows and cells.
namespace sc::office::excel {
    namespace detail {
        // A map's iterator that gives its values, not its pairs: rows and cells in order.
        template<typename map_iterator>
        class values_iterator {
        public:
            using iterator_category = std::bidirectional_iterator_tag;
            using value_type = std::remove_cvref_t<decltype(std::declval<map_iterator>()->second)>;
            using difference_type = std::ptrdiff_t;
            using reference = decltype((std::declval<map_iterator>()->second));
            using pointer = std::add_pointer_t<reference>;

            values_iterator() = default;
            explicit values_iterator(map_iterator at) : at_(at) {}

            reference operator*() const { return at_->second; }
            pointer operator->() const { return &at_->second; }
            values_iterator &operator++() {
                ++at_;
                return *this;
            }
            values_iterator operator++(int) { return values_iterator{at_++}; }
            values_iterator &operator--() {
                --at_;
                return *this;
            }
            values_iterator operator--(int) { return values_iterator{at_--}; }
            bool operator==(const values_iterator &) const = default;

        private:
            map_iterator at_{};
        };
    }
}
