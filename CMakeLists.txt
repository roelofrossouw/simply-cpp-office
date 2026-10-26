cmake_minimum_required(VERSION 3.22)
set(SC_MODULE office)
project(simply-cpp-${SC_MODULE} HOMEPAGE_URL "https://github.com/roelofrossouw/simply-cpp-office" LANGUAGES CXX)
include(cmake/sc_bootstrap.cmake)

option(SC_OFFICE_BUILD_TESTS "Build simply-cpp-office tests" ${PROJECT_IS_TOP_LEVEL})

get_sc_version()
message(STATUS "SimplyCPP ${SC_MODULE} Version: ${SC_VERSION}")

find_or_install_package(sc-core simply-cpp-dev simply-cpp 1.20.0)
# Office files are ZIP archives of XML: sc-archive opens and writes them, sc-web reads the XML.
find_or_install_package(sc-archive simply-cpp-archive-dev simply-cpp-archive 1.0.0)
find_or_install_package(sc-web simply-cpp-web-dev simply-cpp-web 1.4.7)

set(SOURCE_OBJECTS include/office.h include/excel_cell.h include/excel_iterator.h include/excel_style.h)
set(SOURCE_LIBRARIES)
set(SOURCE_LINK_LIBRARIES)

add_sc_object(office
        LINK_LIBRARIES sc::sc-core sc::sc-archive sc::sc-web
        PUBLIC_LINK_LIBRARIES sc::sc-core sc::sc-archive sc::sc-web)
# Excel: its cells, sheets and workbooks, then reading and writing them.
add_sc_object(excel_address)
add_sc_object(excel_range)
add_sc_object(excel_style LINK_LIBRARIES sc::sc-core)
add_sc_object(excel_row)
add_sc_object(excel_sheet)
add_sc_object(excel_workbook SOURCES src/excel_sheet_names.h LINK_LIBRARIES sc::sc-core)
add_sc_object(excel SOURCES src/excel_sheet_names.h LINK_LIBRARIES sc::sc-core sc::sc-archive sc::sc-web)

add_sc_libraries(DESCRIPTION "simply-cpp Office documents")
target_compile_features(sc-office PUBLIC cxx_std_20)
target_compile_features(sc-office-shared PUBLIC cxx_std_20)

install_sc_module(PATH_VARS CMAKE_INSTALL_INCLUDEDIR)
install(FILES cmake/SimplyCppFunctions.cmake cmake/sc_bootstrap.cmake
        DESTINATION ${CMAKE_INSTALL_LIBDIR}/cmake/sc-office
        COMPONENT development)

if (SC_OFFICE_BUILD_TESTS)
    enable_testing()
    add_subdirectory(tests sc-office-test)
endif ()

add_sc_example(sc-office-word-demo)
add_sc_example(sc-office-excel-demo)

package_sc_module(DEV_DEPENDS "simply-cpp-dev (>= 1.20.0~)" "simply-cpp-archive-dev (>= 1.0.0~)" "simply-cpp-web-dev (>= 1.4.7~)")
