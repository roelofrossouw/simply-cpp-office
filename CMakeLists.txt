cmake_minimum_required(VERSION 3.22)
set(SC_MODULE office)
project(simply-cpp-${SC_MODULE} HOMEPAGE_URL "https://github.com/roelofrossouw/simply-cpp-office" LANGUAGES CXX)
include(cmake/sc_bootstrap.cmake)

option(SC_OFFICE_BUILD_TESTS "Build simply-cpp-office tests" ${PROJECT_IS_TOP_LEVEL})

get_sc_version()
message(STATUS "SimplyCPP ${SC_MODULE} Version: ${SC_VERSION}")

find_or_install_package(sc-core simply-cpp-dev simply-cpp 1.20.0)
# Office files are ZIP archives of XML: sc-compression opens them, sc-web reads the XML.
find_or_install_package(sc-compression simply-cpp-compression-dev simply-cpp-compression 0.0.0)
find_or_install_package(sc-web simply-cpp-web-dev simply-cpp-web 1.4.7)

set(SOURCE_OBJECTS include/office.h)
set(SOURCE_LIBRARIES)
set(SOURCE_LINK_LIBRARIES)

add_sc_object(office
        LINK_LIBRARIES sc::sc-core sc::sc-compression sc::sc-web
        PUBLIC_LINK_LIBRARIES sc::sc-core sc::sc-compression sc::sc-web)

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

add_sc_example(sc-office-demo)

package_sc_module(DEV_DEPENDS "simply-cpp-dev (>= 1.20.0~)" "simply-cpp-compression-dev (>= 0.0.0~)" "simply-cpp-web-dev (>= 1.4.7~)")
