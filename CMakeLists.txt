cmake_minimum_required(VERSION 3.20)
project(SimpleXllDiag LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 23)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# --- Adjust to your SDK path ---
set(EXCEL_SDK_ROOT "E:/ExcelSDK/Excel2013XLLSDK")

set(EXCEL_SDK_INCLUDE_DIRS
        "${EXCEL_SDK_ROOT}/INCLUDE"
        "${EXCEL_SDK_ROOT}/SAMPLES/FRAMEWRK"
)
set(EXCEL_SDK_LIBDIR "${EXCEL_SDK_ROOT}/LIB/x64") # use LIB/Win32 for 32-bit Excel

find_library(XLCALL32_LIB NAMES XLCALL32 xlcall32 PATHS "${EXCEL_SDK_LIBDIR}" REQUIRED)
find_library(FRAMEWRK_LIB  NAMES FRAMEWRK  frmwrk32  PATHS "${EXCEL_SDK_LIBDIR}" REQUIRED)

add_library(simplexll SHARED src/simplexll.cpp)
target_include_directories(simplexll PRIVATE ${EXCEL_SDK_INCLUDE_DIRS})
target_link_libraries(simplexll PRIVATE ${XLCALL32_LIB} ${FRAMEWRK_LIB})

# Diagnostics + static CRT
target_compile_definitions(simplexll PRIVATE UNICODE _UNICODE WIN32_LEAN_AND_MEAN NOMINMAX)
if(MSVC)
    target_compile_options(simplexll PRIVATE /W4 /permissive- /EHsc)
    # Static CRT so machines don’t need the VC++ Redistributable
    set_property(TARGET simplexll PROPERTY MSVC_RUNTIME_LIBRARY "MultiThreadedDLL")
endif()

# Export required symbols without a .def file
# (SetExcel12EntryPt comes from FRAMEWRK.lib; the others are in our source)
target_link_options(simplexll PRIVATE
        /EXPORT:SetExcel12EntryPt
        /EXPORT:xlAutoOpen
        /EXPORT:xlAutoClose
        /EXPORT:CalcCircum
)

# Name it .xll
set_target_properties(simplexll PROPERTIES OUTPUT_NAME "simplexll" SUFFIX ".xll")

# Convenience copy
set(DEPLOY_DIR "${CMAKE_BINARY_DIR}/deploy/$<CONFIG>")
add_custom_command(TARGET simplexll POST_BUILD
        COMMAND ${CMAKE_COMMAND} -E make_directory "${DEPLOY_DIR}"
        COMMAND ${CMAKE_COMMAND} -E copy $<TARGET_FILE:simplexll> "${DEPLOY_DIR}/simplexll.xll"
)
