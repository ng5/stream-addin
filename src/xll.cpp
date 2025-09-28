#include <exception>
#include <windows.h>
#include "xlcall.h"
#include "framewrk.h"

extern "C" __declspec(dllexport) double WINAPI CalcCircum(const double x) { return x * 6.283185308; }
extern "C" __declspec(dllexport) int __stdcall xlAutoOpen(void) {
    try {
        static XLOPER12 xDLL{};
        Excel12f(xlGetName, &xDLL, 0);
        int rc = Excel12f(xlfRegister, nullptr, 4, static_cast<LPXLOPER12>(&xDLL),
                          static_cast<LPXLOPER12>(TempStr12(L"CalcCircum")), static_cast<LPXLOPER12>(TempStr12(L"BB")),
                          static_cast<LPXLOPER12>(TempStr12(L"CalcCircum")));
        Excel12f(xlFree, nullptr, 1, static_cast<LPXLOPER12>(&xDLL));
        return 1;
    } catch (const std::exception &e) {
        return 0;
    }
}

extern "C" __declspec(dllexport) int __stdcall xlAutoClose(void) { return 1; }
