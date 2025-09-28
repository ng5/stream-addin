#include <exception>
#include <windows.h>
#include <thread>
#include <random>
#include "xlcall.h"
#include "framewrk.h"

#include <chrono>

extern "C" __declspec(dllexport) double WINAPI CalcCircum(const double x) { return x * 6.283185308; }
extern "C" __declspec(dllexport) void __stdcall RandTick(LPXLOPER12 pAsyncHandle) {
    XLOPER12 asyncHandle = *pAsyncHandle;
    std::thread([asyncHandle]() mutable {
        using namespace std::chrono_literals;
        std::mt19937_64 rng{(GetTickCount64())};
        std::uniform_real_distribution dist(0.0, 1.0);
        XLOPER12 result{};
        result.xltype = xltypeNum;
        result.val.num = dist(rng);
        Excel12(xlAsyncReturn, nullptr, 2, &asyncHandle, &result);
        std::this_thread::sleep_for(1s);
    }).detach();
}
extern "C" __declspec(dllexport) int __stdcall xlAutoOpen(void) {
    try {
        static XLOPER12 xDLL{};
        Excel12f(xlGetName, &xDLL, 0);
        int rc = Excel12f(xlfRegister, nullptr, 4, static_cast<LPXLOPER12>(&xDLL),
                          static_cast<LPXLOPER12>(TempStr12(L"CalcCircum")), static_cast<LPXLOPER12>(TempStr12(L"BB")),
                          static_cast<LPXLOPER12>(TempStr12(L"CalcCircum")));
        rc = Excel12f(xlfRegister, nullptr, 4, static_cast<LPXLOPER12>(&xDLL),
                      static_cast<LPXLOPER12>(TempStr12(L"RandTick")), static_cast<LPXLOPER12>(TempStr12(L">X")),
                      static_cast<LPXLOPER12>(TempStr12(L"RandTick")));
        Excel12f(xlFree, nullptr, 1, static_cast<LPXLOPER12>(&xDLL));
        return 1;
    } catch (const std::exception &e) {
        return 0;
    }
}

extern "C" __declspec(dllexport) int __stdcall xlAutoClose(void) { return 1; }
