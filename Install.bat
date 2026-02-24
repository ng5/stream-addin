@echo off
rem Register an Excel add-in using the OPEN99 trick
rem Usage: Install.bat path\to\addin.xll

setlocal EnableExtensions EnableDelayedExpansion

if "%~1"=="" (
  echo Usage: %~nx0 path\to\addin.xll
  exit /b 1
)

set "ADDIN=%~f1"
if not exist "%ADDIN%" (
  echo Add-in not found: %ADDIN%
  exit /b 1
)

echo Registering add-in: %ADDIN%

rem Known Excel versions to try (per-user keys)
set "VERSIONS=16.0 15.0 14.0 12.0 11.0 10.0 9.0"

for %%V in (%VERSIONS%) do (
  reg query "HKCU\Software\Microsoft\Office\%%V\Excel\Options" >nul 2>&1
  if !ERRORLEVEL! EQU 0 (
    echo -> Adding OPEN99 for Excel %%V
    set "KEY=HKCU\Software\Microsoft\Office\%%V\Excel\Options"
    rem Value should include quotes so Excel treats paths with spaces correctly
    set "DATA=\"%ADDIN%\""
    reg add "!KEY!" /v "OPEN99" /t REG_SZ /d "!DATA!" /f >nul 2>&1
    if !ERRORLEVEL! EQU 0 (
      echo    OK
    ) else (
      echo    Failed to set registry value for %%V
    )
  ) else (
    rem Key not present, skip
  )
)

echo Done. Restart Excel to load the add-in (if Excel was running, restart it).

endlocal
