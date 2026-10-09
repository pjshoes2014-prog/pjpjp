@echo off
setlocal
cd /d "%~dp0"

set "CSC=%WINDIR%\Microsoft.NET\Framework\v3.5\csc.exe"
if not exist "%CSC%" (
  echo .NET Framework 3.5 compiler was not found.
  echo On Windows 7, enable .NET Framework 3.5.1 in Windows Features,
  echo or install the .NET Framework 3.5 development tools.
  pause
  exit /b 1
)

set "WINDOWSBASE=%WINDIR%\Microsoft.NET\Framework\v3.0\Windows Presentation Foundation\WindowsBase.dll"
if not exist "%WINDOWSBASE%" set "WINDOWSBASE=%ProgramFiles%\Reference Assemblies\Microsoft\Framework\v3.0\WindowsBase.dll"
if not exist "%WINDOWSBASE%" set "WINDOWSBASE=%ProgramFiles%\Reference Assemblies\Microsoft\Framework\v3.5\WindowsBase.dll"
if not exist "%WINDOWSBASE%" (
  echo WindowsBase.dll was not found. Install the .NET Framework 3.5.1
  echo development components and run this build again.
  pause
  exit /b 1
)

if not exist "bin" mkdir "bin"

"%CSC%" /nologo /target:winexe /platform:x86 /optimize+ /debug- /codepage:65001 ^
  /win32manifest:app.manifest ^
  /out:"bin\PJShoesSlipRecords.exe" ^
  /reference:"%WINDIR%\Microsoft.NET\Framework\v3.5\System.dll" ^
  /reference:"%WINDIR%\Microsoft.NET\Framework\v3.5\System.Core.dll" ^
  /reference:"%WINDIR%\Microsoft.NET\Framework\v3.5\System.Drawing.dll" ^
  /reference:"%WINDIR%\Microsoft.NET\Framework\v3.5\System.Windows.Forms.dll" ^
  /reference:"%WINDIR%\Microsoft.NET\Framework\v3.5\System.Xml.dll" ^
  /reference:"%WINDIR%\Microsoft.NET\Framework\v3.5\System.Xml.Linq.dll" ^
  /reference:"%WINDOWSBASE%" ^
  /resource:"assets\pj-shoes-logo.jpg,PJShoesSlipRecords.assets.pj-shoes-logo.jpg" ^
  src\SlipRecord.cs ^
  src\SpreadsheetStore.cs ^
  src\SettingsStore.cs ^
  src\LegacyWorkbookImporter.cs ^
  src\PrinterService.cs ^
  src\MainForm.cs ^
  src\Program.cs

if errorlevel 1 (
  echo.
  echo Build failed. Review the compiler errors above.
  pause
  exit /b 1
)

echo.
echo Build complete: bin\PJShoesSlipRecords.exe
echo This is a single 32-bit executable with the PJ Shoes logo embedded.
echo Customer data and printer settings are saved under the current user's
echo Application Data folder, not beside the executable.
pause
exit /b 0
