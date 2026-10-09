# PJ Shoes Slip Printing & Phone Records

A standalone, offline Windows desktop app for PJ Shoes. It records customer
names and phone numbers, prints slips, and searches and manages saved records.
There is no draw or winner-selection feature.

## Windows 7 RTM compatibility

- 32-bit WinForms executable targeting .NET Framework 3.5 / CLR 2.0.
- Windows 7 includes the .NET Framework 3.5.1 feature; enable it if it has
  been turned off. The app does not require Windows 7 SP1 or later updates.
- Uses built-in WinForms, XML, SpreadsheetML, GDI printing, and Windows printer
  enumeration. It does not require Python, Node.js, Electron, SQLite binaries,
  Office automation, a network connection, or a third-party runtime.
- The EXE is x86 and embeds the PJ Shoes logo. Windows supplies .NET and the
  printer driver.

## Features

- Enter slip number, customer name, phone number, and entry date.
- Save slips locally, edit an existing slip, search, view, and delete records.
- Print new or saved receipts to a selected installed printer.
- Dynamically list installed printers, save the selected printer, and send a
  test slip without creating a customer record.
- Open the local data folder for backup.
- Import records from the previous app's `.xlsx` workbook. Import is read-only:
  existing slip numbers are skipped, and the selected workbook is not changed.

## Build on Windows 7 x86

1. Copy this `windows-app` folder to the Windows 7 computer.
2. Enable **Microsoft .NET Framework 3.5.1** in Windows Features if it is off.
3. Open Command Prompt in this folder and run:

   ```bat
   build.bat
   ```

4. The single 32-bit executable will be at `bin\PJShoesSlipRecords.exe`.

The build uses the .NET Framework 3.5 C# compiler already installed with the
framework. No NuGet restore, package download, or internet access is needed.

## First launch and printing

- Customer records are stored under the current Windows user's
  `%APPDATA%\PJ Shoes\Slip Records\pj_shoes_slips.xml`.
- The data is Excel 2003 SpreadsheetML XML and can be opened in Excel. A `.bak`
  copy of the previous workbook is kept when records are saved.
- Printer selection is saved in `settings.xml` in the same data folder.
- Install the Windows driver for the receipt printer, select the correct
  printer here, and choose 80mm paper in that printer's Windows preferences.
  Run **Test print** before issuing customer tokens.
- Use **Import old workbook** to bring forward records from
  `pj_shoes_lucky_draw.xlsx`. The original workbook remains untouched.

## Compatibility test note

The source and build target use APIs available in .NET Framework 3.5 and the
Windows 7 printing stack. A final compatibility claim still requires building
and smoke-testing the generated EXE on a physical Windows 7 RTM x86 computer
with the intended printer driver. The Replit build environment cannot run that
Windows configuration or verify thermal-printer output.
