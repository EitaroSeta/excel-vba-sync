# Changelog
All notable changes to the "excel-vba-sync" extension are documented here.

This file follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/)
and uses [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]
### Planned
- Improve error messages around VBA import/export.
- Add docs: troubleshooting for PowerShell session/language server.

## [0.0.27] - 2025-09-13
### ### Added
- Excel Macro Execution.  
Added ability to execute VBA macros by fully qualified name or by specifying module/procedure names.  
- VBA Code Search (vba_search_code)  
New tool to search across all open Excel workbooks and their VBA modules.  
Supports both plain text and regex search (useRegex).

### ### Improved
- VS Code Extension Integration  
Implemented JSON-RPC communication between the extension and the server.  
Allows temporary display of fetched VBA code with automatic navigation to the matched line.  
If exported .bas, .cls, or .frm files already exist, they are prioritized for opening instead of fetching directly from Excel.  
- Error Handling & Stability  
Added clear JSON-formatted error messages when Excel is not running, or when workbooks/modules are not found.  
Introduced execution timeout (20s) and buffer size limits (2 MB) for more robust process control.
Improved JSON parsing safety for server responses.

## [0.0.26] - 2025-09-03
### ### Changed
- Fine-tuned message text.

## [0.0.25] - 2025-09-02
### ### Changed
- Added emoji to log prefix.
- Updated activity bar icon.

## [0.0.24] - 2025-09-02
### ### Fixed
- Fix the character encoding to UTF-8 when exporting workbook modules by export().

## [0.0.23] - 2025-09-01
### ### Fixed
- Fixed export log output error

## [0.0.22] - 2025-08-31
### ### Fixed
- Fixed issue where INFO logs were output even on errors.

### ### Changed
- Fine-tuned message text.
- Minor changelog output"

### ### Added
- Enabled exporting files via right-click.

## [0.0.21] - 2025-08-30
### ### Fixed
- Fixed import issue from statusbar.

## [0.0.20] - 2025-08-30
### ### Added
- Enabled importing files via right-click.

### ### Fixed
- Aligned ATTRIBUTE output of files with VBE output.
- Fixed import issue with cls files.

## [0.0.19] - 2025-08-29
### ### Added
- Monitor the folder for changes and refresh the directory and file information.

## [0.0.18] - 2025-08-28
### ### Fixed
- Fix the character encoding to UTF-8 when exporting workbook modules.

## [0.0.17] - 2025-08-28
### ### Changed
- Fine-tuned message text.
- Minor README correction

## [0.0.16] - 2025-08-28
### ### Changed
- Fine-tuned message text.
- Added export file extension check（\*.xlsm/\*.xlsb only）
- Added import file extension check（\*.bas/\*.cls/\*.frm only）

## [0.0.15] - 2025-08-26
### ### Changed
- Fine-tuned message text.
- SUnified message logging to **VS Code Output Channel** (all logs/errors are now centralized in the Output panel)
- Added timestamps to messages.

### ### Fixed
- Fixed a bug where a file dialog appeared when no folder was specified during import.

## [0.0.10] - 2025-08-23
### Added
- **Initial public release on VS Code Marketplace.**
- Commands to **Export** / **Import** VBA modules against the *opened* Excel project.
- Localization: **en** / **ja**.

### Notes
- **Limitation**: This tool **replaces existing** modules/classes/forms only; **adding new items is not supported**.  
  To create a new item, add & save a blank module/class/form in VBE, then export it.
- **Caution**: Do **not edit attribute lines** in exported `.frm` files  
  (`VERSION`, `Begin … End`, `Object = …`, `Attribute VB_*`).

