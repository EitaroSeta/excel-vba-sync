# Changelog
All notable changes to the "excel-vba-sync" extension are documented here.

This file follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/)
and uses [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]
### Planned
- Improve error messages around VBA import/export.
- Add docs: troubleshooting for PowerShell session/language server.

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

