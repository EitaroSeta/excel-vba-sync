# Changelog
All notable changes to the "excel-vba-sync" extension are documented here.

This file follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/)
and uses [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

**Note on development process**: Since v0.0.28, implementation has been done via AI-assisted development ("vibe coding") with Claude (Anthropic), reviewed by a human at each step. See [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md) for details.

## [Unreleased]
### Planned
- Improve error messages around VBA import/export.
- Add docs: troubleshooting for PowerShell session/language server.

## [0.0.42] - 2026-07-26
### ### Added
- Disclosed the AI-assisted ("vibe coding" with Claude/Anthropic) development process since v0.0.28 in `CHANGELOG.md`, `README.md`, and `docs/DEVELOPMENT.md`, for transparency. No functional code changes in this release.

## [0.0.41] - 2026-07-26
### ### Added
- Registered this extension's MCP server with VS Code's native MCP server discovery via `contributes.mcpServerDefinitionProviders` (a stable, non-proposed VS Code API as of the installed `@types/vscode` 1.104.0). This lets clients built into VS Code itself, such as Copilot Chat's agent mode, find and start the server automatically -- in addition to, not instead of, the existing manual `.mcp.json`/`claude_desktop_config.json` route via "Excel VBA: Print MCP Server Config (for AI)", which external apps like Claude Code/Desktop still need since they cannot see this contribution point. Reuses the exact same `process.execPath`/`dist-server/server.js`/`MCP_SCRIPTS_DIR`/`ELECTRON_RUN_AS_NODE` values already proven to work via that command. The registration is wrapped in a feature-detection guard (`typeof vscode.lm?.registerMcpServerDefinitionProvider === "function"`) so it's silently skipped on VS Code versions that predate this API, instead of throwing during `activate()` and disabling every other command (the same class of bug fixed for `fs.watch()` in v0.0.29).

### ### Known limitations
- This adds a third independent way `dist-server/server.js` can be launched (alongside the extension's own internal use and any externally-configured Claude Code/Desktop `.mcp.json`). Each MCP client gets its own dedicated stdio server process by design -- this is not a new problem, but it does raise the odds of the pre-existing multi-Excel-process scenario (see `launchedExcelPid`, v0.0.33) occurring in practice if multiple clients happen to be active around the same time.
- Not yet verified end-to-end against an actual VS Code MCP client (e.g. Copilot Chat agent mode) discovering and using this provider -- verified only that it compiles and reuses already-proven config values. Manual verification in an environment with such a client is recommended before relying on this path.

## [0.0.40] - 2026-07-26
### ### Added
- `excel_update_module_code` now uses optimistic concurrency control: the `confirmToken` returned by a `dryRun:true` call is bound to the module's code as it was at that moment, and the tool re-reads the module immediately before writing to recompute and compare the token. If the code has changed in the meantime (e.g. a different MCP client -- Claude Code, Claude Desktop, or VS Code's own Copilot Chat if it's ever wired up via `mcpServerDefinitionProviders` -- wrote to the same module first), the write is now rejected with `ERR_MODULE_CHANGED_SINCE_DRYRUN` instead of silently overwriting that change. Verified against a live workbook: two independent dry-run/confirm sequences against the same module, where the second write completed first -- the first (now-stale) confirmToken was correctly rejected without touching the second write's content.

### ### Changed
- `computeConfirmToken`'s formula changed from `hash(workbook, module, newCode)` to `hash(module, currentCode, newCode)` -- it no longer includes the workbook identifier (unnecessary for uniqueness here) but now includes the current code snapshot, which is what makes the concurrency check above possible.

## [0.0.39] - 2026-07-26
### ### Fixed
- The v0.0.38 CRLF-to-LF normalization did not fix the broken feature tables either (confirmed by re-checking the live page at v0.0.38 -- identical symptom). Since headings and bullet lists render correctly on the Marketplace page but pipe-delimited tables consistently collapse into a single plain paragraph regardless of anchor placement or line-ending style, the most likely explanation is that the Marketplace README renderer does not support GFM pipe tables at all. Replaced both feature tables ("主な機能" / "Features") with bullet lists in the same style as the bullet list already used above them (which does render correctly), instead of continuing to chase table-syntax fixes.

## [0.0.38] - 2026-07-26
### ### Fixed
- The v0.0.36 anchor-inlining fix did not actually resolve the broken feature tables on the Marketplace page (confirmed by re-checking the live page at v0.0.36). The real cause was that `README.md` used CRLF line endings throughout, leaving a stray trailing `\r` on every line -- including the table's delimiter row (`|---|---|`). Headings and bullet lists tolerate a trailing `\r` invisibly, but the delimiter-row check used by the Marketplace renderer's table parser apparently does not, so the whole table fell back to being rendered as one plain paragraph of pipe-delimited text. Normalized `README.md` to LF line endings, which should let the tables parse correctly.

## [0.0.37] - 2026-07-26
### ### Fixed
- Fixed a Mermaid parse error in `docs/DEVELOPMENT.md`'s architecture diagram: the `PS` node was written as `PS ["PowerShell Scripts(.ps1)"]` (a stray space between the node ID and its label brackets), which the Mermaid grammar does not allow -- GitHub showed "Unable to render rich display / Parse error on line 4" instead of the diagram. This typo predates the recent README split and was simply never noticed until now. Fixed to `PS["PowerShell Scripts(.ps1)"]`, matching the other three nodes.

## [0.0.36] - 2026-07-26
### ### Fixed
- v0.0.35's table-of-contents anchors (`<a id="...">` placed alone on the line right before each heading) broke table rendering on the Marketplace page: a standalone HTML tag like that can be parsed as an "HTML block" that continues, verbatim, until the next blank line -- swallowing the following heading and, apparently on the Marketplace renderer, everything after it, so the pipe-delimited feature tables further down showed up as raw text instead of `<table>`s. Fixed by moving each anchor inline into the same line as its heading (`## <a id="x"></a>Heading text`), which is treated as ordinary inline HTML inside a single heading line rather than a standalone HTML block.

## [0.0.35] - 2026-07-26
### ### Changed
- Restructured `README.md`: it previously mixed manual-user, AI-client, and developer content in one long page. Moved the AI-client usage section into `docs/AI_USAGE.md` and the developer/build section into `docs/DEVELOPMENT.md`, with a short summary and link left in place of each. Added an in-page table of contents with anchors so readers can jump directly to the section relevant to them.
- Strengthened the Excel backup warning: moved it to the top of the Important section as a GitHub-flavored-markdown alert (renders as a colored warning box on GitHub), and added an explicit note that the risk of unintended overwrite/execution is higher than normal when an AI agent is driving the tools autonomously via MCP (no per-step human confirmation).

### ### Added
- Added a Disclaimer section covering damage from use of the extension, including actions taken by an AI agent via the MCP server.

## [0.0.34] - 2026-07-25
### ### Fixed
- v0.0.33's published VSIX unintentionally bundled `.mcp.json` (a machine-specific MCP config containing the developer's local Windows username and folder paths), plus internal working files `BRUSHUP_NOTES.md` and `brushup_changes.diff`. Root cause: `.gitignore` only keeps files out of git, not out of the packaged VSIX -- that requires `.vscodeignore`, which these files were missing from. Added them (along with `.claude/` and `.github/`) to `.vscodeignore`; no functional code changes in this release.

## [0.0.33] - 2026-07-25
### ### Fixed
- Real regression found in live testing: repeated MCP-driven Excel auto-launches (from `Get-OrStartExcelApplication`) could leave an orphaned, workbook-less Excel process running. Once more than one Excel process exists, `GetActiveObject("Excel.Application")` can resolve to the wrong one -- this broke the ordinary VS Code Export command with `保存済みの Excel ブックが見つかりません` / `DISP_E_BADINDEX` errors that had nothing to do with the export logic itself, since it landed on the empty automation instance instead of the user's real session.

### ### Added
- `Get-OrStartExcelApplication` now returns `LaunchedProcessId`, populated only when it actually had to launch a new Excel instance (null when reusing one already running). This is surfaced as `launchedExcelPid` in the JSON response of every MCP tool that can trigger an auto-launch (`excel_get_module_code`, `vba_search_code`, `excel_update_module_code`, `excel_list_macros`, `excel_run_macro`), so a calling agent/user can identify -- and if needed manually clean up -- a process this tooling caused to exist, without touching an unrelated pre-existing Excel session. This is identification/groundwork only; the existing "auto-launched Excel stays open and visible" behavior is unchanged.

## [0.0.32] - 2026-07-25
### ### Fixed
- `Get-ModulePublicSubs` (used by `excel_list_macros`) was missing two categories of runnable Subs: (1) Subs declared without the explicit `Public` keyword, which VBA treats as Public by default -- the common case, since most authors omit it; (2) procedures with non-ASCII names (e.g. Japanese identifiers) were matched with an ASCII-only regex and silently skipped even when correctly marked `Public`. Both are now detected correctly; `Private`/`Friend` subs remain excluded.

### ### Improved
- Added a description to every MCP tool and to the less-obvious parameters (via zod `.describe()`). Previously none of the 6 tools had any description text, so an AI client connecting without prior context about this project would see only tool names and bare parameter types -- with no indication of `excel_update_module_code`'s required `dryRun` → `confirmToken` flow, that `workbookPath` auto-launches Excel, that `excel_run_macro` can hang on a blocking dialog, or that a successful response only means the call didn't throw (not that the macro did the intended thing). This is the primary "reference" for AI clients now; see the note in the README's AI client section for why a separate reference doc wasn't added on top of it.

### ### Notes from live MCP testing
- VBA compiles the entire project as a unit. If ANY module in the workbook has a compile error (e.g. an undeclared variable under `Option Explicit`, or otherwise malformed code), `excel_run_macro` will fail/hang for every macro in the project, not just the broken one -- Excel shows a blocking "コンパイルエラー" dialog that requires manual dismissal. This is inherent to VBA, not something these tools can detect or work around; if a macro call unexpectedly hangs or times out, check Excel directly for a stuck compile-error dialog before assuming the target macro itself is at fault.

## [0.0.31] - 2026-07-25
### ### Added
- New MCP tool `excel_update_module_code`: lets an AI client (Claude Code/Desktop, etc.) write code into a VBA module. Uses a `dryRun` → `confirmToken` two-step flow (preview the diff and get a token, then re-call with the token to actually apply it), and always takes a timestamped backup to `.excel-vba-sync-backups` next to the workbook before writing. Reuses the same `VBComponents.Import()`-based logic already fixed for issue #3, so Attribute-line handling stays correct; a module of type Document (Sheet/ThisWorkbook) can still lose shortcut-key attributes on write due to the underlying VBA API constraint, and the response says so explicitly.
- `ExcelUtil.ps1`: `Get-OrStartExcelApplication` (launches Excel if it isn't already running) and `Resolve-TargetWorkbook` (opens a workbook by full path if it isn't already open) so tools no longer require Excel/the target workbook to be manually opened first. All MCP tools gained an optional `workbookPath` parameter to use this.
- `ExcelUtil.ps1`: `Test-VbaTrustAccess` surfaces a distinct `ERR_VBOM_TRUST_DISABLED` error (instead of silently returning nothing) when Excel's "Trust access to the VBA project object model" setting is off.
- New command **Excel VBA: Print MCP Server Config (for AI)** generates a ready-to-paste `.mcp.json`/`claude_desktop_config.json` snippet. The MCP server (`dist-server/server.js`) runs as plain Node and can be used by an AI client standalone, without VS Code running.
- `excel_run_macro` gained `timeoutMs` (default 30s); on timeout the wrapping PowerShell process is killed and `ERR_TIMEOUT` is returned (note: this does not un-stick Excel itself if it's blocked on a dialog).
- `vba_search_code` gained `maxResults` (default 50); overly broad queries against a large project now return `truncated`/`totalMatchCount` instead of an unbounded result set.

### ### Fixed
- Fixed extension activation crash when the folder configured via "Set Export Folder" had since been deleted. `fs.watch()` was throwing synchronously (ENOENT) with no error handling, which aborted `activate()` before commands registered later (Export/Import/etc.) ever got wired up — the window would open but none of those commands would work.
- Tool failure responses are now consistently signaled at the MCP protocol level (`isError: true`) instead of only being visible by inspecting the JSON body — this covers both `{error: "ERR_..."}` and general `{ok: false, ...}` payloads, and also recovers the real error detail from a script's stdout when PowerShell exits non-zero (previously collapsed into a generic "ps failed" message).
- `excel_list_macros` list-mode response unified to `{ok, macros, count}` instead of a bare array with no success/error envelope.

### ### Known limitations
- `ok: true` / no error only means the underlying script completed without throwing — it does not confirm a macro did what was intended (cell writes, file output, Immediate window output aren't observable through these tools). See the README's AI client section for more detail.

## [0.0.30] - 2026-07-25
### ### Fixed
- Fixed Marketplace README rendering: replaced shields.io `visual-studio-marketplace` badges (retired by shields.io, showing as broken placeholders) with `vsmarketplacebadges.dev` equivalents, and removed stray tab characters inside markdown table cells that were breaking table rendering.

## [0.0.29] - 2026-07-25
### ### Fixed
- Fixed garbled Japanese description text for the "Export Folder" setting in `package.json` (#2).
- Fixed import failure (`Attribute <proc>.VB_ProcData.VB_Invoke_Func`) when importing a module containing a procedure with an assigned macro shortcut key. Standard/class module import (`.bas`/`.cls`) now uses `VBComponents.Import()` — the same approach already used for `.frm` — instead of `CodeModule.AddFromString()`, so all Attribute lines (including shortcut key assignments) are preserved on import (#3).
- Fixed a non-fatal `PropertyNotFound` error ("プロパティ 'Visible' が見つかりません") logged to the Output channel during export, caused by `$excel.VBE.MainWindow` sometimes not being available yet (e.g. VBE never opened in the session) when the script tried to bring it to the foreground. The assignment is now wrapped in `try/catch` in `export_opened_vba.ps1`.

### ### Improved
- Hardened the flowchart generation pipeline: escaped single quotes in PowerShell command arguments (paths containing `'` no longer break the command), verify the intermediate `.flow.json` file actually exists before running the Mermaid conversion step, and surface failures via error/info popups in addition to the Output channel.

### ### Known limitations
- Shortcut key assignments on Sheet/ThisWorkbook code-behind procedures (component type = Document) are still lost on import. `VBComponents.Import()` cannot target Document-type components (VBA API constraint), so this path still requires stripping all Attribute lines before `AddFromString()`.

## [0.0.28] - 2025-11-03
### ### Changed
- Create an mmd folder in the export destination and output a simple flowchart in Mermaid format (*.mmd) as an experimental feature.
- Fine-tuned message text.

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

