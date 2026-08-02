# 📤 EXCEL VBA module Sync - VSCode ⇄ Excel
![Installs](https://vsmarketplacebadges.dev/installs/9kv8xiyi.excel-vba-sync.svg)
![Version](https://vsmarketplacebadges.dev/version/9kv8xiyi.excel-vba-sync.svg)
![Rating](https://vsmarketplacebadges.dev/rating-short/9kv8xiyi.excel-vba-sync.svg)

## 目次 / Contents
- [概要 / Overview](#overview)
- [2つの使い方 / Two Ways to Use This](#two-ways)
- [インストール / Install](#install)
- [🤖 AIクライアントから使う（Claude Code / Claude Desktop） / Use from an AI client](#ai-usage)
- [⚠重要 / Important](#important)
- [🛠 開発者向け情報 / Development](#dev)

## <a id="overview"></a>概要（Japanese）

**EXCEL VBA module Sync** は、開いているExcel の VBA モジュールを VSCode 上で編集するための拡張機能です。  
VBA モジュールのVSCodeへのエクスポート、VSCodeで編集した内容のVBAへのインポートが行えます。  
**Winsdows10/11＋Excel＋VSCode環境でのみ動作します。**

この拡張機能は、**①VS Code上で手動編集**するのと、**②AIエージェントに読み書きさせる**のと、2つの使い方ができます。両者はデータの流れが根本的に違うので、詳しくは次の「[2つの使い方](#two-ways)」をご覧ください。

- ✅ 開いているExcelブックから `.bas` / `.cls` / `.frm` 内のコードをエクスポート（保存）
- ✅ VSCode 上で編集
- ✅ 編集したモジュールを 開いているExcelブック にインポート（反映）
- ✅ インポートはモジュール差し替えにて行います
- ✅ エクスポートしたモジュールファイルはgitで管理しやすいようにUTF-8の文字コードで出力されます。
- ✅ Excel マクロ実行（モジュール＋プロシージャ、または完全修飾名で実行可能）
- ✅ VBA コード検索（開いている全ブック・全モジュールを対象に、正規表現やフィルタ指定で検索可能）
- ✅ フローチャート作成（右クリックで選択したモジュールのフローチャートを作成）

### 🔧 主な機能

- **Export All Modules From VBA** — 開いているExcel から全モジュールを抽出・保存
- **Import Module To VBA** — VSCode 上で編集したコードを Excel に反映（単独モジュール/ファイル）
- **Set Export Folder** — エクスポート先フォルダをダイアログにて選択
- **Excel VBA: List & Run Macro** — 開いているブックより指定マクロを実行
- **Excel VBA: Search VBA Code** — 開いているブック・モジュール対象にコード検索（正規表現対応）
- **Generate VBA Flow Chart**（実験的機能） — エクスポート先フォルダに mmdフォルダを作成し、マーメイド形式(*.mmd)で簡易フローチャートを出力
- **Excel VBA: Print MCP Server Config (for AI)** — AIクライアント（Claude Code/Desktop等）向けのMCP接続設定を生成
- **コマンドパレット／ボタン対応** — GUI 操作または `Ctrl+Shift+P` から実行可

---

## Overview (English)

**EXCEL VBA module Sync** is a VSCode extension for editing opened Excel VBA modules.  
You can export VBA modules to VS Code and import the content edited in VS Code back into VBA.
Works in a Windows 10/11 + Excel + VS Code environment only.

This extension can be used in two ways: **① manual editing in VS Code**, or **② letting an AI agent read/write for you**. The two have fundamentally different data flows — see "[Two Ways to Use This](#two-ways)" below for details.

- ✅ Export inner code of  `.bas` / `.cls` / `.frm` from opened Excel
- ✅ Edit VBA modules in VSCode
- ✅ Import modules back into opened Excel
- ✅ Import is performed by replacing the module.
- ✅ Exported module files are saved in UTF-8 encoding, making them easier to manage with Git.
- ✅ Execute Excel macros (by module/procedure or fully qualified name)
- ✅ Search VBA code (across all open workbooks/modules, with regex and filters supported)
- ✅ Generate flowchart (Create a flowchart for the module selected with right-click)

### 🔧 Features

- **Export All Modules From VBA** — Extract and save all VBA modules from opened Excel
- **Import Module To VBA** — Reflect modified code back to opened Excel (module-based/file-based)
- **Set Export Folder** — Change export folder via dialog
- **Excel VBA: List & Run Macro** — Execute macros by name or fully qualified path in the open workbook
- **Excel VBA: Search VBA Code** — Search VBA code in the open workbook (with regex support)
- **Generate VBA Flow Chart** (Experimental) — Create an mmd folder in the export destination and output a simple flowchart in Mermaid format (*.mmd)
- **Excel VBA: Print MCP Server Config (for AI)** — Generate an MCP connection config for AI clients (Claude Code/Desktop, etc.)
- **Command Palette / GUI support** — Use commands or side panel buttons

---

## <a id="two-ways"></a>🔀 2つの使い方 / Two Ways to Use This

### ① 手動で使う（VS Code編集）
Excel VBE の中身は、**エクスポート先フォルダのテキストファイル（`.bas`/`.cls`/`.frm`）を経由して**やり取りします。

```
Excel VBE  --[Export]-->  エクスポートフォルダの .bas/.cls/.frm
                                        │
                                  VS Code で編集
                                        │
Excel VBE  <--[Import]--  編集したファイル
```

**Importボタンを押すまで、VS Code上の変更はExcel側に反映されません。** 逆に、Excel側でVBEを直接編集した場合も、再度Exportするまでテキストファイル側には反映されません。両者は明示的な操作（Export/Import）でのみ同期する、ズレうる2つのコピーです。

### ② AIエージェントから使う（MCP）
AIクライアント（Claude Code・Claude Desktop・Copilot Chat・Codex等）は、**エクスポートフォルダを経由せず、Excelの「今開いているVBAプロジェクト」を直接読み書きします**（内蔵のMCPサーバーがCOM経由で操作）。

```
AIクライアント（Claude Code / Copilot Chat 等）
        │  MCP
        ▼
内蔵MCPサーバー  ──COM経由で直接──>  Excel VBE（今開いているプロジェクト）
```

**AIがコードを書き込むと、Excel VBE上に即座に反映されます**（①のようにImportボタンを押す必要はありません）。書き込みには`dryRun`→`confirmToken`の2段階確認フロー・自動バックアップ・複数クライアント同時利用時の排他制御が入っており、AIエージェントが一発で書き込んでしまうことはありません。詳しくは後述の「[🤖 AIクライアントから使う](#ai-usage)」を参照してください。

**①②は独立しています。** AIがExcelに直接書き込んだ内容を、①のエクスポート済みファイルへ反映したい場合は、改めて手動でExportしてください（自動では同期されません）。

---

### ① Manual editing in VS Code
The VBE's contents move through **text files in the export folder** (`.bas`/`.cls`/`.frm`).

```
Excel VBE  --[Export]-->  .bas/.cls/.frm files in the export folder
                                        │
                                  edit in VS Code
                                        │
Excel VBE  <--[Import]--  the edited files
```

**Changes in VS Code aren't reflected in Excel until you click Import.** Likewise, edits made directly in the VBE aren't reflected in the text files until you Export again. The two are separate copies that only sync on an explicit Export/Import action.

### ② Using an AI agent (MCP)
AI clients (Claude Code, Claude Desktop, Copilot Chat, Codex, etc.) **bypass the export folder entirely** and read/write the "currently open VBA project" in Excel directly (the built-in MCP server operates via COM).

```
AI client (Claude Code / Copilot Chat / etc.)
        │  MCP
        ▼
Built-in MCP server  ──directly via COM──>  Excel VBE (the currently open project)
```

**When an AI writes code, it's reflected in the VBE immediately** — no Import button needed, unlike ①. Writes go through a two-step `dryRun`/`confirmToken` confirmation flow, automatic backups, and concurrency control for multiple simultaneous clients, so an AI agent can't write blind in one shot. See "[Use from an AI client](#ai-usage)" below for details.

**① and ② are independent.** If an AI writes directly to Excel and you also want that reflected in ①'s exported files, Export manually again afterward — it does not happen automatically.

---
## <a id="install"></a>🧩 インストール（VSIX） / Install from VSIX

### From Marketplace
1. [Visual Studio Marketplace - excel-vba-sync](https://marketplace.visualstudio.com/items?itemName=9kv8xiyi.excel-vba-sync)  
2. Visual Studio Code を開き、拡張機能ビューからインストール  

## 拡張機能ビューからできない場合は、以下をお試しください。

### From Marketplace(Powershell)
以下コマンドを実行
```powershell
code --install-extension 9kv8xiyi.excel-vba-sync
```

### From Github(VSCode)
1. https://github.com/EitaroSeta/excel-vba-sync/releases/download/latest/extension.vsix より`extension.vsix`をダウンロード
2. VS Code を開く
3. 拡張機能ビュー（Ctrl+Shift+X / Cmd+Shift+X）を開く
4. 右上の「…」メニュー → **VSIXからのインストール...** を選択
5. ダウンロードした`extension.vsix` ファイルを選択
6. Reloadを実行

### From Github(Powershell)
以下コマンドを実行
```powershell
$URL = "https://github.com/EitaroSeta/excel-vba-sync/releases/download/latest/extension.vsix"
$OUT = "$env:TEMP\extension.vsix"
curl.exe -sS -L -f --retry 3 --retry-delay 2 "$URL" -o "$OUT"
code --install-extension "$OUT"
```

## <a id="ai-usage"></a>🤖 AIクライアントから使う（Claude Code / Claude Desktop） / Use from an AI client

この拡張機能には[Model Context Protocol (MCP)](https://modelcontextprotocol.io/)サーバーが内蔵されており、Claude Code・Claude Desktop・VS Code内蔵のCopilot Chat・Codex等のAIクライアントから、VBAコードの読み取り・検索・マクロ実行に加えて、**AIが書いたコードをExcelのVBAモジュールへ直接書き込む**ことができます（上記「[2つの使い方](#two-ways)」の②）。

### 提供される13のツール / 13 available tools

| ツール / Tool | できること / What it does |
|---|---|
| `ping` | 疎通確認（サーバーが応答するか） / Health check |
| `excel_list_modules` | モジュール名・種別・行数の一覧を軽量取得 / List modules (name, type, line count) |
| `excel_get_module_code` | モジュールのソースコード全体を読み取り / Read a module's full source code |
| `vba_search_code` | 全ブック・全モジュール横断でコード検索（正規表現対応） / Search code across all open workbooks/modules (regex supported) |
| `vba_analyze_flow` | プロシージャの制御フロー（分岐・ループ・呼び出し）を構造化JSONで取得 / Get a procedure's control-flow structure (branches, loops, calls) as structured JSON |
| `vba_render_flowchart` | プロシージャの制御フロー、またはモジュールの呼び出しグラフをMermaid図として取得 / Get a procedure's control flow, or a module's call graph, as a Mermaid diagram |
| `vba_list_dependencies` | Windows API宣言・CreateObject・Shell呼び出し等を正規表現で一覧化 / List Windows API declares, CreateObject calls, Shell calls, etc. via regex matching |
| `vba_list_references` | イベントプロシージャ・シート/名前付き範囲参照を正規表現で一覧化 / List event procedures and sheet/named-range references via regex matching |
| `vba_list_variable_scopes` | 変数・定数宣言をスコープ別に一覧化、または指定変数の使用箇所をスコープを考慮して検索 / List variable/constant declarations by scope, or scope-aware search for a given variable's usages |
| `excel_list_macros` | 実行可能なマクロ（Public Sub）の一覧を取得 / List runnable macros (Public Subs) |
| `excel_run_macro` | マクロを実行 / Run a macro |
| `excel_read_range` | セル範囲の値を読み取り（マクロの実行結果検証等に） / Read cell values from a range (e.g. to verify a macro's effect) |
| `excel_update_module_code` | モジュールのコードを書き込み（`dryRun`/`confirmToken`の安全フロー付き） / Write a module's code (with the `dryRun`/`confirmToken` safety flow) |

セットアップ手順・安全設計・既知の注意点は **[docs/AI_USAGE.md](docs/AI_USAGE.md)** を参照してください。

This extension has a built-in [Model Context Protocol (MCP)](https://modelcontextprotocol.io/) server, allowing AI clients — Claude Code, Claude Desktop, VS Code's own Copilot Chat, Codex, etc. — to read, search, and run VBA code, and even write AI-authored code directly into an Excel VBA module (option ② in "[Two Ways to Use This](#two-ways)" above).

See **[docs/AI_USAGE.md](docs/AI_USAGE.md)** for setup steps, safety design, and known caveats.

## <a id="important"></a>⚠重要 / Important ##

> [!WARNING]
> **EXCELファイルは必ずバックアップしてください。**
> この拡張機能はCOM経由でExcelファイルを外部から操作するため、条件によりファイルを破損させる恐れがあります。特に、AIエージェントがMCP経由で自動的にコードの読み書き・マクロ実行を行う運用では、人が都度確認する場合に比べて**意図しない上書き・実行が起きるリスクが通常の利用より高くなります**。作業前に必ずバックアップを取得してください。
>
> **Always back up your Excel files before use.** This extension operates on Excel files externally via COM, so there is a risk of file corruption depending on conditions. This risk is **higher than normal** when an AI agent is driving the tools autonomously via MCP (automatic code read/write and macro execution), since there is no per-step human confirmation. Always back up your files before starting.

**●エクスポートしたファイルの属性は編集しないでください**
> エクスポートした **`.frm/.cls/.bas`** の **属性行は編集しないでください**。`VERSION`、`Begin … End`、`Object = …`、および `Attribute VB_*`（例：`VB_Name` / `VB_PredeclaredId` / `VB_Exposed` / `VB_Creatable` など）を変更すると、**インポート失敗**や**既存フォームとの紐付け崩れ**が発生します。  

**●モジュールの新規追加はできません**
>既存のモジュール/クラス/フォームを入替えを行う仕組みの為、新規の追加はできません。 VBA上で新規モジュールを追加し、エクスポートしてください。

**●COMエラーについて**
>Excel に長時間触れずに放置した後や、画面ロック復帰直後などにインポート／エクスポートを実行すると、
次のようなエラーが発生する場合があります。  
`STDERR: Call was rejected by callee. (HRESULT からの例外:0x80010001 (RPC_E_CALL_REJECTED))`    
>これは Excel 側が一時的に応答できない状態にあるため、COM 呼び出しが失敗して発生するエラーです。  
この場合は **Excelを再起動**すると解消されます。

**●Do **not** edit attributes of exported files**
> Do **not edit the attribute lines** in exported **`.frm/.cls/.bas`** files. Changing `VERSION`, `Begin … End`, `Object = …`, or any `Attribute VB_*` (e.g., `VB_Name`, `VB_PredeclaredId`, `VB_Exposed`, `VB_Creatable`) can cause **import failures**,  and **loss of linkage** to the original form.  

**●New modules, classes, or forms cannot be added;**
>New modules, classes, or forms cannot be added; this tool only replaces existing ones.If you need to create a new item, first add a blank module/class/form in the VBE, then export it.

**●About COM Error**
>When running import/export operations after leaving Excel idle for a long time or resuming from a screen lock,
you may encounter the following error:  
`STDERR: Call was rejected by callee. (HRESULT 0x80010001)`  
>This occurs because Excel is temporarily unable to respond, causing the COM call to fail.  
**Restarting Excel** will resolve the issue.

### 免責事項 / Disclaimer
本拡張機能は現状有姿（as-is）で提供されます。MCPサーバー経由でAIエージェントが行った操作を含め、本拡張機能の使用によって生じたExcelファイルの破損・データ損失・その他いかなる損害についても、作者は一切の責任を負いません。自己責任でご利用の上、必ず事前にバックアップを取得してください。

This extension is provided "as is", without warranty of any kind. The author accepts no liability for any damage — including but not limited to file corruption or data loss — resulting from use of this extension, including actions taken by an AI agent via the MCP server. Use at your own risk, and always back up your Excel files beforehand.

---

## <a id="dev"></a>🛠 開発者向け情報 / Development (for GitHub users)
このセクションは拡張機能の利用者には不要です。拡張の開発や修正、ローカライズ設定、アーキテクチャ図は **[docs/DEVELOPMENT.md](docs/DEVELOPMENT.md)** を参照してください。

This section is unnecessary for extension users. See **[docs/DEVELOPMENT.md](docs/DEVELOPMENT.md)** for build/dev setup, packaging, localization, and the architecture diagram.

**開発体制について**: v0.0.28以降の実装は、Claude（Anthropic）とのAI協働（vibe coding）によって行われています。詳細は[docs/DEVELOPMENT.md](docs/DEVELOPMENT.md)を参照してください。

**About the development process**: Since v0.0.28, implementation has been done via AI-assisted development ("vibe coding") with Claude (Anthropic). See [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md) for details.

https://github.com/EitaroSeta/excel-vba-sync

