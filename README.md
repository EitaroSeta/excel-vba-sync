# 📤 EXCEL VBA module Sync - VSCode ⇄ Excel
![Installs](https://vsmarketplacebadges.dev/installs/9kv8xiyi.excel-vba-sync.svg)
![Version](https://vsmarketplacebadges.dev/version/9kv8xiyi.excel-vba-sync.svg)
![Rating](https://vsmarketplacebadges.dev/rating-short/9kv8xiyi.excel-vba-sync.svg)

開いているExcelのVBAモジュールを、VS Code上で扱う拡張機能です。
A VS Code extension for working with the VBA modules of an open Excel workbook.

**[📖 日本語で読む](#ja) ｜ [📖 Read in English](#en)**

---

<a id="ja"></a>

## 目次

- [概要](#ja-overview)
- [2つの使い方](#ja-two-ways)
- [内蔵MCPサーバーが提供する機能](#ja-ai-usage)
- [インストール](#ja-install)
- [⚠ 重要](#ja-important)
- [開発者向け情報](#ja-dev)

## <a id="ja-overview"></a>概要

**EXCEL VBA module Sync** は、開いているExcelのVBAモジュールをVS Code上で扱う拡張機能です。

VBAモジュールをVS Codeへエクスポートし、編集した内容をExcelへインポートできます。さらに、内蔵のMCPサーバー経由でAIエージェントに直接読み書きさせることもできます。

**Windows 10/11 ＋ Excel ＋ VS Code 環境でのみ動作します。**

### 使い方は2通り

データの流れが根本的に違います。詳しくは「[2つの使い方](#ja-two-ways)」を参照してください。

| 区分 | ✍️ ① VS Codeで手動編集 | 🤖 ② AIエージェントから操作（MCP） |
|---|---|---|
| **入出力** | エクスポート（UTF-8、Git管理しやすい）／インポート（差し替え） | エクスポート不要。開いているVBAプロジェクトを直接読み書き |
| **編集機能** | VS Code上で自分で編集 | AIが書いたコードを直接書き込み（`dryRun`→確認の2段階、自動バックアップ付き） |
| **検索機能** | VS Codeコマンドパレットから検索（正規表現対応） | 全ブック・全モジュール横断で検索 |
| **コード解析支援** | 簡易フローチャート作成（右クリック、実験的機能） | 制御フロー・呼び出しグラフ（Mermaid）／外部依存・イベント・変数スコープの棚卸し |
| **ワークシート解析** | — | シート・名前付き範囲・フォームコントロールの実在確認／セル数式・条件付き書式・入力規則の可視化 |
| **実行・検証** | VS Codeコマンドパレットから実行 | マクロ実行、セル値の読み取りで結果を確認 |

### 🔧 コマンド一覧（①手動で使う場合）

- **Export All Modules From VBA** — 開いているExcelから全モジュールを抽出・保存
- **Import Module To VBA** — VS Codeで編集したコードをExcelへ反映（単独モジュール／ファイル）
- **Set Export Folder** — エクスポート先フォルダをダイアログで選択
- **Excel VBA: List & Run Macro** — 開いているブックの指定マクロを実行
- **Excel VBA: Search VBA Code** — 開いているブック・モジュールを対象にコード検索（正規表現対応）
- **Generate VBA Flow Chart**（実験的機能） — エクスポート先に`mmd`フォルダを作り、Mermaid形式（`*.mmd`）で簡易フローチャートを出力
- **Excel VBA: Print MCP Server Config (for AI)** — AIクライアント向けのMCP接続設定を生成
- **コマンドパレット／ボタン対応** — GUI操作または`Ctrl+Shift+P`から実行

## <a id="ja-two-ways"></a>🔀 2つの使い方

### ✍️ ① 手動で使う（VS Code編集）

Excel VBEの中身は、**エクスポート先フォルダのテキストファイル**（`.bas`/`.cls`/`.frm`）を経由してやり取りします。

```
Excel VBE  --[Export]-->  エクスポートフォルダの .bas/.cls/.frm
                                        │
                                  VS Code で編集
                                        │
Excel VBE  <--[Import]--  編集したファイル
```

**Importボタンを押すまで、VS Code上の変更はExcelに反映されません。** 逆も同様で、VBEを直接編集してもExportするまでテキストファイルには反映されません。

両者は独立した2つのコピーです。明示的な操作（Export／Import）でのみ同期します。

#### 💬 保存時のインポート提案

Importの押し忘れを防ぐため、**エクスポートフォルダ配下**の`.bas`/`.cls`/`.frm`をVS Codeで保存（`Ctrl+S`）すると、右下に通知が出ます：

> 📄 `Module1.bas` を保存しました。Excelへインポートしますか？　**[インポート]** **[今後表示しない]**

- **[インポート]** — そのファイル1つを既存のImport処理でExcelへ反映します（結果は出力パネルに表示）
- **通知を閉じる（×）** — 何もしません。次に保存したときにまた提案されます
- **[今後表示しない]** — 通知を完全にオフにします（下記の設定がオフに切り替わります）

勝手にインポートすることはありません。編集途中の保存で未完成のコードがExcelへ飛ばないよう、必ずボタンを押したときだけ実行されます。エクスポートフォルダの外のファイルや、`.bas`/`.cls`/`.frm`以外の保存では何も出ません。

オン／オフは設定 `excelVbaSync.importPromptOnSave` で切り替えられます（既定：オン）。設定画面（`Ctrl+,`）で「excel vba sync」と検索するとチェックボックスが見つかります。「今後表示しない」を押した後に復活させたい場合も、ここを再度オンにしてください。

### 🤖 ② AIエージェントから操作（MCP）

AIクライアント（Claude Code・Claude Desktop・Copilot Chat・Codex等）は、**エクスポートフォルダを経由しません**。内蔵のMCPサーバーがCOM経由で、Excelの「今開いているVBAプロジェクト」を直接読み書きします。

```
AIクライアント（Claude Code / Copilot Chat 等）
        │  MCP
        ▼
内蔵MCPサーバー  ──COM経由で直接──>  Excel VBE（今開いているプロジェクト）
```

**AIがコードを書き込むと、VBE上に即座に反映されます。** ①のようなImportボタンは不要です。

ただし一発では書き込めません。書き込みには`dryRun`→`confirmToken`の2段階確認、自動バックアップ、複数クライアント利用時の排他制御が入っています。詳しくは「[内蔵MCPサーバーが提供する機能](#ja-ai-usage)」を参照してください。

**①と②は独立しています。** 既定では自動同期しませんが、橋渡しは2つあります：AIがExcelへ書き込んだ内容は`excel_export_module`ツールでそのモジュールだけエクスポートファイルへ反映でき（AIエージェントが書き込み後に自分で呼びます）、逆方向は上記「[保存時のインポート提案](#ja-two-ways)」が編集済みファイルのインポートを提案します。さらにAI→ファイル方向は、設定`excelVbaSync.mcpSyncMode`（既定`remind`：書き込み応答でAIにエクスポートを指示。`auto`にするとサーバーが自動でエクスポート）が同期を確実にします。いずれも対象ブックのエクスポートフォルダが既に存在するときだけ動くので、エクスポートを使わない運用には影響しません。

## <a id="ja-ai-usage"></a>内蔵MCPサーバーが提供する機能

この拡張機能には [Model Context Protocol (MCP)](https://modelcontextprotocol.io/) サーバーが内蔵されています。

Claude Code・Claude Desktop・Copilot Chat・Codex等のAIクライアントから、VBAコードの読み取り・検索・解析・マクロ実行に加えて、**AIが書いたコードをExcelのVBAモジュールへ直接書き込む**ことができます（上記②）。

### 提供される20のツール

| ツール | できること |
|---|---|
| `ping` | 疎通確認（サーバーが応答するか） |
| `excel_list_modules` | モジュール名・種別・行数の一覧を軽量取得 |
| `excel_get_module_code` | モジュールのソースコード全体を読み取り |
| `vba_search_code` | 全ブック・全モジュール横断でコード検索（正規表現対応） |
| `vba_analyze_flow` | プロシージャの制御フロー（分岐・ループ・呼び出し）を構造化JSONで取得 |
| `vba_render_flowchart` | プロシージャの制御フロー、またはモジュールの呼び出しグラフをMermaid図として取得 |
| `vba_list_dependencies` | Windows API宣言・CreateObject・Shell呼び出し等を一覧化 |
| `vba_list_references` | イベントプロシージャ・シート／名前付き範囲参照を一覧化 |
| `vba_list_variable_scopes` | 変数・定数宣言をスコープ別に一覧化、または指定変数の使用箇所を検索 |
| `excel_list_macros` | 実行可能なマクロ（Public Sub）の一覧を取得 |
| `excel_run_macro` | マクロを実行 |
| `excel_read_range` | セル範囲の値を読み取り（マクロの実行結果検証等に） |
| `excel_update_module_code` | 既存モジュール（UserFormのコード部分も可）へ書き込み、または`moduleType`指定で新規モジュールを作成。`dryRun`/`confirmToken`の安全フロー付き、重複プロシージャ名も警告 |
| `excel_export_module` | 指定モジュール1つをエクスポートフォルダへ書き出し（書き込み後にディスク上のファイルを最新化。①の手動編集と組み合わせる運用向け） |
| `excel_list_worksheets` | 実在するシート一覧を取得（表示名・VBAコード名・表示状態） |
| `excel_list_form_controls` | UserForm内のコントロール一覧を取得（名前・種類） |
| `excel_list_defined_names` | 実在する名前付き範囲の一覧を取得（参照先・壊れているかどうか） |
| `excel_list_formulas` | シート上の数式を取得（同一パターンはグルーピング） |
| `excel_list_conditional_formats` | 条件付き書式のルール一覧を取得（同一パターンはグルーピング） |
| `excel_list_data_validations` | 入力規則の一覧を取得（ドロップダウンの選択肢等） |

コード内容を返すツールは、ハードコードされたパスワード・APIキーらしき値を自動的に`[REDACTED]`へマスクします（常時有効、無効化不可）。対象は`excel_get_module_code`・`excel_update_module_code`・`vba_search_code`・`vba_list_dependencies`・`vba_list_references`・`vba_list_variable_scopes`・`vba_analyze_flow`・`vba_render_flowchart`です。

セットアップ手順・安全設計・既知の注意点は **[docs/AI_USAGE.md](docs/AI_USAGE.md)** を参照してください。

## <a id="ja-install"></a>🧩 インストール

### Marketplaceから

1. [Visual Studio Marketplace - excel-vba-sync](https://marketplace.visualstudio.com/items?itemName=9kv8xiyi.excel-vba-sync) を開く
2. VS Codeの拡張機能ビューからインストール

拡張機能ビューからインストールできない場合は、以下をお試しください。

### Marketplaceから（PowerShell）

```powershell
code --install-extension 9kv8xiyi.excel-vba-sync
```

### GitHubから（VS Code）

1. [extension.vsix](https://github.com/EitaroSeta/excel-vba-sync/releases/download/latest/extension.vsix) をダウンロード
2. VS Codeを開く
3. 拡張機能ビュー（`Ctrl+Shift+X`）を開く
4. 右上の「…」メニュー → **VSIXからのインストール...**
5. ダウンロードした`extension.vsix`を選択
6. Reloadを実行

### GitHubから（PowerShell）

```powershell
$URL = "https://github.com/EitaroSeta/excel-vba-sync/releases/download/latest/extension.vsix"
$OUT = "$env:TEMP\extension.vsix"
curl.exe -sS -L -f --retry 3 --retry-delay 2 "$URL" -o "$OUT"
code --install-extension "$OUT"
```

## <a id="ja-important"></a>⚠ 重要

> [!WARNING]
> **Excelファイルは必ずバックアップしてください。**
>
> この拡張機能はCOM経由でExcelファイルを外部から操作します。条件によってはファイルを破損させる恐れがあります。
>
> 特にAIエージェントがMCP経由で自動的に読み書き・マクロ実行を行う運用では、**意図しない上書き・実行のリスクが通常より高くなります**（人が都度確認しないため）。作業前に必ずバックアップを取得してください。

**● エクスポートしたファイルの属性行は編集しないでください**

> エクスポートした`.frm`/`.cls`/`.bas`の**属性行を編集しないでください**。
>
> `VERSION`、`Begin … End`、`Object = …`、`Attribute VB_*`（`VB_Name`／`VB_PredeclaredId`／`VB_Exposed`／`VB_Creatable`等）を変更すると、**インポート失敗**や**既存フォームとの紐付け崩れ**が発生します。

**● 手動インポートではモジュールの新規追加はできません**

> ①の手動インポートは既存モジュール／クラス／フォームの差し替えを行う仕組みです。新規追加はできません。VBE上で新規モジュールを追加してから、エクスポートしてください。
>
> ②のMCP経由（`excel_update_module_code`）であれば、標準モジュール・クラスモジュールの新規作成は可能です。ただしUserFormの新規作成はできません。

**● COMエラーについて**

> Excelを長時間放置した後や、画面ロック復帰直後にインポート／エクスポートを実行すると、次のエラーが出ることがあります。
>
> `STDERR: Call was rejected by callee. (HRESULT からの例外:0x80010001 (RPC_E_CALL_REJECTED))`
>
> Excelが一時的に応答できない状態のため、COM呼び出しが失敗して発生します。**Excelを再起動**すると解消されます。

**● セキュリティ製品（アンチウイルス／EDR）にブロックされる場合があります**

> この拡張機能は、PowerShell経由でExcelのVBAプロジェクトを読み書きします。これは**マクロ型マルウェアが行う操作と原理的に同じ**です。
>
> そのため環境によっては、アンチウイルスやEDRが動作をブロックしたり警告を出したりする可能性があります（特に企業環境でDefenderのASR〈攻撃対象領域の縮小〉ルールが有効な場合）。
>
> そもそもVBAプロジェクトへのアクセスには、Trust Centerの「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」を**利用者が明示的に有効化する必要があります**（既定は無効）。この設定が既定でオフなのは、ここがマルウェアの侵入経路になりうるためです。
>
> 本拡張機能は、利用者自身の判断でそのゲートを開けた上で使うことを前提としています。ブロックされた場合に除外設定を入れるかどうかも、利用者自身のリスク判断に委ねられます。

### 免責事項

本拡張機能は現状有姿（as-is）で提供されます。

MCPサーバー経由でAIエージェントが行った操作を含め、本拡張機能の使用によって生じたExcelファイルの破損・データ損失・その他いかなる損害についても、作者は一切の責任を負いません。自己責任でご利用の上、必ず事前にバックアップを取得してください。

## <a id="ja-dev"></a>🛠 開発者向け情報

このセクションは拡張機能の利用者には不要です。

拡張の開発や修正、ローカライズ設定、アーキテクチャ図は **[docs/DEVELOPMENT.md](docs/DEVELOPMENT.md)** を参照してください。

**開発体制について**: v0.0.28以降の実装は、Claude（Anthropic）とのAI協働（vibe coding）によって行われています。詳細は [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md) を参照してください。

---

<a id="en"></a>

## Contents

- [Overview](#en-overview)
- [Two Ways to Use This](#en-two-ways)
- [What the built-in MCP server provides](#en-ai-usage)
- [Install](#en-install)
- [⚠ Important](#en-important)
- [Development](#en-dev)

## <a id="en-overview"></a>Overview

**EXCEL VBA module Sync** is a VS Code extension for working with the VBA modules of an open Excel workbook.

You can export VBA modules to VS Code and import your edits back into Excel. You can also let an AI agent read and write them directly, through the built-in MCP server.

**Works only on Windows 10/11 with Excel and VS Code.**

### Two ways to use it

The two have fundamentally different data flows. See "[Two Ways to Use This](#en-two-ways)" for details.

| Area | ✍️ ① Manual editing in VS Code | 🤖 ② Driving it from an AI agent (MCP) |
|---|---|---|
| **Import / export** | Export (UTF-8, Git-friendly) and Import (by replacement) | No export needed. Reads and writes the open VBA project directly |
| **Editing** | You edit it yourself in VS Code | An AI writes code straight in (two-step `dryRun` → confirm, with automatic backups) |
| **Search** | Search from the VS Code Command Palette (regex supported) | Search across all open workbooks and modules |
| **Code analysis** | Generate a simple flowchart (right-click, experimental) | Control flow and call graphs (Mermaid); inventory of external dependencies, events and variable scopes |
| **Worksheet analysis** | — | Verify that sheets, defined names and form controls exist; reveal cell formulas, conditional formatting and data validation |
| **Running / verifying** | Run macros from the VS Code Command Palette | Run macros, then read cell values to check the result |

### 🔧 Commands (for ① manual use)

- **Export All Modules From VBA** — Extract and save all VBA modules from the open workbook
- **Import Module To VBA** — Reflect code edited in VS Code back into Excel (module-based / file-based)
- **Set Export Folder** — Change the export folder via a dialog
- **Excel VBA: List & Run Macro** — Run a macro in the open workbook
- **Excel VBA: Search VBA Code** — Search code across open workbooks and modules (regex supported)
- **Generate VBA Flow Chart** (experimental) — Create an `mmd` folder under the export destination and write a simple flowchart in Mermaid format (`*.mmd`)
- **Excel VBA: Print MCP Server Config (for AI)** — Generate an MCP connection config for AI clients
- **Command Palette / buttons** — Run from the GUI or `Ctrl+Shift+P`

## <a id="en-two-ways"></a>🔀 Two Ways to Use This

### ✍️ ① Manual editing in VS Code

The VBE's contents move through **text files in the export folder** (`.bas`/`.cls`/`.frm`).

```
Excel VBE  --[Export]-->  .bas/.cls/.frm files in the export folder
                                        │
                                  edit in VS Code
                                        │
Excel VBE  <--[Import]--  the edited files
```

**Changes in VS Code are not reflected in Excel until you click Import.** The reverse is also true: edits made directly in the VBE are not reflected in the text files until you Export again.

They are two independent copies. They only sync on an explicit Export or Import.

#### 💬 Import prompt on save

To keep you from forgetting the Import step, saving (`Ctrl+S`) a `.bas`/`.cls`/`.frm` **under the export folder** shows a notification in the bottom-right corner:

> 📄 `Module1.bas` was saved. Import it into Excel?　**[Import]** **[Don't ask again]**

- **[Import]** — runs the existing Import for that one file (the result appears in the output panel)
- **Closing the toast (×)** — does nothing; you will be asked again on the next save
- **[Don't ask again]** — turns the prompt off permanently (flips the setting below)

It never imports on its own: nothing reaches Excel unless you click the button, so a mid-edit save cannot push half-finished code. Saving files outside the export folder, or with other extensions, shows nothing.

The prompt is controlled by the `excelVbaSync.importPromptOnSave` setting (default: on). Search for "excel vba sync" in Settings (`Ctrl+,`) to find the checkbox — re-enable it there if you clicked "Don't ask again" and want the prompt back.

### 🤖 ② Driving it from an AI agent (MCP)

AI clients (Claude Code, Claude Desktop, Copilot Chat, Codex, etc.) **bypass the export folder entirely**. The built-in MCP server reads and writes the currently open VBA project directly, via COM.

```
AI client (Claude Code / Copilot Chat / etc.)
        │  MCP
        ▼
Built-in MCP server  ──directly via COM──>  Excel VBE (the currently open project)
```

**When an AI writes code, it appears in the VBE immediately.** No Import button, unlike ①.

It cannot write in one shot, though. Writes go through a two-step `dryRun`/`confirmToken` confirmation, automatic backups, and concurrency control for multiple clients. See "[What the built-in MCP server provides](#en-ai-usage)" for details.

**① and ② are independent.** Nothing syncs automatically by default, but there are two bridges: an AI agent can refresh a single exported file right after writing to Excel via the `excel_export_module` tool, and in the other direction the "[Import prompt on save](#en-two-ways)" above offers to import your edited file. The AI-to-file direction is made reliable by the `excelVbaSync.mcpSyncMode` setting (default `remind`: each write response instructs the AI to export; `auto` makes the server export by itself). Both act only when the workbook's export folder already exists, so workflows that never export are unaffected.

## <a id="en-ai-usage"></a>What the built-in MCP server provides

This extension has a built-in [Model Context Protocol (MCP)](https://modelcontextprotocol.io/) server.

AI clients such as Claude Code, Claude Desktop, Copilot Chat and Codex can read, search, analyze and run VBA code — and **write AI-authored code directly into an Excel VBA module** (option ② above).

### 20 available tools

| Tool | What it does |
|---|---|
| `ping` | Health check (is the server responding) |
| `excel_list_modules` | List modules (name, component type, line count) |
| `excel_get_module_code` | Read a module's full source code |
| `vba_search_code` | Search code across all open workbooks and modules (regex supported) |
| `vba_analyze_flow` | Get a procedure's control-flow structure (branches, loops, calls) as structured JSON |
| `vba_render_flowchart` | Get a procedure's control flow, or a module's call graph, as a Mermaid diagram |
| `vba_list_dependencies` | List Windows API declares, CreateObject calls, Shell calls, etc. |
| `vba_list_references` | List event procedures and sheet / named-range references |
| `vba_list_variable_scopes` | List variable and constant declarations by scope, or scope-aware search for one variable's usages |
| `excel_list_macros` | List runnable macros (Public Subs) |
| `excel_run_macro` | Run a macro |
| `excel_read_range` | Read cell values from a range (e.g. to verify a macro's effect) |
| `excel_update_module_code` | Overwrite an existing module (including a UserForm's code-behind), or create a new one via `moduleType`. Includes the `dryRun`/`confirmToken` safety flow and a duplicate-procedure-name warning |
| `excel_export_module` | Export one module to the export folder, refreshing the on-disk copy after an MCP write (for workflows that combine ① manual editing) |
| `excel_list_worksheets` | List actual worksheets (display name, VBA CodeName, visibility) |
| `excel_list_form_controls` | List a UserForm's controls (name and type) |
| `excel_list_defined_names` | List actual defined names (what they refer to, whether broken) |
| `excel_list_formulas` | List formulas present in a sheet's cells (grouped by pattern) |
| `excel_list_conditional_formats` | List conditional formatting rules (grouped by pattern) |
| `excel_list_data_validations` | List data validation rules, including dropdown choices |

Tools that return code content automatically mask values that look like hardcoded passwords or API keys as `[REDACTED]` (always on, no opt-out). This applies to `excel_get_module_code`, `excel_update_module_code`, `vba_search_code`, `vba_list_dependencies`, `vba_list_references`, `vba_list_variable_scopes`, `vba_analyze_flow` and `vba_render_flowchart`.

See **[docs/AI_USAGE.md](docs/AI_USAGE.md)** for setup steps, safety design, and known caveats.

## <a id="en-install"></a>🧩 Install

### From Marketplace

1. Open [Visual Studio Marketplace - excel-vba-sync](https://marketplace.visualstudio.com/items?itemName=9kv8xiyi.excel-vba-sync)
2. Install from the VS Code Extensions view

If you cannot install from the Extensions view, try one of the following.

### From Marketplace (PowerShell)

```powershell
code --install-extension 9kv8xiyi.excel-vba-sync
```

### From GitHub (VS Code)

1. Download [extension.vsix](https://github.com/EitaroSeta/excel-vba-sync/releases/download/latest/extension.vsix)
2. Open VS Code
3. Open the Extensions view (`Ctrl+Shift+X` / `Cmd+Shift+X`)
4. Use the "…" menu at the top right → **Install from VSIX...**
5. Select the downloaded `extension.vsix`
6. Reload

### From GitHub (PowerShell)

```powershell
$URL = "https://github.com/EitaroSeta/excel-vba-sync/releases/download/latest/extension.vsix"
$OUT = "$env:TEMP\extension.vsix"
curl.exe -sS -L -f --retry 3 --retry-delay 2 "$URL" -o "$OUT"
code --install-extension "$OUT"
```

## <a id="en-important"></a>⚠ Important

> [!WARNING]
> **Always back up your Excel files before use.**
>
> This extension operates on Excel files externally via COM, so there is a risk of file corruption depending on conditions.
>
> That risk is **higher than normal** when an AI agent drives the tools autonomously via MCP, since there is no per-step human confirmation. Always back up your files before starting.

**● Do not edit the attribute lines of exported files**

> Do **not edit the attribute lines** in exported `.frm`/`.cls`/`.bas` files.
>
> Changing `VERSION`, `Begin … End`, `Object = …`, or any `Attribute VB_*` (`VB_Name`, `VB_PredeclaredId`, `VB_Exposed`, `VB_Creatable`, etc.) can cause **import failures** and **loss of linkage** to the original form.

**● Manual import cannot add new modules**

> The manual import in ① works by replacing an existing module, class or form, so it cannot add new ones. Add the module in the VBE first, then export it.
>
> Via MCP (`excel_update_module_code`) in ②, new standard and class modules *can* be created. Creating a new UserForm is still not supported.

**● About COM errors**

> After leaving Excel idle for a long time, or right after resuming from a screen lock, an import/export may fail with:
>
> `STDERR: Call was rejected by callee. (HRESULT 0x80010001 RPC_E_CALL_REJECTED)`
>
> This happens because Excel is temporarily unable to respond, so the COM call fails. **Restarting Excel** resolves it.

**● Security software (antivirus / EDR) may block this extension**

> This extension reads and writes a workbook's VBA project through PowerShell. That is **fundamentally the same operation macro malware performs**.
>
> Depending on your environment, antivirus or EDR software may therefore block it or raise warnings (particularly in managed environments where Defender's ASR — Attack Surface Reduction — rules are enabled).
>
> Note that accessing the VBA project already requires you to **explicitly enable** "Trust access to the VBA project object model" in the Trust Center (off by default). It is off by default precisely because it is a malware entry point.
>
> This extension assumes you have opened that gate as your own deliberate decision. Likewise, whether to add an antivirus exclusion if you are blocked is your own risk decision to make.

### Disclaimer

This extension is provided "as is", without warranty of any kind.

The author accepts no liability for any damage — including but not limited to file corruption or data loss — resulting from use of this extension, including actions taken by an AI agent via the MCP server. Use at your own risk, and always back up your Excel files beforehand.

## <a id="en-dev"></a>🛠 Development (for GitHub users)

This section is unnecessary for extension users.

See **[docs/DEVELOPMENT.md](docs/DEVELOPMENT.md)** for build/dev setup, packaging, localization, and the architecture diagram.

**On the development process**: since v0.0.28, implementation has been done via AI-assisted development ("vibe coding") with Claude (Anthropic). See [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md) for details.
