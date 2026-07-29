[← README に戻る](../README.md)

# 🤖 AIクライアントから使う（Claude Code / Claude Desktop） / Use from an AI client

この拡張機能には[Model Context Protocol (MCP)](https://modelcontextprotocol.io/)サーバーが内蔵されており、AIクライアントからVBAコードの読み取り・検索・マクロ実行・**書き込み**ができます。

- MCPサーバー本体（`dist-server/server.js`）はNode単体で動作 → **VS Codeを開いていなくても**AIクライアントから直接起動・利用可能
- ただしExcel本体とVBA実行環境はWindows上に必要

## 1. セットアップ

### 1.1 外部AIクライアント（Claude Code / Claude Desktop）

1. VS Codeでこの拡張機能をインストールした状態で、コマンドパレット（`Ctrl+Shift+P`）から **`Excel VBA: Print MCP Server Config (for AI)`** を実行
2. 生成されたJSON（クリップボードにコピー済み）を設定ファイルへ反映
   - Claude Code: プロジェクトルートの`.mcp.json`
   - Claude Desktop: `claude_desktop_config.json`
   - 既存の設定がある場合はファイル全体を上書きせず、`mcpServers`オブジェクトに`excel-vba-sync`のエントリだけを追記する
3. AIクライアントを再起動 → 9ツール（`ping` / `excel_list_modules` / `excel_get_module_code` / `excel_list_macros` / `excel_run_macro` / `vba_search_code` / `vba_analyze_flow` / `excel_update_module_code` / `excel_read_range`）が使用可能に

### 1.2 VS Code内蔵・VS Code拡張機能型のAIクライアント（Copilot Chat / Codex 等）

VS Code標準のMCP連携機能（`contributes.mcpServerDefinitionProviders`）でこの拡張機能自身がMCPサーバーを自動登録するため、**手動設定は不要**（拡張機能をインストールするだけで検出される。`onStartupFinished`のような明示的な起動イベントも不要）。

実機で自動検出を確認済み：
- **Copilot Chat**（エージェントモード）
- **Codex**（VS Code拡張機能版）

補足：
- ツール選択画面には`vba-excel-mcp`（サーバーがMCPプロトコルで名乗る識別子）として表示される。人間向け表示名（`title`）・使い方要約（`instructions`）も設定済みだが、表示するかはクライアント次第
- 自動検出されない場合はクライアント独自の手動登録フォーム（Codexの「カスタムMCPに接続する」等）を試す。値は上記「Print MCP Server Config」の出力（`command`/`args`/`env`）をそのまま使える

### 1.3 複数クライアントを同時に使う場合の注意

Claude Code/Desktop（手動設定）とVS Code内蔵クライアント（自動検出）を同時に使うと、それぞれ独立した`server.js`プロセスが起動する（MCPの標準的な設計）。複数クライアントが同時にExcel操作を行うと、4.2「複数Excelプロセスに関する注意」のリスクが顕在化しやすい。

## 2. ツールの使い方のコツ

- **モジュール一覧だけ欲しい場合** → `excel_list_modules`を使う。`vba_search_code`に`.*`のような全行マッチ正規表現を投げない（コード内容を読まないぶん軽量・高速。力技の全文検索で巨大な結果を返しエラーになった実例あり）
- **ワークブック全体のマクロ一覧が欲しい場合** → `excel_list_macros`の`moduleName`を省略する。1回の呼び出しで全モジュール分を返す。モジュールごとに1回ずつ呼ぶループは避ける（Excel解決処理がモジュール数だけ繰り返され、応答が返ってこないように見えた実例あり）

## 3. 書き込み（`excel_update_module_code`）の安全設計

- `dryRun:true`で呼ぶと書き込まずに差分プレビュー＋`confirmToken`を返す
- 同じ`confirmToken`を付けて再度呼ぶと初めて書き込まれる（AIの一発書き込みを防ぐ2段階フロー）
- `confirmToken`はdry-run時点のコードに紐付き、書き込み直前に再読込・再計算して照合（楽観的排他制御）。dry-run〜確定の間に別クライアントが同じモジュールを書き換えていた場合は`ERR_MODULE_CHANGED_SINCE_DRYRUN`で拒否（後勝ち上書きを防止。実機で競合再現・正しい拒否を確認済み）
- 書き込み前に現在のコードを`.excel-vba-sync-backups`へタイムスタンプ付きで自動退避
- Sheet/ThisWorkbook（Documentモジュール）はショートカットキー等のAttribute情報がVBA API制約上失われる（レスポンスに警告あり）

**書き込み前の簡易静的解析（`lintWarnings`）**
`dryRun`レスポンスに含まれる。正規表現ベースの簡易チェック（advisory、書き込みはブロックしない）：
- `Select`/`Activate`/`Selection`/`ActiveSheet`/`ActiveWorkbook`の使用
- `Option Explicit`漏れ
- `UsedRange`依存
- 裸の`End`
- 200行超の長いプロシージャ
- オブジェクト代入時の`Set`忘れ
- `PtrSafe`なしの`Declare`
- ファイル番号の固定使用

未使用変数・デッドコード検出等のより広範なルールは未実装。

## 4. Excelプロセスの自動解決

### 4.1 workbookPathによる自動起動・自動オープン

`workbookPath`（フルパス）を指定すると、Excel未起動なら自動起動、対象ブック未オープンなら自動オープンする。Trust Centerで「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」が無効だと`ERR_VBOM_TRUST_DISABLED`（自動有効化は不可）。

対象ワークブックの状態による違い：

| 状態 | Excel起動 | ワークブックOpen | `launchedExcelPid` | 速度 |
|---|---|---|---|---|
| ① 既に開かれている | 既存を再利用 | 何もしない | 付かない | 最速 |
| ② Excel起動中・対象未オープン | 既存を再利用 | 自動でOpen | 付かない | 中速 |
| ③ Excel自体が未起動 | 新規起動（表示状態のまま） | 起動後に自動でOpen | **付く** | 最も遅い |

- `launchedExcelPid`の有無＝「このツール呼び出しでExcelを新規起動したか」の唯一の手がかり（③のみ）
- ①の場合、ユーザーの操作中の状態（編集・選択・ダイアログ表示等）がそのままAI操作に影響しうる
- `workbook`（表示名のみ）は①でしか解決できない。②③から自動で開くには`workbookPath`が必須 → **常に`workbookPath`を渡すことを推奨**
- ③で新規起動したExcelも処理後は開いたまま（自動終了しない意図的な設計）

### 4.2 複数Excelプロセスに関する注意

- `workbookPath`で自動起動したインスタンスは処理後も開いたままなので、ユーザーが別途Excelを起動する（済み含む）と2プロセス並存し、`GetActiveObject`がどちらを掴むか不定になる。Export/Importコマンドが「ブックが見つからない」等の不可解なエラーを起こすことがある → タスクマネージャーで`EXCEL.EXE`の多重起動を確認
- **サーバー内部での多重起動対策（v0.0.52で対応）**：複数のMCPツールが同時（並列）に呼ばれると、以前は各呼び出しが独立にExcelの起動状態を確認し、お互いの起動完了前に「Excelが起動していない」と誤判定して複数プロセスを起動してしまう不具合があった。現在はExcel/COMに触れる処理をサーバー内部で1つずつ順番に実行するようキュー化済み（実機で3並列呼び出し→Excelプロセス1つのみを確認済み）

## 5. マクロ実行（`excel_run_macro`）に関する注意

- `timeoutMs`（既定30秒）超過で`ERR_TIMEOUT`。ただしこれはこちら側の待ちを止めるだけで、`MsgBox`等でExcel自体が固まっている状態は解消されない → タイムアウト時はExcel画面を直接確認
- `isError`が返らない（例外なく完走した）＝意図通り動作したことの保証ではない。セル操作・ファイル出力等の結果はこのツールからは検証できない
- VBAは**プロジェクト全体を一括コンパイル**するため、どこか1モジュールにコンパイルエラーがあると無関係なマクロも全て失敗・ハングする → 原因不明のハング時はExcel側のコンパイルエラーダイアログを確認

## 6. セル値の読み取り（`excel_read_range`）

`excel_run_macro`は完走の保証のみなので、`excel_read_range`でマクロ実行後の**実際のセル値を読んで自己検証**できる（書く→実行→読んで確認→違えば直して再実行、のループ）。

- シート名（`sheet`）とセル範囲（`range`、例: `"A1:C10"`）を指定する読み取り専用ツール
- `Range.Value2`使用（`.Value`と違い日付・通貨型への暗黙変換なし）
- 行×列の2次元配列で返る（単一セルも`[[値]]`という1x1配列）
- VBAプロジェクトオブジェクトモデルへのアクセス（Trust Center設定）は**不要**（通常のExcelオブジェクトモデルのみ使用）
- 列全体・行全体（`"A:A"`等）は遅くなりうるので範囲を絞ることを推奨
- 書き込み（セルへの値の書き込み）は未実装

## 7. VBAの制御フロー解析（`vba_analyze_flow`）

生のコードを読み直して脳内でパースする代わりに、分岐・ループ・GoTo/ラベル・呼び出しを構造化JSONとして取得できる。既存の「Generate VBA Flow Chart」コマンドが使う`scripts/VBA-FlowJson.ps1`をそのまま外部プロセスとして再利用（COM経由で取得したライブなモジュールコードを一時ファイルに書き出して解析させる）。

- `procedure`省略時：モジュール内の全プロシージャの`{name, kind, startLine, endLine}`一覧のみを軽量に返す（`excel_list_macros`の`moduleName`省略パターンと同じ設計）。存在するプロシージャ名を知らない場合はまずこちらを使う
- `procedure`指定時：該当プロシージャの詳細フロー（`calls`、Mermaid用の`nodes`/`edges`/`loopSpans`）を返す
- **Phase 1の制約**：モジュール横断の呼び出し解決はしない。他モジュールへの`calls`は常に`resolved:false`になるが、これは「呼び出し先が存在しない」という意味ではなく「このツールが他モジュールを確認していない」という意味
- ディスクへの保存は一切しない（常にその場でJSON応答を返すのみ）
- `sourceHash`（正規化済みコードのSHA256）が応答に含まれる。将来の鮮度検知機能向けの下地で、現時点では比較には使われない
- 読み取り専用。VBAプロジェクトオブジェクトモデルへのアクセス（Trust Center設定）が必要

## 8. AIエージェント向けの「リファレンス」について

各ツール・各パラメータの説明文（description）がMCPプロトコル経由でAIクライアントに自動的に渡るため、これが実質的なリファレンスとして機能する。コードと説明文が常に同期する方が別ファイル保守よりズレるリスクが低いため、別途リファレンス文書は用意していない。
