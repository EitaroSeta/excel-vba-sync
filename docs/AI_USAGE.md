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
3. AIクライアントを再起動 → 12ツール（`ping` / `excel_list_modules` / `excel_get_module_code` / `excel_list_macros` / `excel_run_macro` / `vba_search_code` / `vba_analyze_flow` / `vba_render_flowchart` / `vba_list_dependencies` / `vba_list_references` / `excel_update_module_code` / `excel_read_range`）が使用可能に

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
- モジュール横断の呼び出し解決に対応済み：ワークブック内の全モジュールを一時フォルダに並べてから解析するため、他モジュールで定義された関数への呼び出しも`resolved:true`になる。`resolved:false`は「ワークブック内のどこにも呼び出し先が見つからなかった」ことを意味する
- ディスクへの保存は一切しない（常にその場でJSON応答を返すのみ）
- `sourceHash`（正規化済みコードのSHA256）が応答に含まれる。将来の鮮度検知機能向けの下地で、現時点では比較には使われない
- 読み取り専用。VBAプロジェクトオブジェクトモデルへのアクセス（Trust Center設定）が必要

`vba_analyze_flow`のJSONを人間が見て分かる図にしたい場合は`vba_render_flowchart`を使う。既存の「Generate VBA Flow Chart」コマンドと同じ`Convert-FlowJsonToMermaid.ps1`をそのまま再利用し、Mermaidの`flowchart TD`テキストを返す。

- `procedure`指定時：そのプロシージャの詳細フローチャート
- `procedure`省略時：モジュール内の全プロシージャの呼び出し関係を示すコールグラフ（`vba_analyze_flow`と同様、他モジュールへの呼び出しも解決される）
- こちらもディスクには一切保存しない（一時フォルダで生成→読み取り→即削除）。返ってきたMermaidテキストをMarkdownプレビューやmermaid.live等に貼り付けて図として見る

## 8. 外部・プラットフォーム依存のリストアップ（`vba_list_dependencies`）

VBAコード内の「Excel外部への依存」を、簡易な正規表現マッチで一覧化する。`excel_update_module_code`の`lintWarnings`と同程度の厳密さ（簡易テキストマッチ、本物のVBAパーサーではない）。

- 検出対象：Windows API宣言（`Declare Sub`/`Declare Function`）、`CreateObject`/`GetObject`によるCOM自動化呼び出し、VBA組み込みの`Shell`関数呼び出し、`Application.Run`（他ブックの文字列指定を含む動的マクロ呼び出し）、VBAネイティブのファイルI/O（`Open`/`Kill`/`FileCopy`/`MkDir`/`RmDir`）、`Scripting.FileSystemObject`のファイル/フォルダ操作メソッド（`OpenTextFile`/`CreateTextFile`/`CopyFile`/`DeleteFile`/`MoveFile`/`CreateFolder`/`DeleteFolder`/`MoveFolder`）、`Workbooks.Open`（外部ブック参照）
- ProgID・呼び出し先名・参照先パスが文字列リテラルなら抽出、変数指定なら`dynamic:true`
- FSOのメソッド検出はメソッド名だけで判定している（呼び出し元の変数が本当にFSOのインスタンスかは確認していない）。`fileIo`の各件に付く`methodNameOnly:true`がこのケースの目印で、同名メソッドを持つ別のオブジェクトも拾ってしまう可能性がある
- 行全体がコメント（`'`または`Rem`で始まる）の場合はスキップされる。ただしコード行末尾のインラインコメント（`Kill "x" ' メモ`等）は文字列リテラルとの区別が難しいため除外していない
- `Application.Run`の検出は`vba_analyze_flow`を補完する：呼び出し元が見つからないプロシージャは、実は`Application.Run`で名前指定されて動的に呼ばれているだけかもしれない。「未使用」と断定する前にこちらも確認するとよい
- `module`省略時：ワークブック内の全モジュールを1回のCOMセッションでスキャン
- 検出0件のモジュールはレスポンスから省略される
- 読み取り専用。Office Scripts等への移行検討時に「VBA外で何をしているか」「他にどのファイルに依存しているか」を洗い出す用途を想定

## 9. イベントプロシージャ・シート/名前定義参照のリストアップ（`vba_list_references`）

「このコードは何をきっかけに動くか」「このシート・名前付き範囲を変更/削除したら何が壊れるか」という影響調査向け。`vba_list_dependencies`と同じ正規表現ベース・読み取り専用のアプローチだが、対象が「外部依存」ではなく「ワークブック内部への参照」である点が異なるため別ツールにしている。

- 検出対象：イベントプロシージャ（`Workbook_*`/`Worksheet_*`/`UserForm_*`。VBA自身の命名規則をそのまま使うので、個別のイベント名一覧を持つ必要がない。加えてレガシーな`Auto_Open`/`Auto_Close`）、シート参照（`Worksheets("名前")`/`Sheets("名前")`）、名前付き範囲の可能性がある参照（`Range("...")`/`Names("...")`）
- 埋め込みActiveXコントロールのイベント（`CommandButton1_Click`等）は検出しない（コントロール名の把握にDesigner/OLEObjectsの読み取りが必要で、このツールのスコープ外）
- `Range("...")`はセルアドレスそのもの（`A1`、`B2:C10`、`A:A`等）に見える場合は除外する。実際のコードでは大半のRange呼び出しが単なるセル参照であり、含めると名前付き範囲の兆候が埋もれてしまうため。変数指定の動的な`Range(変数)`も同様の理由で対象外（`Names(...)`は逆に、Namesコレクションを使っている時点で名前付き範囲がらみと言えるため動的指定でも検出する）
- `Range("プレフィックス" & 行番号)`のような文字列連結によるセル範囲構築は、閉じ引用符の直後が`)`または`,`でない（＝リテラルが引数全体ではない）場合は除外される
- イベントプロシージャの検出は`vba_analyze_flow`・`vba_list_dependencies`の`Application.Run`検出と合わせて考えるとよい：呼び出し元が見つからないプロシージャも、①イベントプロシージャである、②`Application.Run`で動的に呼ばれている、のどちらかであれば未使用ではない
- `module`省略時：ワークブック内の全モジュールを1回のCOMセッションでスキャン。検出0件のモジュールはレスポンスから省略される

## 10. AIエージェント向けの「リファレンス」について

各ツール・各パラメータの説明文（description）がMCPプロトコル経由でAIクライアントに自動的に渡るため、これが実質的なリファレンスとして機能する。コードと説明文が常に同期する方が別ファイル保守よりズレるリスクが低いため、別途リファレンス文書は用意していない。
