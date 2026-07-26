[← README に戻る](../README.md)

# 🤖 AIクライアントから使う（Claude Code / Claude Desktop） / Use from an AI client

この拡張機能には[Model Context Protocol (MCP)](https://modelcontextprotocol.io/)サーバーが内蔵されており、Claude Code・Claude Desktopなどの外部AIクライアントから、VBAコードの読み取り・検索・マクロ実行に加えて、**AIが書いたコードをExcelのVBAモジュールへ直接書き込む**ことができます（`excel_update_module_code`ツール）。

**セットアップ手順**
1. VS Codeでこの拡張機能をインストールした状態で、コマンドパレット（`Ctrl+Shift+P`）から **`Excel VBA: Print MCP Server Config (for AI)`** を実行する
2. 生成されたJSON（クリップボードにコピー済み）を、Claude Codeなら`.mcp.json`、Claude Desktopなら`claude_desktop_config.json`の`mcpServers`に貼り付ける
3. AIクライアントを再起動すると、`ping` / `excel_get_module_code` / `excel_list_macros` / `excel_run_macro` / `vba_search_code` / `excel_update_module_code`の6ツールが使えるようになる

このMCPサーバー（`dist-server/server.js`）はNode単体で動作するため、**VS Codeを開いていなくても**AIクライアントから直接起動・利用できます（ただしExcel本体とVBA実行環境はWindows上に必要です）。

**VS Code自身のCopilot Chat（エージェントモード）からの利用**
上記のセットアップ手順は外部AIクライアント（Claude Code/Desktopなど）向けです。VS Code自体に組み込まれたAIクライアント（Copilot Chatのエージェントモード等）向けには、`contributes.mcpServerDefinitionProviders`経由でこの拡張機能自身がMCPサーバーを自動的に登録します。手動での`.mcp.json`設定は不要で、拡張機能をインストールするだけで検出されます（起動される内容は上記の手動設定と同一）。`onStartupFinished`のような明示的な起動イベントを追加しなくても自動的に有効化されることを、Copilot Chatのツール選択画面で実機確認済みです。

Copilot Chatのツール選択画面には`vba-excel-mcp`という名前（サーバー自身がMCPプロトコルで名乗る識別子）で表示されます。あわせて人間向けの表示名（`title: "Excel VBA Sync"`）と、サーバーの使い方を要約した`instructions`（`workbookPath`を優先すること、`dryRun`/`confirmToken`の書き込みフロー等）も設定していますが、これらをクライアントのUIがどう表示するかはクライアント次第です。

**複数クライアント同時利用時の注意**
Claude Code/Desktop（手動設定）とVS Code内蔵のAIクライアント（自動検出）を同時に使う場合、それぞれが独立した`server.js`プロセスを起動します（MCPの標準的な設計であり不具合ではありません）。同時に複数のクライアントがExcel操作を行うと、上記「複数Excelプロセスに関する注意」のリスクが顕在化しやすくなる点にご留意ください。

**Excel未起動でも自動解決**
`excel_get_module_code` / `excel_list_macros` / `excel_run_macro` / `vba_search_code` / `excel_update_module_code`はいずれも`workbookPath`（ワークブックのフルパス）を指定できます。指定した場合、Excelが起動していなければ自動的に起動し、対象ブックが未オープンであれば自動的に開いてから操作します。Excelのトラストセンターで「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」が有効になっている必要があり、無効な場合は`ERR_VBOM_TRUST_DISABLED`という分かりやすいエラーが返ります（この設定はプログラムから自動的に有効化することはできません）。

**書き込み（`excel_update_module_code`）の安全設計**
- `dryRun:true`でまず呼び出すと、実際には書き込まず、現在のコードとの差分プレビュー＋`confirmToken`が返る
- 同じ`confirmToken`を付けて再度呼び出すことで、初めて実際にExcelへ書き込まれる（AIが誤って一発で書き込んでしまうことを防ぐ2段階フロー）
- `confirmToken`はdry-run時点のモジュールのコードに紐付いており、書き込み直前にモジュールを再読込してトークンを再計算・照合する（楽観的排他制御）。dry-runから確定までの間に**別のクライアント（別のAI/別セッション等）が同じモジュールを先に書き換えていた場合**、トークンが一致せず`ERR_MODULE_CHANGED_SINCE_DRYRUN`で書き込みが拒否される（気づかずに上書きすることを防ぐ。実機で2クライアントの競合を再現し、後勝ちの上書きが正しく拒否されることを確認済み）
- 書き込み前には自動的に現在のコードがワークブックと同じフォルダの`.excel-vba-sync-backups`にタイムスタンプ付きで退避される
- Sheet/ThisWorkbookのコードビハインド（Documentモジュール）に書き込んだ場合、ショートカットキー割り当てなどのプロシージャ単位のAttribute情報はVBAのAPI制約上失われる（レスポンスに警告が入る）

**マクロ実行（`excel_run_macro`）に関する注意**
- `timeoutMs`（既定30秒）を超えると呼び出しは`ERR_TIMEOUT`で打ち切られるが、これは**こちら側の待ちを止めるだけ**で、`MsgBox`・`InputBox`・フォーム表示等でExcel自体がダイアログ待ちで固まっている状態は解消されない。タイムアウトが出た場合はExcel画面を直接確認すること
- ツールが`isError`を返さなかった（＝例外を投げずに完走した）からといって、マクロが**意図した通りに**動作したことは保証されない。セル操作・ファイル出力・イミディエイトウィンドウ出力などの結果は、このMCPサーバーからは直接検証できないため、実行後に何を確認すべきかは呼び出し側が判断する必要がある
- VBAは**プロジェクト全体を一括コンパイル**する仕様のため、ワークブック内のどこか1モジュールにでもコンパイルエラー（`Option Explicit`下での未宣言変数など）があると、**それとは無関係なマクロを呼んでも全て失敗・ハングする**。呼び出しが原因不明にハングしたりエラーになったりした場合は、まずExcel側にコンパイルエラーのダイアログが出ていないか確認すること

**複数Excelプロセスに関する注意**
`workbookPath`指定でExcelが未起動の場合に自動起動されたインスタンスは、処理後も**開いたままになります**（意図的な設計）。この状態でユーザーが別途Excelを手動で起動する（または既に起動していた）と、Excelプロセスが2つ並存し、`GetActiveObject`がどちらを掴むか不定になり、VS Code UIのExport/Importコマンドなどが「保存済みのブックが見つからない」といった不可解なエラーを起こすことがあります。これを見分けやすくするため、ツール呼び出し中に**新規でExcelを自動起動した場合のみ**、レスポンスの`launchedExcelPid`にそのプロセスIDが入ります（既存のExcelを再利用した場合は含まれません）。心当たりのないエラーが出た場合は、タスクマネージャーで`EXCEL.EXE`が複数起動していないか確認してください。

**AIエージェント向けの「リファレンス」について**
このMCPサーバーの各ツール・各パラメータには説明文（description）を実装済みで、MCPプロトコル経由でAIクライアントに自動的に渡ります。これが実質的なリファレンスとして機能するため、別途リファレンス文書は用意していません（コードと説明文が常に同期している方が、別ファイルを保守してズレるリスクより信頼できるため）。
