# VBAフローチャート機能のMCPツール化 - 検討経緯・要件・仕様（引き継ぎ用）

別セッション（excel-vba-sync専用）でこの続きを検討・実装するための引き継ぎドキュメント。

## 背景・経緯

「誰も手をつけていないニッチなMCP/APIを作りたい」という相談から、複数のアイデアを検討・却下してきた（詳細は作業場ルートの`niche-mcp-ideas.md`参照）。その過程で「MCP Switchboard」（複数MCPを束ねて必要なツールだけ動的公開するメタMCP）を検討したが、MetaMCP・Microsoft mcp-gateway・Envoy AI Gateway・lazy-mcp等の既存OSS/エンタープライズ製品が出揃っており、かつClaude Code自身がdeferred tools+ToolSearchで同種の仕組みをネイティブ実装済みと判明し、没案とした。

その流れで「そういえばexcel-vba-syncにMermaidフローチャート生成機能を実装していた」という話が出て、これを別の切り口で活用できないか検討を始めた。

## 既存資産（実装済み・変更不要）

- `scripts/VBA-FlowJson.ps1`: `.bas`/`.cls`/`.frm`ファイルを解析し、手続き（Sub/Function/Property）ごとの制御フローをJSON化する。
  - If/ElseIf/Else/End If、Do/Loop、For/Next、Select Case/Case/End Select、With/End With、GoTo/ラベル、Exit/Return、Err.Raiseを検出し、`LoopHierarchy`/`IfHierarchy`クラスでネストを階層管理
  - フォルダ内全モジュールから`BuildSymbolTable`でシンボルテーブルを構築し、モジュール横断の呼び出し先解決（`calls`の`resolved`フラグ）を行う
  - 出力: `{ input, symbols, procedures: [{ name, kind, startLine, endLine, calls, mermaid: { nodes, edges, loopSpans } }], mermaid_global: { callgraph } }` 形式のJSON
  - CLI引数: `-FolderPath`, `-FilePath`, `-OutputPath`または`-OutputFolder`, `-Encoding`
- `scripts/Convert-FlowJsonToMermaid.ps1`: 上記JSONを読み込み、手続きごとの`.mmd`（Mermaid `flowchart TD`構文）ファイルと、モジュール横断の呼び出しグラフ`.mmd`を生成する
  - CLI引数: `-JsonPath`, `-OutDir`
- どちらも十分にテスト・調整されたロジック（ネストしたIf/Loopの分岐ラベル正規化、ループ終端の逆矢印表現など、かなり作り込まれている）。**再実装せず、既存スクリプトをそのまま外部プロセスとして再利用する方針**（`import_single_module.ps1`がVBA書き込みで確立した「既存ロジックの再利用」原則を踏襲）。

## 競合確認の結果

- **[Visustin](https://www.aivosto.com/visustin.html)**（Aivosto社）: VBA/VBScript/ASPコードをコピペするとフローチャート化する20年以上の実績がある商用デスクトップツール。Visio/Word/PowerPointへエクスポート可能。**ただしAIエージェントと非連携の単体描画ツール**。
- Mermaid描画系MCP（mermaid-mcp等）多数存在するが、いずれも「自然言語→Mermaid図」の汎用レンダリングであり、VBAコードの制御フロー解析（CFG抽出）はしていない。
- レガシーコード可視化へのAI活用（Microsoft Azure-Samples/Legacy-Modernization-Agents、IBM watsonx Code Assistant for Z等）はCOBOL/メインフレーム領域で先行しているが、Excel VBA・COMのライブ接続には非対応。

**結論**: 「VBAコード→フローチャート」という変換自体はVisustinという強力な先客がいるため単体の新製品としては厳しいが、excel-vba-syncが既に持つ「開いているExcelブックへのCOMライブ接続」と組み合わせ、かつ**人間向けの絵ではなくAIエージェントの推論材料として構造化JSONを使う**という方向性は、Visustinにも他のMCPにも真似できない独自の立ち位置になる。新規の独立プロダクトではなく、**excel-vba-syncの機能強化**として位置づける。

## 差別化の核心

- Visustin: 静的コード→絵。人間が見て理解する用途。AIとは無関係。
- 本構想: ライブComモジュール→構造化JSON→（必要なら）Mermaid。**AIエージェントがJSON構造を直接読んで「このGoToはどこに飛ぶか」「到達不能な分岐はないか」「この条件がTrueの時に何が起きるか」に答える**、または**編集前後のCFGを比較して意図しない制御フロー変化を検知する**、という用途に使う。後者はAIエージェント統合でなければ原理的に成立しない機能。

## 決定済みの設計方針

### 1. ツールは2段階に分割する（1本の巨大ツールにしない）
- `vba_analyze_flow`（コア）: モジュール＋プロシージャ名を指定し、構造化JSON（nodes/edges/loopSpans/calls）を返す
- `vba_render_flowchart`（派生、後回し可）: 同じJSONをMermaidテキストに変換する

JSONを中間形式として独立させる理由: Mermaid以外の派生用途（複雑度メトリクス算出、編集前後の差分比較、自然文要約など）に後から展開しやすいため。Mermaid直行だとこれらが「Mermaidテキストの再パース」という余計な手間になる。

### 2. 粒度はプロシージャ単位
モジュール全体を一括で返さず、`excel_get_module_code`と同様に「モジュール名＋プロシージャ名」を指定して該当分だけ返す。プロシージャ名を省略した場合は`{name, kind, startLine, endLine}`の一覧のみを軽量に返す（`excel_list_modules`/`excel_list_macros`の「一覧専用ツールを用意する」「moduleName省略時は一括取得モード」という既存の教訓と同じ設計）。

### 3. 保存はデフォルトなし、明示指定時のみ
- デフォルト: ツール応答としてその場で返すだけ（ディスクに何も残さない）
- オプション（`save: true`等）: 既存のエクスポート先規則（`extension.ts`の`resolveExportRoot()`が使う`vbaExport/<Workbook>/`配下）に`<Module>.<Proc>.flow.json`・`.mmd`を書き出す
- 保存する意味: Gitで差分管理すれば「このリファクタで制御フローがどう変わったか」をMermaidテキストのdiffで追える

### 4. 鮮度検知（将来拡張）
`ExcelUtil.ps1`の`Get-NormalizedTextHash`と同じ仕組みで、生成時点のコードのハッシュをJSONに埋め、次回呼び出し時に現在の`CodeModule`のハッシュとズレていたら鮮度切れとして警告・再生成を促す。

### 5. 差分モード（将来拡張、最大の差別化ポイント）
「このVBA修正案を適用する前後でCFGを比較し、意図しない分岐変化がないか検出する」機能。`excel_update_module_code`のdry-run/confirmTokenフローと組み合わせられると、静的ツール（Visustin）には不可能な「AIエージェントによる安全なリファクタリング支援」になる。

## 実装方針（未着手、中断時点のメモ）

- **既存の`VBA-FlowJson.ps1`/`Convert-FlowJsonToMermaid.ps1`は変更しない**。COM経由で取得したライブコードを一時ファイル（正しいモジュール名・拡張子で命名）に書き出し、既存スクリプトを外部プロセスとしてそのまま呼び出す（`import_single_module.ps1`と同じ「既存ロジックのブラックボックス再利用」方針）。
- Phase 1のスコープ: モジュール横断のシンボルテーブル解決（他モジュールへの呼び出し解決）は後回しにしてよい。一時フォルダに対象モジュール1つだけ置けば、呼び出しは`resolved:false`になるだけで、CFG自体（分岐・ループ構造）の正しさには影響しない。
- `server.ts`への追加ツール名は仮に`excel_analyze_vba_flow`（既存ツールの命名規則`excel_xxx`/`vba_xxx`と整合させるかは要検討。既存は`excel_get_module_code`, `excel_list_modules`, `excel_list_macros`, `excel_run_macro`, `vba_search_code`, `excel_update_module_code`, `excel_read_range`の7種）。
- 実装時は`.claude/skills/excel-vba-sync-dev/`の以下を必ず参照:
  - `references/encoding-rules.md`（`.ps1`/`server.ts`のエンコーディング規則。新規`.ps1`はUTF-8 BOM必須、既存`server.ts`はShift-JIS）
  - `references/com-automation-pitfalls.md`（COM多重起動レース対策の`excelOpQueue`、`psq()`によるPowerShell文字列エスケープ、`Write-Host`のstdout汚染対策等）
  - `references/mcp-server-ops.md`（変更反映には`dist-server`の子プロセスを手動kill、環境変数を3経路で同期させる必要がある等）
- 未確認事項: `VBA-FlowJson.ps1`・`Convert-FlowJsonToMermaid.ps1`自体の実ファイルエンコーディング（BOM有無）を`Get-Content -Encoding Byte -TotalCount 4`等で確認してから着手すること（既存のSJIS/UTF8BOM分類リストに明記がないため）。

## 未解決の設計論点

- ツール名の最終決定（`excel_analyze_vba_flow` vs `vba_analyze_flow`など、既存命名規則との整合）
- 一時ファイル・一時JSONの配置場所（OSの一時フォルダで良いか、後片付けのタイミング）
- 保存機能のパラメータ名・デフォルト値
- 差分モードのインターフェース（2つのコード文字列を渡す？ dry-run結果と組み合わせる？）
