<#
    File: export_opened_vba.ps1
    Description: EXCELのVBAモジュールからフローチャートJSONを作成するスクリプト
    Author: Eitaro SETA
    License: MIT License
    Copyright (c) 2025 Eitaro SETA

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.
#>

<# 
VBA-FlowJson.ps1  (minimal & robust)
- 指定された文字コードの .bas/.cls/.frm を解析し、対象1ファイルの手続きごとに
  Mermaidフローチャート用の nodes/edges と、関数呼び出し(calls)をJSON出力。
- 依存の強い構文は排し、.Count を一切使わない（すべて .Length か ForEach に統一）
- Encodingパラメータでファイル文字コードを指定可能（デフォルト: UTF8）

使用例:
  # UTF-8で読み込み（デフォルト）
  pwsh -File "VBA-FlowJson.ps1" "c:\folder" "c:\folder\file.bas" "output.json"
  
  # Shift_JISで読み込み  
  pwsh -File "VBA-FlowJson.ps1" "c:\folder" "c:\folder\file.bas" "output.json" -Encoding "Shift_JIS"
  
  # その他の文字コード例: Default, ASCII, Unicode, UTF32, etc.
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)][string]$FolderPath,
  [Parameter(Mandatory=$true)][string]$FilePath,
  [Parameter()][string]$OutputPath,
  [Parameter()][string]$OutputFolder,
  [Parameter()][string]$Encoding = "UTF8"
)

# ロケール取得
$locale = (Get-UICulture).Name.Split('-')[0]
$defaultLocale = "en"
$localePath = Join-Path $PSScriptRoot "..\locales\$locale.json"

if (-Not (Test-Path $localePath)) {
    # 指定ロケールがなければ英語にフォールバック
    $localePath = Join-Path $PSScriptRoot "..\locales\$defaultLocale.json"
}

$messages = Get-Content $localePath -Encoding UTF8 -Raw | ConvertFrom-Json

# UTF-8で出力されるよう設定
$OutputEncoding = [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$InformationPreference = 'Continue'

function Get-Timestamp {
    Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
}

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# UTF-8 BOMなしでファイルに出力する関数
function Out-FileUtf8NoBom {
  param(
    [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
    [string]$Content,
    [Parameter(Mandatory=$true)]
    [string]$FilePath
  )
  
  $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
  [System.IO.File]::WriteAllText($FilePath, $Content, $utf8NoBom)
}

# ユーティリティ
# 配列化ユーティリティ
function Ensure-Array { param($x) if ($null -eq $x) { return @() }; if ($x -is [System.Collections.IEnumerable] -and -not ($x -is [string])) { return @($x) } return @($x) }

# 指定された文字コードでのファイル読み込み
function ReadFileLines { 
  param([string]$Path, [string]$FileEncoding = "UTF8") 
  if(-not (Test-Path -LiteralPath $Path -PathType Leaf)){ 
    throw "File not found: $Path" 
  } 
  return @(Get-Content -LiteralPath $Path -Encoding $FileEncoding) 
}

# UTF-8でのファイル読み込み（後方互換性のため保持）
function ReadUtf8Lines { param([string]$Path) return ReadFileLines -Path $Path -FileEncoding "UTF8" }

#コメント除去
function StripComments {
  param([string]$Line)
  if([string]::IsNullOrEmpty($Line)){ return "" }

  # 先頭が ' または Rem で始まる行はコメント行として空文字を返す
  if($Line -match '^\s*(?:''|Rem\s)'){ return "" }

  # 行内コメントの除去（文字列リテラル内の ' は除外）
  $line = [regex]::Replace($line, "(?<![""'])(?:'|Rem\s).*", "")

  # 文字列リテラル内の ' を保護しつつコメント除去
  $parts = $Line -split '(")'
  for($i=0; $i -lt $parts.Length; $i+=2){
    $p = $parts[$i]; $idx = $p.IndexOf("'"); if($idx -ge 0){ $parts[$i] = $p.Substring(0,$idx) }
  }
  return ($parts -join "")
}

# 継続行結合（行末の _ を結合）
function JoinContinuations {
  param([string[]]$Lines)
  $Lines = @(Ensure-Array $Lines)
  if($Lines.Length -eq 0){ return @() }
  $buf = New-Object System.Collections.Generic.List[string]
  $acc = ""
  foreach($ln in $Lines){
    if($null -eq $ln){ continue }
    $trim = $ln.TrimEnd()
    if($trim -match '\s+_\s*(?:''.*)?$'){
      $acc += ($trim -replace '\s+_\s*(?:''.*)?$',' ')
    } else {
      if($acc){ $buf.Add($acc + $trim); $acc = "" } else { $buf.Add($ln) }
    }
  }
  if($acc){ $buf.Add($acc) }
  return $buf.ToArray()
}

# モジュール名取得 
function GetModuleName {
  param([string[]]$Lines,[string]$Fallback)
  $Lines = @(Ensure-Array $Lines)
  $max = [math]::Min(200, [math]::Max(0, $Lines.Length-1))
  for($i=0; $i -le $max; $i++){
    $ln = [string]$Lines[$i]
    if($ln -match '^\s*Attribute\s+VB_Name\s*=\s*"(?<n>[^"]+)"'){ return $Matches.n }
  }
  return [IO.Path]::GetFileNameWithoutExtension($Fallback)
}

# Edge存在チェック
function HasEdge {
  param(
    [Parameter(Mandatory)]$from,
    [Parameter(Mandatory)]$to,
    [string]$label
  )
  foreach($e in $edges){
    if( [string]$e.from -eq [string]$from -and [string]$e.to -eq [string]$to ){
      if([string]::IsNullOrEmpty($label) -or [string]$e.label -eq $label){ return $true }
    }
  }
  return $false
}

# ユーティリティ：特定 from/to のエッジを一括削除（List なので後ろから消す）
function Remove-Edges {
  param([string]$from,[string]$to)
  for($i = $edges.Count-1; $i -ge 0; $i--){
    $e = $edges[$i]
    if(([string]$e.from -eq [string]$from) -and ([string]$e.to -eq [string]$to)){
      $edges.RemoveAt($i)
    }
  }
}

# ---- シンボルテーブル構築（フォルダ内の全モジュール） ----
function BuildSymbolTable {
  param([string]$Folder, [string]$FileEncoding = "UTF8")
  if(-not (Test-Path -LiteralPath $Folder -PathType Container)){ throw "Folder not found: $Folder" }
  $table = @{}
  $files = @(Get-ChildItem -LiteralPath $Folder -File -Recurse -Include *.bas,*.cls,*.frm)
  foreach($f in $files){
    $lines  = @(JoinContinuations (ReadFileLines -Path $f.FullName -FileEncoding $FileEncoding))
    $module = GetModuleName -Lines $lines -Fallback $f.Name
    if(-not $table.ContainsKey($module)){ $table[$module] = New-Object System.Collections.Generic.HashSet[string] }
    foreach($ln in $lines){
      #if($ln -match '^\s*(?:(?:Public|Private|Friend)\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+(?<name>[A-Za-z_][A-Za-z0-9_]*)\s*\('){
      if($ln -match '^\s*(?:(?:Public|Private|Friend)\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+(?<name>[\p{L}_][\p{L}\p{N}_]*)\s*\('){
        [void]$table[$module].Add($Matches.name)
      }
    }
  }
  $out = @{}; foreach($k in $table.Keys){ $out[$k] = @($table[$k]) }
  return $out
}

# 手続き抽出（対象ファイル）
function ParseProcedures {
  param([string[]]$Lines)
  $Lines = @(Ensure-Array $Lines)
  #$rxHead = '^\s*(?:(?:Public|Private|Friend)\s+)?(?<kind>Sub|Function|Property\s+(?:Get|Let|Set))\s+(?<name>[A-Za-z_][A-Za-z0-9_]*)\s*\('
  $rxHead = '^\s*(?:(?:Public|Private|Friend)\s+)?(?<kind>Sub|Function|Property\s+(?:Get|Let|Set))\s+(?<name>[\p{L}_][\p{L}\p{N}_]*)\s*\('
  $rxEnd  = '^\s*End\s+(Sub|Function|Property)\b'
  $list = New-Object System.Collections.Generic.List[pscustomobject]
  $in = $false; $cur = $null
  for($i=0; $i -lt $Lines.Length; $i++){
    $line = [string]$Lines[$i]
    if(-not $in -and $line -match $rxHead){
      $in = $true
      $cur = [pscustomobject]@{
        name = $Matches.name
        kind = ($Matches.kind -replace '\s+',' ')
        startLine = $i+1
        endLine = $null
        body = New-Object System.Collections.Generic.List[string]
      }
      continue
    }
    if($in -and $line -match $rxEnd){
      $cur.endLine = $i+1
      $list.Add($cur)
      $in = $false; $cur = $null
      continue
    }
    if($in){ $cur.body.Add($line) }
  }
  return $list.ToArray() # ← 配列保証（Lengthあり）
}

# 手続き解析：ノード/エッジ/呼び出し
function AnalyzeProcedure {
  param(
    [string[]]$BodyLines,
    [int]$StartLineOffset,
    [string]$ThisModule,
    [hashtable]$SymbolTable
  )
  $BodyLines = @(Ensure-Array $BodyLines)
  $usedIds = New-Object 'System.Collections.Generic.HashSet[string]'   # 使ったID集合
  $loopSpans = New-Object System.Collections.Generic.List[object]

  # kind が一致するまで上から Pop して返す。見つからなければ $null
  function Pop-UntilKind {
    param($stack, [string]$kind)
    while ($stack.Count -gt 0) {
      $x = $stack.Pop()
      if ($null -ne $x -and ($x.PSObject.Properties.Name -contains 'kind') -and $x.kind -eq $kind) {
        return $x
      }
    }
    return $null
  }

  # スタックを壊さず、最も近い一致 kind を上から探して返す（見つからなければ $null）
  function Peek-NearestKind {
    param($stack, [string]$kind)
    foreach ($x in $stack.ToArray()) { # 先頭がトップ
      if ($null -ne $x -and ($x.PSObject.Properties.Name -contains 'kind') -and $x.kind -eq $kind) {
        return $x
      }
    }
    return $null
  }

  # ノード/エッジ/呼び出しリスト
  $nodes = New-Object System.Collections.Generic.List[pscustomobject]
  $edges = New-Object System.Collections.Generic.List[pscustomobject]
  $calls = New-Object System.Collections.Generic.List[pscustomobject]

  # === 階層管理システム ===
  # ループ階層管理クラス
  class LoopHierarchy {
    [System.Collections.Generic.Stack[object]]$loopStack
    [int]$nextLoopId
    
    LoopHierarchy() {
      $this.loopStack = New-Object System.Collections.Generic.Stack[object]
      $this.nextLoopId = 1
    }
    
    # 新しいループを開始
    [object] StartLoop([string]$type, [string]$headNodeId, [int]$line) {
      $loopInfo = [pscustomobject]@{
        id = $this.nextLoopId++
        type = $type
        headNodeId = $headNodeId
        line = $line
        level = $this.loopStack.Count
        endNodeId = $null
        innerIds = New-Object System.Collections.Generic.List[string]
      }
      $this.loopStack.Push($loopInfo)
      return $loopInfo
    }
    
    # 現在のループを終了
    [object] EndLoop([string]$endNodeId) {
      if ($this.loopStack.Count -eq 0) { return $null }
      $loopInfo = $this.loopStack.Pop()
      $loopInfo.endNodeId = $endNodeId
      return $loopInfo
    }
    
    # 現在のループレベル取得
    [int] GetCurrentLevel() {
      return $this.loopStack.Count
    }
    
    # 現在のループ情報取得
    [object] GetCurrentLoop() {
      if ($this.loopStack.Count -eq 0) { return $null }
      return $this.loopStack.Peek()
    }
    
    # Exit For処理用：指定レベルのループを検索
    [object] FindLoopAtLevel([int]$level) {
      $stackArray = $this.loopStack.ToArray()
      for ($i = 0; $i -lt $stackArray.Length; $i++) {
        if ($stackArray[$i].level -eq $level) {
          return $stackArray[$i]
        }
      }
      return $null
    }
  }

  # If階層管理クラス
  class IfHierarchy {
    [System.Collections.Generic.Stack[object]]$ifStack
    [int]$nextIfId
    
    IfHierarchy() {
      $this.ifStack = New-Object System.Collections.Generic.Stack[object]
      $this.nextIfId = 1
    }
    
    # 新しいIfブロックを開始
    [object] StartIf([string]$condNodeId, [int]$line) {
      $ifInfo = [pscustomobject]@{
        id = $this.nextIfId++
        condNodeId = $condNodeId
        line = $line
        level = $this.ifStack.Count
        hasElse = $false
        elseIfCount = 0
        joinNodeId = $null
      }
      $this.ifStack.Push($ifInfo)
      return $ifInfo
    }
    
    # ElseIf追加
    [void] AddElseIf() {
      if ($this.ifStack.Count -eq 0) { return }
      $ifInfo = $this.ifStack.Peek()
      $ifInfo.elseIfCount++
    }
    
    # Else追加
    [void] AddElse() {
      if ($this.ifStack.Count -eq 0) { return }
      $ifInfo = $this.ifStack.Peek()
      $ifInfo.hasElse = $true
    }
    
    # Ifブロック終了
    [object] EndIf([string]$joinNodeId) {
      if ($this.ifStack.Count -eq 0) { return $null }
      $ifInfo = $this.ifStack.Pop()
      $ifInfo.joinNodeId = $joinNodeId
      return $ifInfo
    }
    
    # 現在のIfレベル取得
    [int] GetCurrentLevel() {
      return $this.ifStack.Count
    }
    
    # 現在のIf情報取得
    [object] GetCurrentIf() {
      if ($this.ifStack.Count -eq 0) { return $null }
      return $this.ifStack.Peek()
    }
  }

  # 階層管理インスタンス作成
  $loopHierarchy = [LoopHierarchy]::new()
  $ifHierarchy = [IfHierarchy]::new()

  # ノード追加ユーティリティ
  $AddNode = {
    param([string]$type,[string]$text,[int]$line,[string]$comment=$null)

    # 文字化け修正 - ElseIfと区別するためより具体的に修正
    $cleanText = $text
    if ($cleanText -eq 'Else?.*' -or $cleanText -match '^Else\s*$') {
      $cleanText = 'Else'
    }

    #$id = "n$line"
    # 基本IDは行番号ベース（互換性維持）
    $baseId = "n{0}" -f $line
    $id = $baseId
    $idx = 1
    while ($usedIds.Contains($id)) {
      $idx++
      $id = "{0}_{1}" -f $baseId, $idx   # n43_2, n43_3 ...
    }
    $usedIds.Add($id) | Out-Null

    $nodes.Add([pscustomobject]@{
      id=$id; type=$type; text=([string]$cleanText).Trim(); line=$line; comment=$comment
    })

    return $id
  }
  # エッジ追加ユーティリティ
  $AddEdge = {
    param([string]$from,[string]$to,[string]$label="")
    if ([string]::IsNullOrWhiteSpace([string]$from)) { return }
    if ([string]::IsNullOrWhiteSpace([string]$to))   { return }
    $edges.Add([pscustomobject]@{ from=$from; to=$to; label=$label })| Out-Null
  }

  # ラベル位置マップ作成
  $labels = @{}
  $labelNodeMap = @{}     # ラベル → ノードID
  for($i=0; $i -lt $BodyLines.Length; $i++){
    $code = StripComments $BodyLines[$i]
    if($code -match '^\s*([A-Za-z_][A-Za-z0-9_]*)\s*:\s*$'){ $labels[$Matches[1].ToLower()] = $i }
  }

  # 解析開始
  $prevId = & $AddNode 'start' 'Start' $StartLineOffset
  $stack = New-Object System.Collections.Stack

  # 呼び出し検出用正規表現
  $rxCall1 = '(?i)\bCall\s+([A-Za-z_][A-Za-z0-9_]*)(?:\s*\()'
  $rxCall2 = '(?i)\b([A-Za-z_][A-Za-z0-9_]*)\.([A-Za-z_][A-Za-z0-9_]*)\s*\('
  $rxCall3 = '(?i)\b([A-Za-z_][A-Za-z0-9_]*)\s*\('

  # 各行解析
  for($i=0; $i -lt $BodyLines.Length; $i++){
    $lineNum = $StartLineOffset + $i
    $trim = (StripComments $BodyLines[$i]).Trim()

    # StripComments で既に先頭コメント行は空文字になっているので、空行チェックで除外される
    if([string]::IsNullOrWhiteSpace($trim)){ continue }
    if($trim -match '^\s*[A-Za-z_][A-Za-z0-9_]*\s*:\s*$'){ continue }

    # 1行 If:  If ... Then <stmt>   （行末に End If が無い、":" での続きも無しのケース）
    if ($trim -match '^\s*If\s+(.+?)\s+Then\s+(.+?)\s*$') {
      $condExpr = $Matches[1].Trim()
      $thenStmt = $Matches[2].Trim()

      # 条件ノード
      $cid = & $AddNode 'cond' ("If " + $condExpr + "?") $lineNum
      & $AddEdge $prevId $cid

      # Then 側ノード（Exit など中身は通常の分類でOK。ここでは 'end' に寄せず 'op' としておいても可）
      # Exit/Return だけは Yes で接続し、prev を更新しない
      if ($thenStmt -match '^\s*(Exit\s+(Sub|Function|Property)|Return)\b') {
        $eid = & $AddNode 'end' $thenStmt $lineNum
        & $AddEdge $cid $eid 'Yes'
        # Exit は終端：prev は更新しない
      } else {
        $tid = & $AddNode 'op' $thenStmt $lineNum
        & $AddEdge $cid $tid 'Yes'
        # Then ノードから続けたいなら：$prevId = $tid
      }

      # 即席の合流点を作成：No は必ずこちらへ
      $joinId = & $AddNode 'join' 'End If' $lineNum
      & $AddEdge $cid $joinId 'No'

      # Then 側が Exit でない場合は Then ノードからも合流へ
      if ($PSBoundParameters.ContainsKey('tid')) { & $AddEdge $tid $joinId }

      # 以降の連続処理は合流点から
      $prevId = $joinId
      continue
    }
    
    # If/ElseIf/Else/End If の検出とstack処理（階層管理対応）
    if($trim -match '^\s*If\s+(.+)\s+Then\s*(?:$|:)'){
      $cid = & $AddNode 'cond' ("If " + $Matches[1].Trim() + "?") $lineNum
      & $AddEdge $prevId $cid
      
      # 階層管理システムにIf登録
      $ifInfo = $ifHierarchy.StartIf($cid, $lineNum)
      
      # ★ ends を持つフレームで管理
      $stack.Push([pscustomobject]@{
        kind='If'; cond=$cid; hasElse=$false; line=$lineNum;
        ends = New-Object System.Collections.Generic.List[string]
        needsYes = $true;
        needsNo = $true;  # ★ Else処理のために最初はtrueに設定
        needsElse = $false;
        hierarchyId=$ifInfo.id
     })
      $prevId = $cid; continue
    }

# ElseIf の検出とstack処理（階層管理対応）
if($trim -match '^\s*ElseIf\s+(.+)\s+Then\s*$'){
  if($stack.Count -gt 0 -and $stack.Peek().kind -eq 'If'){
    $top = $stack.Pop()
    # 直前ブランチの最終ノードを ends に退避
    if($null -ne $prevId){ $top.ends.Add($prevId) }

    # 階層管理システムでElseIf追加
    $ifHierarchy.AddElseIf()

    $eid = & $AddNode 'cond' ("ElseIf " + $Matches[1].Trim() + "?") $lineNum
    # 元の条件ノードから ElseIf 条件への接続は 'No' で行う
    & $AddEdge $top.cond $eid 'No'
    $stack.Push([pscustomobject]@{
      kind='If';
      cond=$eid;
      hasElse=$top.hasElse;
      line=$top.line;
      ends=$top.ends;
      needsYes=$true;
      needsNo=$true;
      needsElse=$false;
      hierarchyId=$top.hierarchyId
    })
    # 元の条件ノードのneedsNoを無効化（ElseIfが処理済み）
    $top.needsNo = $false
    $prevId = $eid
  } 
  continue
}
    # Else (If のみ) - 階層管理対応
    if($trim -match '^\s*Else\s*$'){
      if($stack.Count -gt 0 -and $stack.Peek().kind -eq 'If'){
        $top = $stack.Peek()
        # ★ 直前ブランチの最終ノードを ends に退避
        if($null -ne $prevId){ $top.ends.Add($prevId) }
        
        # 階層管理システムでElse追加
        $ifHierarchy.AddElse()
        
        if($top.needsNo){
          $elseNodeId = & $AddNode 'op' 'Else' $lineNum
          & $AddEdge $top.cond $elseNodeId 'No'
          $top.needsNo = $false
          $top.hasElse = $true
          $prevId = $elseNodeId  # ★ Else処理ノードに続く処理のためにprevIdを更新
        }
      } 
      continue
    }

    # If文の最終処理（End If）- 階層管理対応
    if($trim -match '^\s*End\s+If\s*$'){
      if($stack.Count -gt 0 -and $stack.Peek().kind -eq 'If'){
        $top = $stack.Pop()
        
        # 階層管理システムでIf終了処理
        $joinId = & $AddNode 'join' 'End If' $lineNum
        $ifInfo = $ifHierarchy.EndIf($joinId)
    
        # 現在位置も ends に追加
        if($null -ne $prevId){ $top.ends.Add($prevId) }
        
        # 全ての分岐の終端を合流点へ接続
        foreach($endId in $top.ends){
          & $AddEdge $endId $joinId
        }
        
        # Else が無い場合は条件ノードから直接合流点への No エッジを追加
        if(-not $top.hasElse -and $top.needsNo){
          & $AddEdge $top.cond $joinId 'No'
        }
        
        $prevId = $joinId
      }
      continue
    }

    # Do/Loop の検出とstack処理（階層管理対応）
    if($trim -match '^\s*(Do\s+While|Do\s+Until|Do\b)'){
      $did = & $AddNode 'loop' ($trim.Split(':')[0]) $lineNum
      & $AddEdge $prevId $did
      
      # 階層管理システムにLoop登録
      $loopInfo = $loopHierarchy.StartLoop('Do', $did, $lineNum)
      
      $stack.Push([pscustomobject]@{ 
        kind='Do'; 
        head=$did; 
        line=$lineNum;
        hierarchyId=$loopInfo.id
      })
      $prevId = $did; continue
    }
    
    # Loop のとき：Popし、Doまでを捨てる（階層管理対応）
    if($trim -match '^\s*Loop\b'){
      $top = Pop-UntilKind $stack 'Do'
      if ($null -ne $top -and ($top.PSObject.Properties.Name -contains 'head')) {
        $eid = & $AddNode 'loopEnd' 'Loop End' $lineNum
        & $AddEdge $prevId $eid
        
        # 階層管理システムでLoop終了処理
        $loopInfo = $loopHierarchy.EndLoop($eid)

        # ループスパンを登録（階層ID付き）
        $span = [pscustomobject]@{
          headId = $top.head -replace '^n', ''
          endId = $eid -replace '^n', ''
          start = [int]$top.line
          end = [int]$lineNum
          hierarchyId = if($loopInfo) { $loopInfo.id } else { 0 }
          level = if($loopInfo) { $loopInfo.level } else { 0 }
        }
        $loopSpans.Add($span)
        Write-Verbose "Added hierarchical Do-Loop span: headId=$($span.headId), endId=$($span.endId), level=$($span.level), hierarchyId=$($span.hierarchyId)"

        $prevId = $eid
      } else {
        Write-Verbose "Skip Loop at line $lineNum (no matching Do on stack)"
      }
      continue
    }

    # For/Next の検出とstack処理（階層管理対応）
    if($trim -match '^\s*For\s+(.+)'){
      $forCond = $Matches[1].Trim()
      $fid = & $AddNode 'loop' ("For " + $forCond) $lineNum
      & $AddEdge $prevId $fid

      # 階層管理システムにループ登録
      $loopInfo = $loopHierarchy.StartLoop('For', $fid, $lineNum)
      
      # 従来のスタック処理も維持（互換性のため）
      $stack.Push([pscustomobject]@{
        kind='For'; 
        loopId=$fid; 
        line=$lineNum;
        condition=$forCond;
        hierarchyId=$loopInfo.id
      })
      $prevId = $fid; continue
    }

    # Next のとき：Popし、Forまでを捨てる
    #＃if($trim -match '^\s*Next\b'){
    #  $top = Pop-UntilKind $stack 'For'
    #  if ($null -ne $top -and ($top.PSObject.Properties.Name -contains 'head')) {
    #    & $AddEdge $prevId $top.head 'next'
    #    $join = & $AddNode 'join' 'For Next End' $lineNum
    #    & $AddEdge $top.head $join 'exit'
    #    $prevId = $join
    #  } else {
    #    Write-Verbose "Skip Next at line $lineNum (no matching For on stack)"
    #  }
    #  continue
    #}

    # For Next処理でスパン登録（階層管理対応）
    if($trim -match '^\s*Next(?:\s+(.+))?\s*$'){
      if($stack.Count -gt 0 -and $stack.Peek().kind -eq 'For'){
        $top = $stack.Pop()
        $eid = & $AddNode 'loopEnd' 'For Next End' $lineNum
        & $AddEdge $prevId $eid

        # 階層管理システムでループ終了処理
        $loopInfo = $loopHierarchy.EndLoop($eid)

        # ループスパンを登録（階層ID付き）
        $span = [pscustomobject]@{
          headId = $top.loopId -replace '^n', ''
          endId = $eid -replace '^n', ''
          start = [int]$top.line
          end = [int]$lineNum
          hierarchyId = if($loopInfo) { $loopInfo.id } else { 0 }
          level = if($loopInfo) { $loopInfo.level } else { 0 }
        }
        $loopSpans.Add($span)
        Write-Verbose "Added hierarchical loop span: headId=$($span.headId), endId=$($span.endId), level=$($span.level), hierarchyId=$($span.hierarchyId)"
    
        $prevId = $eid
      }
      continue
    }



    # Select Case の検出とstack処理
    if($trim -match '^\s*Select\s+Case\s+(.+)'){
      $sid = & $AddNode 'switch' ("Select Case " + $Matches[1].Trim()) $lineNum
      & $AddEdge $prevId $sid
      # ★ ends と inCase を持たせる
      $stack.Push([pscustomobject]@{
        kind='Select'; head=$sid; line=$lineNum;
        ends = New-Object System.Collections.Generic.List[string];
        inCase = $false 
      })
      $prevId = $sid; continue
    }

    # Case のとき：Popしない。Select を上から探す
    if($trim -match '^\s*Case\s+(.+)'){
      $sel = Peek-NearestKind $stack 'Select'
      if ($null -ne $sel -and ($sel.PSObject.Properties.Name -contains 'head')) {
        # ★ 直前のCaseブランチをクローズ（末尾をendsへ）
        if($sel.inCase -and $null -ne $prevId -and $prevId -ne $sel.head){
          $sel.ends.Add($prevId)
        }
        # ★ 新しいCaseノードを作って head から繋ぐ
        $cid = & $AddNode 'case' ("Case " + $Matches[1].Trim()) $lineNum
        & $AddEdge $sel.head $cid
        $prevId = $cid
        $sel.inCase = $true
        # 更新を反映（PSCustomObjectは参照型なので上書き不要だが保険で入れ替え）
        $null = $stack.Pop(); $stack.Push($sel)
      } else {
        Write-Verbose "Skip Case at line $lineNum (no active Select on stack)"
      }
      continue
    }

    # Case Else
    if($trim -match '^\s*Case\s*Else\b'){
      $sel = Peek-NearestKind $stack 'Select'
      if ($null -ne $sel -and ($sel.PSObject.Properties.Name -contains 'head')) {
        if($sel.inCase -and $null -ne $prevId -and $prevId -ne $sel.head){
          $sel.ends.Add($prevId)
        }
        $cid = & $AddNode 'case' 'Case Else' $lineNum
        & $AddEdge $sel.head $cid
        $prevId = $cid
        $sel.inCase = $true
        $null = $stack.Pop(); $stack.Push($sel)
      } else {
        Write-Verbose "Skip Case Else at line $lineNum (no active Select on stack)"
      }
      continue
    }

    # End Select のとき：Select を取り除く（Select までを捨てる）
    if($trim -match '^\s*End\s*Select\b'){
      $sel = Pop-UntilKind $stack 'Select'
      $join = & $AddNode 'join' 'End Select' $lineNum

      if ($null -ne $sel) {
        # ★ 最後に開いていたCaseの末尾を回収
        if($sel.inCase -and $null -ne $prevId -and $prevId -ne $sel.head){
          $sel.ends.Add($prevId)
        }
        # ★ これまでのケース末尾をすべて合流へ
        foreach($e in $sel.ends){ & $AddEdge $e $join }

        # ★ ブランチが1つも無い/空だった場合の保険（head→join）
        if($sel.ends.Count -eq 0){
          & $AddEdge $sel.head $join
        }
      } else {
        # セーフティ：スタックが壊れていた場合でも現在位置から合流
        if($null -ne $prevId){ & $AddEdge $prevId $join }
      }

      #& $AddEdge $prevId $join
      $prevId = $join
      continue
    }

    # With ブロック開始
    if($trim -match '^\s*With\s+(.+)$'){
      $wid = & $AddNode 'block' ("With " + $Matches[1].Trim()) $lineNum
      & $AddEdge $prevId $wid
      $stack.Push([pscustomobject]@{ kind='With'; head=$wid; line=$lineNum })
      $prevId = $wid
      continue
    }

    # End With
    if($trim -match '^\s*End\s*With\b'){
      $top = Pop-UntilKind $stack 'With'
      if ($null -ne $top) {
        $join = & $AddNode 'join' 'End With' $lineNum
        & $AddEdge $prevId $join
        $prevId = $join
      }
      continue
    }

    # Goto (ラベルへのジャンプ)
    if($trim -match '^\s*Goto\s+([A-Za-z_][A-Za-z0-9_]*)'){
      $gid = & $AddNode 'goto' $trim $lineNum
      & $AddEdge $prevId $gid
      $lbl = $Matches[1].ToLower()
    
      # ★ まずノードIDに直接飛ばす
      if($labelNodeMap.ContainsKey($lbl)){
        & $AddEdge $gid $labelNodeMap[$lbl] 'goto'
      }elseif($labels.ContainsKey($lbl)){
      #if($labels.ContainsKey($lbl)){
        $targetLine = $StartLineOffset + $labels[$lbl]
        & $AddEdge $gid ("n$targetLine") 'goto'
      }
      $prevId = $gid; continue
    }

    # ラベル行:  MyLabel:
    if($trim -match '^\s*([A-Za-z_][A-Za-z0-9_]*)\s*:\s*$'){
      $lbl = $Matches[1].ToLower()
      $labels[$lbl] = $i   # ここは既存どおり: 行インデックス等を記録


      # ★ ここを追加：順次実行の “前のノード” を断ち切る
      $#prevId = $null
      # ★ ここを追加：独立アンカー（開始ノード）を作る
      $lid = & $AddNode 'label' ("Label: " + $Matches[1]) $lineNum
      $labelNodeMap[$lbl] = $lid
      # ★ ラベル配下の最初の文とつながるように、$prevId をアンカーにする
      $prevId = $lid
      continue
    }

    # Exit/Return の検出とstack処理
    if($trim -match '^\s*(Exit\s+(Sub|Function|Property)|Return)\b'){
      $eid = & $AddNode 'end' $trim $lineNum

      # ★ 直前ノードが条件なら "Yes" で結ぶ（片側を固定）
      $prevNode = $nodes | Where-Object { $_.id -eq $prevId } | Select-Object -First 1
      if($prevNode -and $prevNode.type -eq 'cond'){
        & $AddEdge $prevId $eid 'Yes'
      } else {
        & $AddEdge $prevId $eid
      }
      # Exit のあとに続くノードは無いので prevId は更新しない
      $prevId = $null
      continue      
    }

    # Err.Raise の検出とstack処理（エラー終了として扱う）
    if($trim -match '^\s*Err\.Raise\b'){
      $eid = & $AddNode 'end' $trim $lineNum

      # ★ 直前ノードが条件なら "Yes" で結ぶ（ElseIfやIf条件が真の場合）
      $prevNode = $nodes | Where-Object { $_.id -eq $prevId } | Select-Object -First 1
      if($prevNode -and $prevNode.type -eq 'cond'){
        & $AddEdge $prevId $eid 'Yes'
      } else {
        & $AddEdge $prevId $eid
      }
      # Err.Raise のあとに続くノードは無いので prevId は更新しない
      $prevId = $null
      continue      
    }

    # 呼び出し検出
    $hit = New-Object System.Collections.Generic.List[pscustomobject]
    if($trim -match $rxCall1){ $hit.Add([pscustomobject]@{ module=$null; name=$Matches[1] }) }
    foreach($m in [regex]::Matches($trim,$rxCall2)){
      $hit.Add([pscustomobject]@{ module=$m.Groups[1].Value; name=$m.Groups[2].Value })
    }
    foreach($m in [regex]::Matches($trim,$rxCall3)){
      $name = $m.Groups[1].Value
      if($trim -match '^\s*(Dim|Set|Let|If|ElseIf|While|Do|For|Select|Case|With|Else|End|Public|Private|Friend|Const|Option)\b'){ continue }
      if($trim -match '^\s*[A-Za-z_][A-Za-z0-9_]*\s*=\s*'+[regex]::Escape($name)+'\s*\('){ continue }
      if($trim -match '\.\s*'+[regex]::Escape($name)+'\s*\('){ continue }
      $hit.Add([pscustomobject]@{ module=$null; name=$name })
    }
    if($hit.count -gt 0){
      foreach($c in $hit){
        $target=$null; $resolved=$false
        if($c.module){
          $target = "$($c.module).$($c.name)"
          if($SymbolTable.ContainsKey($c.module) -and ($SymbolTable[$c.module] -contains $c.name)){ $resolved=$true }
        } else {
          if($SymbolTable.ContainsKey($ThisModule) -and ($SymbolTable[$ThisModule] -contains $c.name)){
            $target="$ThisModule.$($c.name)"; $resolved=$true
          } else {
            foreach($m in $SymbolTable.Keys){
              if($SymbolTable[$m] -contains $c.name){ $target="$m.$($c.name)"; $resolved=$true; break }
            }
            if(-not $target){ $target = $c.name }
          }
        }
        $calls.Add([pscustomobject]@{ target=$target; resolved=$resolved; line=$lineNum })
      }
      $cid = & $AddNode 'call' $trim $lineNum
      
      # If文スタック処理：関数呼び出しでも条件からの直接接続を考慮
      if($stack.Count -gt 0 -and $stack.Peek().kind -eq 'If'){
        $top = $stack.Peek()

        # needsYesが真で、かつElse処理を経由していない場合のみ条件から直接接続
        if($top.needsYes -and -not $top.hasElse){
          & $AddEdge $top.cond $cid 'Yes'
          $top.needsYes = $false  # 最初のYes分岐のみに適用
        } else {
          # 通常の前のノードからの接続
          if($prevId -and $prevId -ne $top.cond){
            & $AddEdge $prevId $cid
          }
        }
      } else {
        # If文スタック外では通常の接続
        & $AddEdge $prevId $cid
      }
      
      $prevId = $cid
      continue
    }

    # If文スタック処理：条件からの直接接続が必要な場合のみ
    if($stack.Count -gt 0 -and $stack.Peek().kind -eq 'If'){
      $top = $stack.Peek()
      $oid = & $AddNode 'op' $trim $lineNum

      # needsYesが真で、かつElse処理を経由していない場合のみ条件から直接接続
      if($top.needsYes -and -not $top.hasElse){
        & $AddEdge $top.cond $oid 'Yes'
        $top.needsYes = $false  # 最初のYes分岐のみに適用
      } else {
        # 通常の前のノードからの接続
        if($prevId -and $prevId -ne $top.cond){
          & $AddEdge $prevId $oid
        }
      }
      
      $prevId = $oid
      continue  # If文context下では通常文処理をスキップ
    }

    # その他の通常文
    $oid = & $AddNode 'op' $trim $lineNum
    
    # デバッグ: 通常文処理
    if($stack.Count -gt 0 -and $stack.Peek().kind -eq 'If'){
      $msg = '[{0}] {1}' -f (Get-Timestamp), ($messages.'flowjson.warn.ifContextProcessing' -f $lineNum, $trim)
      Write-Host $msg
    }
    
    & $AddEdge $prevId $oid
    $prevId = $oid
  }
 
  # 残ったスタックは無視して終了ノードへ
  $endLine = $StartLineOffset + $BodyLines.Length
  $endId   = & $AddNode 'end' 'End' $endLine
  & $AddEdge $prevId $endId

  return [pscustomobject]@{ nodes=$nodes; edges=$edges; calls=$calls; loopSpans=$loopSpans }
}

# 実行部分
# パラメーターの正規化
$msg = '[{0}] {1}' -f (Get-Timestamp), ($messages.'flowjson.info.originalPaths' -f $FolderPath, $FilePath)
Write-Host $msg

# 引用符を除去
$FolderPath = $FolderPath -replace '^"(.*)"$', '$1'
$FilePath = $FilePath -replace '^"(.*)"$', '$1'

$msg = '[{0}] {1}' -f (Get-Timestamp), ($messages.'flowjson.info.cleanedPaths' -f $FolderPath, $FilePath)
Write-Host $msg

# フォルダパスの正規化と存在チェック
$FolderPath = [System.IO.Path]::GetFullPath($FolderPath)
$msg = '[{0}] {1}' -f (Get-Timestamp), ($messages.'flowjson.info.fullFolderPath' -f $FolderPath, (Test-Path -LiteralPath $FolderPath -PathType Container))
Write-Host $msg

if(-not (Test-Path -LiteralPath $FolderPath -PathType Container)){ 
  Write-Warning "FolderPath not found, using file directory: $FolderPath"
  $FolderPath = [System.IO.Path]::GetDirectoryName($FilePath)
  $msg = '[{0}] {1}' -f (Get-Timestamp), ($messages.'flowjson.info.newFolderPath' -f $FolderPath)
  Write-Host $msg
}

$msg = '[{0}] {1}' -f (Get-Timestamp), ($messages.'flowjson.info.finalFilePath' -f $FilePath, (Test-Path -LiteralPath $FilePath -PathType Leaf))
Write-Host $msg

if(-not (Test-Path -LiteralPath $FilePath -PathType Leaf)){ throw "FilePath not found: $FilePath" }

$symbolTable = BuildSymbolTable -Folder $FolderPath -FileEncoding $Encoding

$raw      = @(ReadFileLines -Path $FilePath -FileEncoding $Encoding)
$lines    = @(JoinContinuations $raw)
$module   = GetModuleName -Lines $lines -Fallback $FilePath
$procs    = @(ParseProcedures -Lines $lines)   # ← 配列保証

# グローバル出力
$procOut  = New-Object System.Collections.Generic.List[pscustomobject]
$gNodes   = New-Object System.Collections.Generic.HashSet[string]
$gEdges   = New-Object System.Collections.Generic.List[pscustomobject]

# 各プロシージャのフローを解析
foreach($p in $procs){
  $flow = AnalyzeProcedure -BodyLines $p.body.ToArray() -StartLineOffset $p.startLine -ThisModule $module -SymbolTable $symbolTable
  $procOut.Add([pscustomobject]@{
    name=$p.name; kind=$p.kind; startLine=$p.startLine; endLine=$p.endLine;
    calls=$flow.calls;
    mermaid=[pscustomobject]@{ direction='TD'; nodes=$flow.nodes; edges=$flow.edges; loopSpans=$flow.loopSpans }
  })
  foreach($c in $flow.calls){
    $from = "$module.$($p.name)"; $to=[string]$c.target
    [void]$gNodes.Add($from); [void]$gNodes.Add($to)
    $gEdges.Add([pscustomobject]@{ from=$from; to=$to; resolved=$c.resolved })
  }
}

# 全体出力オブジェクト構築
$out = [pscustomobject]@{
  input=[pscustomobject]@{
    folderPath=(Resolve-Path -LiteralPath $FolderPath).Path
    filePath  =(Resolve-Path -LiteralPath $FilePath).Path
    moduleName=$module
  }
  symbols=$symbolTable
  procedures=$procOut
  mermaid_global=[pscustomobject]@{
    direction='TD'
    callgraph=[pscustomobject]@{ nodes=@($gNodes); edges=$gEdges }
  }
}

# JSON出力
$json = $out | ConvertTo-Json -Depth 20

if($OutputPath){ 
  $json | Out-FileUtf8NoBom -FilePath $OutputPath
  $msg = '[{0}] {1}' -f (Get-Timestamp), ($messages.'flowjson.info.jsonWritten' -f $OutputPath)
  Write-Host $msg
} elseif($OutputFolder) {
  # OutputFolderが指定された場合、モジュール名.flow.jsonとして出力
  $baseName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
  $outputPath = Join-Path $OutputFolder "$baseName.flow.json"
  $json | Out-FileUtf8NoBom -FilePath $outputPath
  $msg = '[{0}] {1}' -f (Get-Timestamp), ($messages.'flowjson.info.jsonWritten' -f $outputPath)
  Write-Host $msg
} else { 
  $json 
}
