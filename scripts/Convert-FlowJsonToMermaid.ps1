<#
    File: export_opened_vba.ps1
    Description: フローチャートJSONからmermaid形式のフローチャートを作成するスクリプト
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

param(
  [Parameter(Mandatory=$true)] [string]$JsonPath,
  [Parameter()] [string]$OutDir = "."
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

# ユーティリティ
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

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

# --- 安全ユーティリティ ---
function To-Array {
  param($x)
  if ($null -eq $x) { return @() }
  if ($x -is [System.Collections.IEnumerable] -and -not ($x -is [string])) { return @($x) }
  return ,$x
}

function SafeSort {
  param(
    $seq,
    $scriptBlock1,
    $scriptBlock2
  )

  $a = @($seq)
  # 0/1 件はそのまま配列で返す
  if ($a.Length -eq 0) { return @() }
  if ($a.Length -eq 1) { return ,$a[0] }

  # 必ず配列で返す
  try {
    $r = $a | Sort-Object $key1, $key2
    return @($r)                  
  } catch {
    return $a
  }
}

function SafeIndex {
  param($arr, [int]$i)
  $a = @($arr)   # ★必ず配列化
  if ($i -lt 0 -or $i -ge $a.Length) { return $null }
  return $a[$i]
}

function Is-Blank { param($s) return [string]::IsNullOrWhiteSpace([string]$s) }

function CountOf($x) {
  if ($null -eq $x) { return 0 }
  if ($x -is [System.Collections.IEnumerable] -and -not ($x -is [string])) {
    try { return (@($x)).Count } catch { return 1 }
  }
  return 1
}

function Get-OutEdges {
  param($map, [string]$fromId)

  if ($null -eq $map) { return @() }
  if (-not ($map -is [hashtable])) { return @() }
  if (-not $map.ContainsKey($fromId)) { return @() }

  $v = $map[$fromId]
  if ($null -eq $v) { return @() }

  # List だったら ToArray で取り出す
  if ($v -is [System.Collections.IEnumerable] -and -not ($v -is [string])) {
    try {
      return @($v | ForEach-Object { $_ })
    } catch {
      return ,$v
    }
  }

  # 単体オブジェクトの場合
  return ,$v
}


# ノードIDをMermaid書式に変換（日本語可読性を保持）
function SafeId {
  param([string]$s)
  if ([string]::IsNullOrEmpty($s)) { return "" }
  
  $sb = New-Object System.Text.StringBuilder
  foreach ($ch in $s.ToCharArray()) {
    if ($ch -match '[A-Za-z0-9_]') {
      [void]$sb.Append($ch)              # ASCII英数・アンダースコアはそのまま
    }
    elseif ($ch -eq '.' -or [char]::IsWhiteSpace($ch) -or $ch -eq '-') {
      [void]$sb.Append('_')              # 区切り記号は _
    }
    elseif ($ch -match '[(){}[\]|&<>"''`~!@#$%^*+=;:,/?\\]') {
      # Mermaid禁則文字のみUnicodeエスケープ
      [void]$sb.Append(('_u{0:X4}' -f [int][char]$ch))
    }
    else {
      # 日本語文字（ひらがな、カタカナ、漢字など）はそのまま保持
      [void]$sb.Append($ch)
    }
  }

  $id = $sb.ToString()
  $id = $id -replace '_{2,}','_'         # 連続 _ を1つに
  if ($id -match '^[0-9]') { $id = '_' + $id }  # 先頭が数字なら _
  return $id
}

# テキストをMermaid書式に変換（日本語可読性保持）
function Escape-Text {
  param(
    [string]$s,
    [switch]$ForCond  # ← 条件ノード用フラグ
  )
  if([string]::IsNullOrEmpty($s)){ return "" }
  if ($ForCond) {
    # 条件ノード：
    #   - <, >, <> はそのまま（比較演算子として保持）
    #   - 内部の " は &quot; に変換
    #   - Mermaid構文衝突する角括弧は () に
    #   - 改行や余分な空白を潰す
    #   - 日本語文字はそのまま保持
    $s = $s -replace '"','&quot;'
    $s = $s -replace '\[','('
    $s = $s -replace '\]',')'
    $s = $s -replace '\s+',' '
    $s = $s -replace "`r","" -replace "`n"," "
    return $s.Trim()
  } else {
    # 通常ノード：
    #   - HTMLエンティティ変換（最小限のみ）
    #   - Mermaid構文と衝突する文字のみエスケープ
    #   - 日本語文字はそのまま保持
    $s = $s -replace '&','&amp;'
    $s = $s -replace '<','&lt;'
    $s = $s -replace '>','&gt;'
    $s = $s -replace '"','&quot;'
    $s = $s -replace '\[','('
    $s = $s -replace '\]',')'
    $s = $s -replace '\s+',' '
    return $s.Trim()
  }
}

# ノードの役割判定
# 合流点（joinノード）かどうか
function Is-Join {
  param($n)
  # JSONが 'join' を付けてくれるのが理想。古いJSONでもテキストでフォールバック。
  return ($n.type -eq 'join' -or ($n.text -match '^(End If|Loop End|End Select)$'))
}

# ループの“見出し”だけ台形にする。終端（Loop End）は台形にしない。
function Is-LoopHeader {
  param($n)
  return ($n.type -eq 'loop' -and ($n.text -notmatch '^Loop End$'))
}

# ループの終端（Loop End）は逆台形にする
function Is-LoopEndJoin {
  param($n)
  return ($n.type -eq 'join' -and ($n.text -match '^(Loop End|For Next End)$'))
}

# ループ内にいるかどうか（行番号ベース）
function Is-InsideLoop {
  param([int]$fromLine, [object]$span)  # span: start/end
  return ($fromLine -ge $span.start -and $fromLine -le $span.end)
}

# === 階層管理対応の新しいユーティリティ関数 ===

# 階層対応の重複エッジ処理クラス
class HierarchicalEdgeProcessor {
  [hashtable]$processedNodes
  [hashtable]$conditionLabels
  
  HierarchicalEdgeProcessor() {
    $this.processedNodes = @{}
    $this.conditionLabels = @{}
  }
  
  # ノード処理済みチェック
  [bool] IsNodeProcessed([string]$nodeId, [string]$nodeType) {
    if($nodeType -eq 'op' -or $nodeType -eq 'call') {
      return $this.processedNodes.ContainsKey($nodeId)
    }
    return $false
  }
  
  # ノード処理済み登録
  [void] MarkNodeProcessed([string]$nodeId, [string]$nodeType) {
    if($nodeType -eq 'op' -or $nodeType -eq 'call') {
      $this.processedNodes[$nodeId] = $true
    }
  }
  
  # 条件ラベル重複チェック
  [bool] IsConditionLabelDuplicate([string]$fromId, [string]$label, [string]$toId, $edges) {
    if(-not $label) { return $false }
    
    $key = "$fromId|$label"
    if($this.conditionLabels.ContainsKey($key)) {
      # 既存のtoIdと異なる場合は重複
      return $this.conditionLabels[$key] -ne $toId
    } else {
      $this.conditionLabels[$key] = $toId
      return $false
    }
  }
}

# If-ElseIf-Else構造の3分岐処理
function Process-IfElseIfElse {
  param($condNode, $edges, $nodeMap)
  
  $fromId = $condNode.id
  $outEdges = @($edges | Where-Object { [string]$_.from -eq $fromId })
  
  if($outEdges.Count -ge 2) {
    # Yes/No/ElseIf分岐の適切な処理
    $yesEdges = @($outEdges | Where-Object { $_.label -match '^(Yes|True)$' })
    $noEdges = @($outEdges | Where-Object { $_.label -match '^(No|False)$' })  
    $elseIfEdges = @($outEdges | Where-Object { $_.label -match '^ElseIf' })
    
    # 3つの分岐が存在する場合の正規化
    if($yesEdges.Count -gt 0 -and $noEdges.Count -gt 0 -and $elseIfEdges.Count -gt 0) {
      # Write-Host "Found proper 3-branch If-ElseIf-Else structure from $fromId"
      return $true
    }
  }
  
  return $false
}

# 行のユニーク化（重複行を追加しない）
function Add-LineUnique {
  param(
    [System.Collections.Generic.List[string]]$list,
    [System.Collections.Generic.HashSet[string]]$seen,
    [string]$s,
    $nodeMap = $null
  )

  if ([string]::IsNullOrWhiteSpace($s)) { return }

  # 改行・タブ・余分な空白を除去し末尾をセミコロン統一
  $line = ($s -replace '[\r\n]+', '' -replace '\s+', ' ').Trim()
  if ($line -notmatch ';$') { $line += ';' }

  # 完全一致で重複防止
  if ($seen.Contains($line)) { return }

  # --- エッジ構造の補正 ---
  #if ($line -match '(\S+)\s*-->\s*(\S+);') {
  if ($line -match '(\S+)\s*-->(\|[^|]+\|)?\s*(\S+);') {
    $fromFull = $Matches[1]
    #$toFull   = $Matches[2]
    $label = $Matches[2]  # ラベル部分（|Yes|など）
    $toFull = $Matches[3]

    $fromId = if ($fromFull -match '_(n\d+)$') { 
      $fromFull -replace '^.*_(n\d+)$', '$1' 
    } else { 
      $fromFull 
    }

    $toId = if ($toFull -match '_(n\d+)$') { 
      $toFull -replace '^.*_(n\d+)$', '$1' 
    } else { 
      $toFull 
   }
    
    $fromType = $null; if ($nodeMap -and $nodeMap.ContainsKey($fromId)) { $fromType = $nodeMap[$fromId].type }
    $toType   = $null; if ($nodeMap -and $nodeMap.ContainsKey($toId))   { $toType   = $nodeMap[$toId].type }

    # 台形→逆台形の戻り線は出さない
    if ($fromType -eq 'loop' -and $toType -eq 'loopEnd') { return }
  }

  # --- 正常登録 ---
  [void]$list.Add($line)
  [void]$seen.Add($line)

}

# ノード行を生成
function Node-Line {
  param($prefix,$n)

  $id = SafeId "$prefix`_$($n.id)"


  # ★ ループ終端は逆台形（[\ text /]）で描く（引用符ナシ・余分な空白ナシ）
  if (Get-Command Is-LoopEndJoin -ErrorAction SilentlyContinue) {
    if (Is-LoopEndJoin $n) {
      $t = Escape-Text $n.text
      $t = $t -replace '"','' -replace '\s+',' '    # " と余分な空白除去
      return "$id[\""$t""/]"
    }
  }
  
  # 合流点はここで早期確定（ラベル無しの丸）
  if (Get-Command Is-Join -ErrorAction SilentlyContinue) {
    if (Is-Join $n) { return "$id(( ))" }
  }

  # ループの“見出し”だけは台形にする（必要な場合のみ）
  if (Get-Command Is-LoopHeader -ErrorAction SilentlyContinue) {
    if (Is-LoopHeader $n) {
      $t = Escape-Text $n.text
      # [/ text \] の公式記法。ダブルクォートは "" で表現
      return "$id[/""$t""\]"
    }
  }

  # 通常ノード
  switch ($n.type) {
    'label' { 
      return "$id([\""$text""])"
    }
    'spacer' {
      return "$id[ ]"
    } 
    'start' {
      $t = Escape-Text $n.text
      # 楕円（開始）ラベル
      return "$id([""$t""])"
    }
    'end' {
      $t = Escape-Text $n.text
      # 楕円（終端）ラベル
      return "$id([""$t""])"
    }
    'cond' {
      # 条件：{"..."} で必ずダブルクォートを付ける
      #    -ForCond では < > <> や " をエンコードしない実装にしておく
      $t = Escape-Text $n.text -ForCond
      return "$id{""$t""}"
    }
    'loop' {
      # ループ条件ノード（見出し以外に loop が来る場合）は、条件と同じひし形で統一
      $t = Escape-Text $n.text -ForCond
      return "$id{""$t""}"
    }
    'loopEnd' {  # ← これを追加
      $t = Escape-Text $n.text
      return "$id[\""$t""/]"
    }
    'call' {
      $t = Escape-Text $n.text
      # サブルーチン
      return "$id[[""$t""]]"
    }
    'switch' {
      $t = Escape-Text $n.text
      return "$id{{""$t""}}"
    }
    'join' {
      $t = Escape-Text $n.text
      #return "$id[\\""For Next End""/]"
        # ループ終端かどうかをテキストで判定
      if ($t -match '^(Loop End|For Next End)$') {
        return "$id[\""$t""/]"  # ← 逆台形
      } else {
        return "$id(( ))"      # ← 通常の合流点
      }
    }
    default {
      $t = Escape-Text $n.text
      # 通常処理ノードは "..." で統一
      return "$id[""$t""]"
    }
  }
}

# エッジラベルのマッピング
# If Not 条件かどうか
function Test-IfNot {
  param([string]$condText)
  return ($condText -match '(?i)^\s*If\s+Not\b')
}

# ノードIDから行番号を取得
#function Get-LineNumFromNodeId {
#  param([string]$nodeId)
#  if ($nodeId -match 'n(\d+)$') { return [int]$Matches[1] } else { return [int]::MaxValue }
#}

# If/Else/ElseIf の unlabeled を Yes/No に補完（If Not 反転対応）
function Map-CondLabelXX {
  param(
    [string]$fromId,
    [string]$toId,
    [object]$fromNode,
    [object[]]$siblings
  )
  # すでに No ラベルがある（ElseIf など）→ 残りの unlabeled は Yes
  $hasExplicitNo = $false
  foreach($s in $siblings){
    if(([string]$s.label) -match '^(?i)no$'){ $hasExplicitNo = $true; break }
  }
  if ($hasExplicitNo) { return 'Yes' }

  # 典型 If/Else（2枝とも unlabeled）
  #$unlabeled = @($siblings | Where-Object { [string]::IsNullOrWhiteSpace([string]$_.label) })
  $unlabeled = @($siblings | Where-Object { Is-Blank([string]$_.label) })  # ← 配列化
  #if ($unlabeled.Count -ge 2) {
  if ((CountOf $unlabeled) -ge 2) {
    #$sorted = $unlabeled | Sort-Object { Get-LineNumFromNodeId([string]$_.to) }
    $sorted = @($unlabeled | Sort-Object { [int]$nodeMap[[string]$_.to].line }, { [string]$_.to })  # ← 配列化
    $thenId = [string]$sorted[0].to
    $elseId = [string]$sorted[-1].to
    $isNot  = Test-IfNot ([string]$fromNode.text)

    if ($toId -eq $thenId) { if ($isNot) { return 'No' } else { return 'Yes' } }
    if ($toId -eq $elseId) { if ($isNot) { return 'Yes' } else { return 'No' } }
  }
  return $null
}

# ループの next/loop/exit を Yes/No に変換
function Map-LoopLabel {
  param(
    [string]$RawLabel,
    [string]$LoopHeaderText
  )
  # Do Until のときは継続/脱出の意味が反転
  $isUntil = ($LoopHeaderText -match '(?i)\bUntil\b')

  switch -regex ($RawLabel) {
    '^(loop|next)$' {
      if ($isUntil) { return 'No' } else { return 'Yes' }
    }
    '^exit$' {
      if ($isUntil) { return 'Yes' } else { return 'No' }
    }
    default {
      return $null
    }
  }
}

# エッジ行を生成
function Edge-Line {
  param(
    $prefix,
    $e,
    $nodeMap,
    $edgesByFrom = $null,
    $loopEndMap = $null,
    $loopSpans = $null,
    $condPlan = $null,
    $loopPlan = $null
  )

  # 受け取りが無ければ空にして落ちないように
  if($null -eq $edgesByFrom){ $edgesByFrom = @{} }
  if($null -eq $loopEndMap ){ $loopEndMap  = @{} }
  if($null -eq $loopSpans  ){ $loopSpans   = New-Object System.Collections.Generic.List[pscustomobject] }
  if($null -eq $condPlan   ){ $condPlan    = @{} }
  if($null -eq $loopPlan   ){ $loopPlan    = @{} }

  $fromId = [string]$e.from
  $toId   = [string]$e.to

  $from = SafeId "$prefix`_$fromId"
  $to   = SafeId "$prefix`_$toId"

  $lab = [string]$e.label
  $fromNode = $nodeMap[$fromId]
  $toNode   = $nodeMap[$toId]
  if($lab -eq 'Yes' -or $lab -eq 'No' -or $lab -eq 'Else' -or $lab -eq 'ElseIf'){
  
    # If Not判定の処理
    if(Test-IfNot $fromNode.text){
      if($lab -eq 'Yes'){ $lab = 'No' }
      elseif($lab -eq 'No'){ $lab = 'Yes' }
    }
    return "$from -->|$lab| $to"
  }

  # 兄弟エッジの取得は“安全に”
  $siblings = @()
  if($edgesByFrom.ContainsKey($fromId)){ $siblings = $edgesByFrom[$fromId] }

  # --- A) 見出し(loop) → 終端(join) の exit は「描かない」 ---
  if ($null -ne $fromNode -and $fromNode.type -eq 'loop' -and $lab -eq 'exit') {
  # 逆台形(join)に行く場合だけ描く（Noラベル無し）、他は抑止
    if ($toNode -and ($toNode.type -eq 'join' -or $toNode.text -match '^(Loop End|For Next End)$')) {
      return "$from --> $to"
    } else {
      return ""
    }
  }

  # --- B) 本体 → 台形(loop) の next/loop を「逆台形」に付け替え（内側限定） ---
  if ($fromNode -and ($lab -eq 'next' -or $lab -eq 'loop')) {
    # そのループの span を取得
    $span = $null
    $fromLine = [int]$fromNode.line
    foreach ($sp in $loopSpans) {
      if ($sp -and ($sp.PSObject.Properties.Name -contains 'headId')) {
        #if ($sp.headId -eq $toId) { $span = $sp; break }
        #$fromLine = [int]$fromNode.line
        if($fromId -eq 'n284'){
          # Write-Host "  Checking span: headId=$($sp.headId) vs toId=$toId"
        }
        # fromLineがスパンの範囲内かつ、ループの戻り先がtoIdと一致するかをチェック
        if ($fromLine -ge $sp.start -and $fromLine -le $sp.end) {
          # さらに、戻り先がループの開始点と一致するかをチェック
          if ($sp.headId -eq $toId) {
            $span = $sp
            if($fromId -eq 'n284'){
              # Write-Host "  MATCH FOUND: $($sp | ConvertTo-Json -Compress)"
            }
            break
          }
        }
      }
      if($fromId -eq 'n284'){
        # Write-Host "Selected span: $($span | ConvertTo-Json -Compress)"
      }
    }
    if ($span -and $span.endId) {
      $fromLine = [int]$fromNode.line

      if (Is-InsideLoop -fromLine $fromLine -span $span) {
          # nextは逆台形に向かう（台形には向かわない）
          $to2 = SafeId "$prefix`_$($span.endId)"
          # Write-Host "SUCCESS: Redirecting to endId: $to2"

          return "$from --> $to2"
      } else {
        # Write-Host "FAILED: Not inside loop"
      }
    } else {
      # Write-Host "FAILED: No span or endId"
    }
      # Write-Host "=========================="
      # スパン処理に該当しない場合は、nextラベルなしで通常接続
      return "$from --> $to"     
  }

  # ループ条件の Yes/No 変換
  if ($null -ne $fromNode -and $fromNode.type -eq 'loop') {
    $mapped = $null
    #if ([string]::IsNullOrWhiteSpace($lab)) {
    if (Is-Blank($lab)) {
    $mapped = Map-LoopLabel -RawLabel '' -LoopHeaderText $fromNode.text
    } else {
      $mapped = Map-LoopLabel -RawLabel $lab -LoopHeaderText $fromNode.text
    }
    if ($null -ne $mapped) { return "$from -->|$mapped| $to" }
  }

  # --- ★ cond→join は「自動ラベリングしない」：この判定を先に置く（順序が重要） ---
  if ($fromNode -and $fromNode.type -eq 'cond' -and $toNode -and $toNode.type -eq 'join') {
    #if ([string]::IsNullOrWhiteSpace($lab)) { return "$from --> $to" }
    if (Is-Blank($lab)) { return "$from --> $to" }
    else { return "$from -->|$lab| $to" }
  } 

  # unlabeled cond→X は condPlan に従って確定
  #if ($fromNode -and $fromNode.type -eq 'cond' -and [string]::IsNullOrWhiteSpace()) {
  if ($fromNode -and $fromNode.type -eq 'cond' -and (Is-Blank ($lab))) {
    $plan = $condPlan[$fromId]

    if ($plan) {
      if ($plan.then -and $toId -eq $plan.then) { return "$from -->|Yes| $to" }
      if ($plan.else -and $toId -eq $plan.else) { return "$from -->|No| $to" }
    }
    return "$from --> $to"
  }

  # --- ループ見出し(loop) の出力統一 ---
  if ($fromNode -and $fromNode.type -eq 'loop') {
    $lp = $loopPlan[$fromId]

    # ループ本体への枝
    if ($lp -and $lp.bodyFirst -and $toId -eq $lp.bodyFirst) {
      return "$from --> $to"
    }

    # 逆台形(end) への枝は削除
    if ($lp -and $lp.endId -and $toId -eq $lp.endId) {
      return "$from -.->|No| $to"
      return ""
    }

    # 誤って生成された「見出し→見出し」などは素通し
    if (Is-Blank($lab)) { return "$from --> $to" }
    else { return "$from -->|$lab| $to" }

    return ""
  }

  # デフォルト（ここに必ず落ちるように）
  if (Is-Blank($lab)) { return "$from --> $to" }
  else { return "$from -->|$lab| $to" }

}

# メイン処理
# JSON読み込み
if(-not (Test-Path -LiteralPath $JsonPath -PathType Leaf)){ throw "File not found: $JsonPath" }
$doc = Get-Content -LiteralPath $JsonPath -Raw -Encoding UTF8 | ConvertFrom-Json

# 出力先
$null = New-Item -ItemType Directory -Path $OutDir -Force


# 手続きごとのフローチャート
foreach($p in $doc.procedures){
  $prefix = "$($doc.input.moduleName).$($p.name)"
  $lines = [System.Collections.Generic.List[string]]::new()
  $seen  = [System.Collections.Generic.HashSet[string]]::new()

  $lines.Add("flowchart TD")

  # ノード/エッジを必ず配列に
  $procNodes = To-Array $p.mermaid.nodes
  $procEdges = To-Array $p.mermaid.edges

  # ノード辞書・fromごとのエッジ
  $nodeMap = @{}
  foreach($n in $p.mermaid.nodes) {
     $nodeMap[$n.id] = $n
  }

  $edgesByFrom = @{}
  foreach($e in $p.mermaid.edges){
    $fid = [string]$e.from
    if(-not $edgesByFrom.ContainsKey($fid)){
      $edgesByFrom[$fid] = New-Object System.Collections.Generic.List[object]
    }
    $edgesByFrom[$fid].Add($e)
  }

  # fromId -> edges[]
  #$loopSpans = @{}
  $loopSpans   = New-Object System.Collections.Generic.List[pscustomobject] 

  # 2) 逆台形判定（Loop End / For Next End）
  $loopEndMap = @{}
  foreach($n in $procNodes){
    if($n.type -eq 'join' -and ($n.text -match '^(Loop End|For Next End)$')){
      $loopEndMap[[string]$n.id] = $true
    }
    if(Is-LoopEndJoin $n){
      $loopEndMap[[string]$n.id] = $true
    }
  }
  # ループスパン情報の収集
  foreach($n in $procNodes){
    if(Is-LoopEndJoin $n){
      $endId = [string]$n.id
      $endLine = [int]$n.line
    
      # 対応するループヘッダーを探す
      $headId = $null
      $startLine = [int]::MaxValue
    
      foreach($n2 in $procNodes){
        if($n2.type -eq 'loop' -and [int]$n2.line -lt $endLine){
          # より近いヘッダーを選ぶ
          if([int]$n2.line -gt $startLine -or $startLine -eq [int]::MaxValue){
            $headId = [string]$n2.id
            $startLine = [int]$n2.line
          }
        }
      }
    
      if($headId){
        $span = [pscustomobject]@{
          headId = $headId
          endId = $endId
          start = $startLine
          end = $endLine
        }
        $loopSpans.Add($span)
        # Write-Host "Added loop span: headId=$headId, endId=$endId, start=$startLine, end=$endLine"
      }
    }
  }
  # 3) 条件の then/else 事前確定（condPlan）
  $condPlan = @{}
  foreach($fromId in To-Array @($edgesByFrom.Keys)){
    $fNode = $nodeMap[$fromId]
    if($fNode -and $fNode.type -eq 'cond'){
      $outs = @()
      foreach($e2 in Get-OutEdges -map $edgesByFrom -fromId $fromId){
        $tn = $nodeMap[[string]$e2.to]
        if(-not ($tn -and $tn.type -eq 'join')){ $outs += ,$e2 }
      }
      #if($outs.Count -ge 1){
      if((CountOf $outs) -ge 1){
        $sorted = SafeSort $outs { [int]$nodeMap[[string]$_.to].line } { [string]$_.to }
        $thenId = [string]$sorted[0].to
        $elseId = $null
        if((CountOf $sorted) -ge 2){ $elseId = [string](SafeIndex $sorted 1).to }
        $isNot = ($fNode.text -match '^\s*"?\s*(If|ElseIf)\s+Not\b')
        if($isNot){ $tmp=$thenId; $thenId=$elseId; $elseId=$tmp }
        $condPlan[$fromId] = [pscustomobject]@{ then=$thenId; else=$elseId }
      }
    }
  }

  # 4) ★ここで $loopPlan を作る（←この位置！）
  $loopPlan = @{}   # headId => @{ bodyFirst=<id>; endId=<id> }

  foreach($fromId in To-Array @($edgesByFrom.Keys)){

    $fNode = $nodeMap[$fromId]

    if($fNode -and $fNode.type -eq 'loop'){

      $outs = Get-OutEdges -map $edgesByFrom -fromId $fromId
      if((CountOf $outs) -gt 0){
        $endId = $null
        foreach($e2 in $outs){
          if($loopEndMap.ContainsKey([string]$e2.to)){ $endId = [string]$e2.to; break }
        }

        $cand = @()
        foreach($e2 in $outs){
          $tn = $nodeMap[[string]$e2.to]
          if(-not ($tn -and $tn.type -eq 'join')){ $cand += ,$e2 }
        }

        $bodyFirst = $null


        if((CountOf $cand) -gt 0){
          #$sorted = SafeSort $cand { [int]$nodeMap[[string]$_.to].line } { [string]$_.to }
          $sorted = SafeSort $cand `
            { $n = $nodeMap[[string]$_.to]; if($n -and $n.line) { [int]$n.line } else { [int]::MaxValue } } `
            { [string]$_.to }
        if ((@($sorted)).Length -ge 1) {
          $first = SafeIndex $sorted 0
          if ($first) { $bodyFirst = [string]$first.to }
        }
          $bodyFirst = [string](SafeIndex $sorted 0).to
        }

        $loopPlan[$fromId] = [pscustomobject]@{ bodyFirst=$bodyFirst; endId=$endId }
      }
    }
  }


  # 5) ノード出力
  foreach($n in $p.mermaid.nodes){
    $nl = (Node-Line -prefix $prefix -n $n)
    #if(-not [string]::IsNullOrWhiteSpace($nl)){
    if(-not (Is-Blank($nl))){
      $lines.Add(("  {0};" -f $nl))
    }
  }

  # 6) エッジ出力（階層管理システム対応）
  # 階層対応エッジプロセッサー初期化
  $edgeProcessor = [HierarchicalEdgeProcessor]::new()

  foreach($e in $p.mermaid.edges){
    $skipEdge = $false
    $fromNode = $nodeMap[[string]$e.from]
    $toNode = $nodeMap[[string]$e.to] 
    
    # === 階層管理システム対応の重複エッジチェック ===
    # 1. 条件ノードからの重複ラベルチェック
    if (-not $skipEdge -and $fromNode -and $fromNode.type -eq 'cond' -and $e.label) {
      if($edgeProcessor.IsConditionLabelDuplicate([string]$e.from, $e.label, [string]$e.to, $p.mermaid.edges)) {
        $skipEdge = $true
        # Write-Host "SKIPPED duplicate condition label: $($e.from) →|$($e.label)| $($e.to)"
      }
    }
    
    # 2. op/callノードからの重複エッジチェック（階層管理版）
    if (-not $skipEdge -and $fromNode -and ($fromNode.type -eq 'op' -or $fromNode.type -eq 'call')) {
      $fromId = [string]$e.from
      if($edgeProcessor.IsNodeProcessed($fromId, $fromNode.type)) {
        $skipEdge = $true
        # Write-Host "SKIPPED duplicate edge from $($fromNode.type) node: $fromId → $($e.to)"
      } else {
        $edgeProcessor.MarkNodeProcessed($fromId, $fromNode.type)
      }
    }
    
    # 3. If-ElseIf-Else 3分岐構造の処理
    if (-not $skipEdge -and $fromNode -and $fromNode.type -eq 'cond') {
      Process-IfElseIfElse -condNode $fromNode -edges $p.mermaid.edges -nodeMap $nodeMap | Out-Null
    }

    if (-not $skipEdge) {

      # 2. 既存のループチェック処理...
      foreach($headId in $loopPlan.Keys){
        $lp = $loopPlan[$headId]
        if($lp -and $lp.endId){
          # 台形→逆台形の直接エッジをスキップ
          if([string]$e.from -eq $headId -and [string]$e.to -eq $lp.endId){
            $skipEdge = $true
            break
          }
        }
      }
    }

    # === ElseIf構造の不正なエッジをフィルタリング（ここに追加） ===
    if (-not $skipEdge -and $fromNode -and $fromNode.type -eq 'call') {
      $toNode = $nodeMap[[string]$e.to]
    
      # ElseIf処理 → Else処理の不正なパターンを検出
      if ($fromNode.text -match '^maskOct\(i\)\s*=\s*0$' -and 
          $toNode -and $toNode.text -match '^\(\(.*\)\)|\^|And|Or') {
        $skipEdge = $true
        # Write-Host "SKIPPED invalid ElseIf→Else edge: $($e.from) → $($e.to)"
      }
    }

    if(-not $skipEdge){     
      if ([string]$e.from -eq 'n338') {
        # Write-Host "Processing n338 edge: $($e.from) → $($e.to)"
      }   
      $el = Edge-Line -prefix $prefix -e $e -nodeMap $nodeMap -edgesByFrom $edgesByFrom -loopEndMap $loopEndMap -loopSpans $loopSpans -condPlan $condPlan -loopPlan $loopPlan
      #if($el){ $lines.Add(("  {0};" -f $el)) }
      Add-LineUnique -list $lines -seen $seen -s $el -nodeMap $nodeMap
    }  

  }

  # === ここにループバック線生成を移動 ===
  if ($p.mermaid.loopSpans) {
    # Write-Host "=== HIERARCHICAL LOOP SPANS DEBUG ==="
    # Write-Host "loopSpans count: $($p.mermaid.loopSpans.Count)"

    # 階層レベル順にソート（内側のループから処理）
    $sortedSpans = $p.mermaid.loopSpans | Sort-Object { 
      if(($_ -ne $null) -and ($_.PSObject.Properties.Name -contains 'level') -and ($_.level -ne $null)) { 
        [int]$_.level 
      } else { 
        0 
      }
    } -Descending

    foreach ($span in $sortedSpans) {
      # Write-Host "Processing hierarchical span: headId=$($span.headId), endId=$($span.endId)"

      if ($span -and $span.headId -and $span.endId) {
        $loopHead = SafeId "$prefix`_n$($span.headId)"
        $loopEnd = SafeId "$prefix`_n$($span.endId)"

        # Write-Host "Generated IDs: loopHead=$loopHead, loopEnd=$loopEnd"
  
        # 階層レベルに応じたスタイルのループバック線を生成（従来の形式も対応）
        $backEdge = "$loopEnd -.-> $loopHead"
        # Write-Host "Generated hierarchical edge: $backEdge"

        # 階層重複チェック付きで追加
        $edgeKey = "$loopEnd->$loopHead"
        if (-not $seen.Contains($backEdge) -and -not $seen.Contains($edgeKey)) {
          [void]$lines.Add($backEdge)
          [void]$seen.Add($backEdge)
          [void]$seen.Add($edgeKey)
          # Write-Host "SUCCESSFULLY added hierarchical loop: $backEdge"
        } else {
          # Write-Host "SKIPPED (already exists): $backEdge"
        }
      } else {
        # Write-Host "SKIPPED (invalid span): headId=$($span.headId), endId=$($span.endId)"
      }
    }
    # Write-Host "========================"
  } else {
    $msg = '[{0}] {1}' -f (Get-Timestamp), $messages.'mermaid.info.noLoopSpans'
    Write-Host $msg
  }

  # 7)（任意）戻り矢印 end→head を合成して 1 本だけ追加
  foreach($headId in $loopPlan.Keys){
    $lp = $loopPlan[$headId]
    if($lp -and $lp.endId){
      $from = SafeId "$prefix`_$($lp.endId)"   # 逆台形
      $to   = SafeId "$prefix`_$headId"        # 台形
      #$lines.Add("  $from --> $to;")
      $back = "$from -.-> $to"
      Add-LineUnique -list $lines -seen $seen -s $back -nodeMap $nodeMap
    
      # 逆台形から次のノードへの接続を検索
      $nextNodes = @()
      foreach($e in $p.mermaid.edges){
        if([string]$e.from -eq $lp.endId){
          $toNode = $nodeMap[[string]$e.to]
          if($toNode -and $toNode.type -ne 'join' -and [string]$e.to -ne $headId){
            $nextNodes += ,$e
          }
        }
      }
    }
  }

  # 何本書けたか確認
  # $written = ($lines | Where-Object { $_ -match '\-\-\>' }).Count

  $mmdPath = Join-Path $OutDir ("{0}.{1}.mmd" -f $doc.input.moduleName,$p.name)
  ($lines -join "`r`n") | Out-FileUtf8NoBom -FilePath $mmdPath
  $msg = '[{0}] {1}' -f (Get-Timestamp), ($messages.'mermaid.info.mmdWritten' -f $mmdPath)
  Write-Host $msg
}

# 全体の呼び出しグラフ（簡易版）
$cg = $doc.mermaid_global.callgraph
$glines = New-Object System.Collections.Generic.List[string]
$gseen = New-Object System.Collections.Generic.HashSet[string]  # 追加
$glines.Add("flowchart TD")

# ノード
foreach($n in $cg.nodes){
  $id = $n -replace '[^A-Za-z0-9_\.]','_'
  $txt = Escape-Text $n
  $glines.Add(("  {0};" -f "$id[ $txt ]"))  
}

# エッジ
foreach($e in $cg.edges){
  $from = $e.from -replace '[^A-Za-z0-9_\.]','_'
  $to   = $e.to   -replace '[^A-Za-z0-9_\.]','_'

  # 解決済みなら実線、未解決なら点線で描く
  if($e.resolved){
    #$glines.Add(("  {0};" -f "$from --> $to"))
    #$glines = "$from --> $to"
    $edgeLine = "$from --> $to"
  } else{
    #$glines.Add(("  {0};" -f "$from -- ""unresolved"" --> $to"))
    $edgeLine = "$from -- ""unresolved"" --> $to"
  }
  #Add-LineUnique -list $glines -seen $seen -s $edgeLine -nodeMap $nodeMap
  Add-LineUnique -list $glines -seen $gseen -s $edgeLine -nodeMap $null

}

# ファイル出力
$callPath = Join-Path $OutDir ("{0}.callgraph.mmd" -f $doc.input.moduleName)
($glines -join "`r`n") | Out-FileUtf8NoBom -FilePath $callPath
$msg = '[{0}] {1}' -f (Get-Timestamp), ($messages.'mermaid.info.mmdWritten' -f $callPath)
Write-Host $msg
