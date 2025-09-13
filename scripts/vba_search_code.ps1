<#
    File: export_opened_vba.ps1
    Description: 開いているEXCELのVBAモジュール内を文字列検索するスクリプト
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
  [Parameter(Mandatory=$true)][string]$Query,
  [switch]$UseRegex,
  [switch]$CaseSensitive,
  [string]$WorkbookFilter,
  [string]$ModuleFilter,
  [int]$MaxResults = 200,
  [int]$ContextLines = 2
)

# --- Force UTF-8 (no BOM) for stdout/stderr ---
try {
  [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
  $OutputEncoding           = [Console]::OutputEncoding
} catch {}

$ErrorActionPreference = 'Stop'

function J([object]$o){ $o | ConvertTo-Json -Depth 8 }
function Get-ExportExt([int]$Type){ switch ($Type) { 1{'bas'} 3{'frm'} default{'cls'} } }

# ★ 改名: Split-Lines → SplitLines
function SplitLines([string]$s){
  [System.Text.RegularExpressions.Regex]::Split($s, "\r\n|\n|\r")
}

# ★ 改名: Count-NewlinesUpTo → CountNewlinesUpTo
function CountNewlinesUpTo([string]$s, [int]$idx){
  if ($idx -le 0) { return 0 }
  $sub = $s.Substring(0, [Math]::Min($idx, $s.Length))
  return ([System.Text.RegularExpressions.Regex]::Matches($sub, "\r\n|\n|\r")).Count
}

try {
  # 正規表現準備
  if ($UseRegex) { $pat = $Query } else { $pat = [regex]::Escape($Query) }
  if (-not $CaseSensitive) { $pat = '(?i)' + $pat }
  $rx = [regex]::new($pat)

  # Excel 取得
  try {
    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
  } catch {
    Write-Output (J @{ ok=$false; error='excel_not_found' }); exit 0
  }
  if (-not $excel) { Write-Output (J @{ ok=false; error='excel_null' }); exit 0 }

  $hits = New-Object System.Collections.Generic.List[object]
  $diag = New-Object System.Collections.Generic.List[object]

  foreach ($wb in @($excel.Workbooks)) {
    if ($WorkbookFilter -and $wb.Name -ne $WorkbookFilter) { continue }

    try { $vbp = $wb.VBProject } catch {
      $diag.Add([pscustomobject]@{ wb=$wb.Name; vbProject='unavailable' }) | Out-Null
      continue
    }
    if (-not $vbp) { $diag.Add([pscustomobject]@{ wb=$wb.Name; vbProject='null' }) | Out-Null; continue }

    foreach ($c in @($vbp.VBComponents)) {
      if ($ModuleFilter -and $c.Name -ne $ModuleFilter) { continue }

      try { $cm = $c.CodeModule } catch { $cm = $null }
      if (-not $cm) { $diag.Add([pscustomobject]@{ wb=$wb.Name; mod=$c.Name; codeModule='null' }) | Out-Null; continue }

      try { $total = $cm.CountOfLines } catch { $diag.Add([pscustomobject]@{ wb=$wb.Name; mod=$c.Name; countOfLines='failed' }) | Out-Null; continue }
      if ($total -le 0) { $diag.Add([pscustomobject]@{ wb=$wb.Name; mod=$c.Name; empty=$true }) | Out-Null; continue }

      # 1) Proc 単位で検索
      $line = 1; $procHit = $false
      while ($line -le $total) {
        $kind = $null; try { $proc = $cm.ProcOfLine([int]$line, [ref]$kind) } catch { $proc = $null }
        if ($proc) {
          $start = $cm.ProcStartLine($proc, $kind)
          $count = $cm.ProcCountLines($proc, $kind)
          $code  = $cm.Lines($start, $count)
          if ($rx.IsMatch($code)) {
            $arr = SplitLines $code
            $rel = 1
            for ($i=0; $i -lt $arr.Length; $i++) { if ($rx.IsMatch($arr[$i])) { $rel = $i+1; break } }
            $matchLine = $start + $rel - 1

            $from = [Math]::Max(1, $rel - $ContextLines)
            $to   = [Math]::Min($arr.Length, $rel + $ContextLines)
            $snippet = ($arr[($from-1)..($to-1)] -join ' ').Trim()

            $compType  = [int]$c.Type
            $exportExt = Get-ExportExt $compType
            $hits.Add([pscustomobject]@{
              workbook   = $wb.Name
              module     = $c.Name
              proc       = $proc
              startLine  = $start
              matchLine  = $matchLine
              snippet    = $snippet
              qualified  = "'$($wb.Name)'!$($c.Name).$proc"
              compType   = $compType
              exportExt  = $exportExt
            }) | Out-Null

            if ($hits.Count -ge $MaxResults) { break }
            $procHit = $true
          }
          $line = $start + $count
        } else { $line++ }
        if ($hits.Count -ge $MaxResults) { break }
      }

      if ($hits.Count -ge $MaxResults) { break }

      # 2) モジュール全体フォールバック
      if (-not $procHit) {
        try { $moduleCode = $cm.Lines(1, $total) } catch { $moduleCode = $null }
        if (-not [string]::IsNullOrEmpty($moduleCode) -and $rx.IsMatch($moduleCode)) {
          $m = [regex]::Match($moduleCode, $rx)
          $firstIdx = $m.Index
          $absLine  = (CountNewlinesUpTo $moduleCode $firstIdx) + 1

          $kind2 = $null; $proc2 = $null; try { $proc2 = $cm.ProcOfLine([int]$absLine, [ref]$kind2) } catch {}
          $start2 = if ($proc2) { $cm.ProcStartLine($proc2, $kind2) } else { 1 }

          $arrAll = SplitLines $moduleCode
          $from = [Math]::Max(1, $absLine - $ContextLines)
          $to   = [Math]::Min($arrAll.Length, $absLine + $ContextLines)
          $snippet2 = ($arrAll[($from-1)..($to-1)] -join ' ').Trim()

          $compType  = [int]$c.Type
          $exportExt = Get-ExportExt $compType
          $hits.Add([pscustomobject]@{
            workbook   = $wb.Name
            module     = $c.Name
            proc       = $proc2
            startLine  = $start2
            matchLine  = $absLine
            snippet    = $snippet2
            qualified  = if ($proc2) { "'$($wb.Name)'!$($c.Name).$proc2" } else { "'$($wb.Name)'!$($c.Name)" }
            compType   = $compType
            exportExt  = $exportExt
          }) | Out-Null
        } else {
          $diag.Add([pscustomobject]@{ wb=$wb.Name; mod=$c.Name; moduleFallback='no_match' }) | Out-Null
        }
      }
    }

    if ($hits.Count -ge $MaxResults) { break }
  }

  Write-Output (J @{
    ok            = $true
    query         = $Query
    useRegex      = [bool]$UseRegex
    caseSensitive = [bool]$CaseSensitive
    count         = $hits.Count
    hits          = $hits
    diag          = $diag
  })
  exit 0

} catch {
  Write-Output (J @{ ok=$false; error='unexpected'; detail="$($_.Exception.Message)" })
  exit 0
}
