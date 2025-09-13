<#
    File: FindAndRun-ExcelMacroByModule.ps1
    Description: 起動中の Excel から 指定モジュールの Public Sub 一覧を取得し、任意の マクロを実行できる PowerShell スクリプト
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
.SYNOPSIS
  起動中の Excel から指定モジュールの Public Sub を列挙／実行する

.PARAMETER ModuleName
  対象の VBA モジュール名（VB_Name）。例: Module1

.PARAMETER BasPath
  対応する .bas ファイルの絶対パス（省略可）。
  内容ハッシュで一致するブックを優先的に特定するために使用。

.PARAMETER ListOutput
  "JSON" を指定すると一覧を JSON で出力（既定: JSON）。他形式は未実装。

.PARAMETER ProcName
  実行したい Public Sub 名。指定された場合、そのマクロを実行する。

.EXAMPLE
  # 一覧出力（JSON）
  powershell -ExecutionPolicy Bypass -File .\FindAndRun-ExcelMacroByModule.ps1 `
    -ModuleName Module1 -ListOutput JSON

.EXAMPLE
  # Book1.xlsm の Module1.DoWork を実行
  powershell -ExecutionPolicy Bypass -File .\FindAndRun-ExcelMacroByModule.ps1 `
    -Qualified  "'Book1.xlsm'!Module1.DoWork"
    
.EXAMPLE
  # Module1.DoWork を実行
  powershell -ExecutionPolicy Bypass -File .\FindAndRun-ExcelMacroByModule.ps1 `
    -ModuleName Module1 -ProcName DoWork
#>

[CmdletBinding(DefaultParameterSetName='ByName')]
param(
  # === ByName（一覧/名前指定）===
  [Parameter(Mandatory=$true, ParameterSetName='ByName')]
  [string]$ModuleName,

  [Parameter(ParameterSetName='ByName')]
  [string]$ProcName,

  [Parameter(ParameterSetName='ByName')]
  [string]$WorkbookName,

  [Parameter(ParameterSetName='ByName')]
  [string]$BasPath,

  [Parameter(ParameterSetName='ByName')]
  [ValidateSet("JSON")]
  [string]$ListOutput = "JSON",

  # === ByQualified（完全修飾で即実行）===
  [Parameter(Mandatory=$true, ParameterSetName='ByQualified')]
  [string]$Qualified,

  # === 共通スイッチ ===
  [Parameter(ParameterSetName='ByName')]
  [Parameter(ParameterSetName='ByQualified')]
  [switch]$ActivateExcel,

  [Parameter(ParameterSetName='ByName')]
  [Parameter(ParameterSetName='ByQualified')]
  [switch]$ShowVBE,

  [Parameter(ParameterSetName='ByName')]
  [Parameter(ParameterSetName='ByQualified')]
  [switch]$ShowStatus
)

$ErrorActionPreference = 'Stop'
# --- 出力の文字化け対策（UTF-8） ---
try {
  [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
  $OutputEncoding = [System.Text.Encoding]::UTF8
} catch {}

# 共通ユーティリティ ExcelUtil.ps1 を同梱ディレクトリから読み込む
$utilPath = Join-Path $PSScriptRoot 'ExcelUtil.ps1'
if (-not (Test-Path -LiteralPath $utilPath)) {
  Write-Output (ConvertTo-Json @{ ok=$false; error="ExcelUtil.ps1 not found"; path=$utilPath })
  exit 0
}

. $utilPath   # ← dot-source（必ず先に実行される位置に）

# COM 再試行ヘルパ（既にあれば重複定義しない）
if (-not (Get-Command Invoke-Com -ErrorAction SilentlyContinue)) {
  function Invoke-Com {
    param([scriptblock]$Action, [int]$MaxTry = 50, [int]$DelayMs = 150)
    for ($i=0; $i -lt $MaxTry; $i++) {
      try { return & $Action } catch [System.Runtime.InteropServices.COMException] {
        if ($_.Exception.HResult -eq -2147418111) { Start-Sleep -Milliseconds $DelayMs; continue } # RPC_E_CALL_REJECTED
        throw
      }
    }
    throw "COM 呼び出しの再試行が上限に達しました"
  }
}


# Excel レディ待ち
if (-not (Get-Command Wait-ExcelReady -ErrorAction SilentlyContinue)) {
  function Wait-ExcelReady { param($App, [int]$TimeoutSec = 20)
    $sw = [Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSec) {
      try {
        if ((Invoke-Com { $App.Ready }) -and (Invoke-Com { $App.Interactive })) { return }
      } catch {}
      Start-Sleep -Milliseconds 200
    }
  }
}

# User32 定義がなければ追加
if (-not ("User32" -as [type])) {
  Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class User32 {
  [DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
}
"@
}

# Microsoft.VisualBasic をロード（AppActivate用）
if (-not ("Microsoft.VisualBasic.Interaction" -as [type])) {
  Add-Type -AssemblyName Microsoft.VisualBasic
}

# === メイン処理 ===
[MessageFilter]::Register()
try {
  $excel = Get-ExcelSafe
  Wait-ExcelReady -App $excel

  # Qualified指定を判定して実行
  if (-not [string]::IsNullOrWhiteSpace($Qualified)) {
    if ($ActivateExcel) { 
      Show-ExcelFront -Excel $excel
    }
    if ($ShowVBE) {
      try {
        $excel.VBE.MainWindow.Visible = $true 
      } catch {}
    }
    try {
      Wait-ExcelReady -App $excel
      Invoke-Com { $excel.Run($Qualified) } | Out-Null
      Write-Output (ConvertTo-Json @{ ok = $true; ran = $Qualified } -Depth 5)
      return
  #    exit 0
    } catch {
      Write-Output (ConvertTo-Json @{ ok = $false; ran = $Qualified; lastError = "$($_.Exception.Message)" } -Depth 8)
      return
  #    exit 0
    }
  }

  # BasPath 一致でブックを特定できるなら優先
  $targetWb = $null
  if ($BasPath) {
    $targetWb = Find-DetectWorkbookByBas -Excel $excel -ModuleName $ModuleName -BasPath $BasPath
  }

  $results = @()
  $wbs = Invoke-Com { @($excel.Workbooks) }
  foreach ($wb in $wbs) {
    if ($null -ne $targetWb -and ($wb.Name -ne (Invoke-Com { $targetWb.Name }))) { continue }
    try {
      $vbp = Invoke-Com { $wb.VBProject }
      $comps = Invoke-Com { @($vbp.VBComponents) }
      foreach ($c in $comps) {
        if ((Invoke-Com { $c.Name }) -ne $ModuleName) { continue }
        $results += Get-ModulePublicSubs -VBComponent $c -WorkbookName (Invoke-Com { $wb.Name })
      }
    } catch {
      # VBE へのアクセス権が無い場合等
      continue
    }
  }

  if ($ProcName) {
    # 実行要求
    if (-not $results -or -not ($results | Where-Object { $_.Proc -eq $ProcName })) {
      Write-Output ("{0}" -f (ConvertTo-Json @{ error="macro not found"; module=$ModuleName; proc=$ProcName; count=$results.Count } -Depth 5))
      exit 1
    }

    # 候補を絞る
    $filtered = $results | Where-Object { $_.Proc -eq $ProcName }

    # WorkbookName が指定されていたらさらに絞り込み
    if ($WorkbookName) {
        $filtered = $filtered | Where-Object { $_.WorkbookName -eq $WorkbookName }
    }

    if (-not $filtered -or $filtered.Count -eq 0) {
        Write-Output (ConvertTo-Json @{ ok=$false; error="macro not found"; module=$ModuleName; proc=$ProcName; workbook=$WorkbookName; count=$results.Count } -Depth 6)
        exit 0
    }

    # 実行対象を決める（BasPathで特定されていればそのブック、そうでなければ最初の一致）
    ##$target = $results | Where-Object { $_.Proc -eq $ProcName } | Select-Object -First 1
    # ★最初の一致（ここで Book2 が指定されていれば Book2 だけ残る）
    $target = $filtered | Select-Object -First 1
    #$wbName = $target.WorkbookName
    $qual   = $target.Qualified  # 例：'Book1.xlsm'!Module1.DoWork

    $qual = $target.Qualified  # 例：'Book1.xlsm'!Module1.aaa

    function Invoke-MacroRobust {
      param([object]$Excel, [string]$Qualified, [string]$Module, [string]$Proc)

      $tryForms = @(
      $Qualified,                   # 'Book1.xlsm'!Module1.aaa
      "$Module.$Proc",              # Module1.aaa
      $Proc                         # aaa
      )

      foreach ($form in $tryForms) {
        Write-Host "[DBG] Trying form: <$form>"
        try {
          Wait-ExcelReady -App $Excel
          Invoke-Com { $Excel.Run($form) } | Out-Null
          return @{ ok = $true; ran = $form }
        } catch {
          $msg = $_.Exception.Message
          $hr  = if ($_.Exception.HResult) { ('0x{0:X8}' -f ($_.Exception.HResult -band 0xffffffff)) } else { $null }
          # 1004（ランタイムエラー）等は次の表記でリトライ
          $last = @{ tried = $form; error = $msg; hresult = $hr }
        }
      }
      return @{ ok = $false; ran = $Qualified; lastError = $last }
   }

    # 実行直前
    if ($ActivateExcel) {
      Show-ExcelFront -Excel $excel
    }
    if ($ShowVBE) {
      try { $excel.VBE.MainWindow.Visible = $true } catch {}
    }
    if ($ShowStatus) {
      try { $excel.StatusBar = "Running $ModuleName.$ProcName ..." } catch {}
    }

    #$result = Invoke-MacroRobust -Excel $excel -Qualified $qual -Module $ModuleName -Proc $ProcName
    # ★フォールバック禁止：Qualified だけ実行
    try {
        Wait-ExcelReady -App $excel
        Invoke-Com { $excel.Run($qual) } | Out-Null
        $result = @{ ok = $true; ran = $qual }
    } catch {
        $result = @{ ok = $false; ran = $qual; lastError = $_.Exception.Message }
    }

    # 実行後
    if ($ShowStatus) {
      try {
        $msg = if ($result.ok) { "Done: $($result.ran)" } else { "Failed: $($result.ran)" }
        Invoke-Com { $excel.StatusBar = $msg } | Out-Null
        Start-Sleep -Milliseconds 800
        Invoke-Com { $excel.StatusBar = $null } | Out-Null
      } catch {}
    }
    
    Write-Output (ConvertTo-Json $result -Depth 8)
    exit 0   # ★ 失敗でも 0 で終了（Node 側で JSON を必ず受け取れる）

  }
  else {
    # 一覧出力
    if ($ListOutput -eq "JSON") {
      Write-Output (ConvertTo-Json $results -Depth 5)
    } else {
      $results | Format-Table -AutoSize
    }
  }

} finally {
  [MessageFilter]::Revoke()
}
