<#
    File: import_single_module.ps1
    Description: MCP write-tool (excel_update_module_code) 用の単一モジュールインポートエントリポイント。
                 import_opened_vba.ps1 の Import-ModuleToVBProject / Import-FormToVBProject を
                 -FunctionsOnly dot-source で再利用し、フォルダ一括インポートを経由せずに
                 「1ワークブック・1モジュール」だけを安全に差し替える。
    Author: Eitaro SETA
    License: MIT License
#>

param (
    [string]$WorkbookPath,                                   # 対象ワークブックのフルパス（優先。Resolve-TargetWorkbookに渡す）
    [string]$WorkbookName,                                   # フルパス未指定時のフォールバック（既にオープン済みのブック名で照合）
    [Parameter(Mandatory=$true)][string]$ModuleName,        # 対象モジュール名（ModuleType未指定時は既存モジュールのみ。指定時は新規作成するモジュール名）
    [Parameter(Mandatory=$true)][string]$SourceCodePath,    # 書き込むコード本文が入ったUTF-8テキストファイルのパス
    [string]$ModuleType,                                     # 指定時のみ新規作成モード。"standard"=StdModule(.bas) / "class"=Class(.cls)
    [string]$ScriptsDir = $PSScriptRoot
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$OutputEncoding           = [Console]::OutputEncoding

$script:launchedExcelPid = $null

function Write-ResultAndExit {
    param([hashtable]$Result, [int]$ExitCode)
    # このツール呼び出しでExcelを新規起動していた場合、そのPIDを常に結果へ含める。
    # 呼び出し側が「AI経由で自動起動されたExcelプロセス」を識別し、必要なら後始末できるようにするため。
    if ($script:launchedExcelPid) { $Result["launchedExcelPid"] = $script:launchedExcelPid }
    ($Result | ConvertTo-Json -Depth 6)
    exit $ExitCode
}

. (Join-Path $ScriptsDir "ExcelUtil.ps1")
. (Join-Path $ScriptsDir "import_opened_vba.ps1") -FunctionsOnly

if (-not (Test-Path -LiteralPath $SourceCodePath)) {
    Write-ResultAndExit @{ ok = $false; error = "ERR_SOURCE_CODE_NOT_FOUND"; path = $SourceCodePath } 1
}

try {
    $r = Get-OrStartExcelApplication
    $excel = $r.App
    $script:launchedExcelPid = $r.LaunchedProcessId
} catch {
    Write-ResultAndExit @{ ok = $false; error = "excel_not_found" } 1
}

try {
    $wb = Resolve-TargetWorkbook -App $excel -WorkbookPath $WorkbookPath -WorkbookName $WorkbookName
} catch {
    Write-ResultAndExit @{ ok = $false; error = "$($_.Exception.Message)" } 1
}

try {
    Test-VbaTrustAccess -Workbook $wb | Out-Null
} catch {
    Write-ResultAndExit @{ ok = $false; error = "$($_.Exception.Message)" } 1
}

$vbproject = $wb.VBProject
if ($vbproject.Protection -ne 0) {
    Write-ResultAndExit @{ ok = $false; error = "ERR_WORKBOOK_PROTECTED"; workbook = $wb.Name } 1
}

if ($ModuleType) {
    # --- 新規モジュール作成モード（$ModuleType指定時のみ。既存の上書きパスとは別経路） ---
    $existing = $null
    try { $existing = $vbproject.VBComponents.Item($ModuleName) } catch {}
    if ($existing) {
        Write-ResultAndExit @{ ok = $false; error = "module_already_exists"; module = $ModuleName } 1
    }

    $newType = if ($ModuleType -eq "class") { 2 } else { 1 }  # vbext_ComponentType: 1=StdModule, 2=Class
    try {
        $newComp = $vbproject.VBComponents.Add($newType)
        $newComp.Name = $ModuleName
    } catch {
        Write-ResultAndExit @{ ok = $false; error = "create_failed"; module = $ModuleName; detail = "$($_.Exception.Message)" } 1
    }

    $sourceText = Get-Content -LiteralPath $SourceCodePath -Raw -Encoding UTF8
    try {
        $newComp.CodeModule.AddFromString($sourceText)
    } catch {
        Write-ResultAndExit @{ ok = $false; error = "create_failed"; module = $ModuleName; detail = "$($_.Exception.Message)" } 1
    }

    Write-ResultAndExit @{ ok = $true; workbook = $wb.Name; module = $ModuleName; componentType = $newType; created = $true } 0
}

$targetComp = $null
try {
    $targetComp = $vbproject.VBComponents.Item($ModuleName)
} catch {
    Write-ResultAndExit @{ ok = $false; error = "module_not_found"; module = $ModuleName } 1
}

# vbext_ComponentType: 1=StdModule, 2=Class, 3=MSForm, 100=Document(Sheet/ThisWorkbook)
$compType = $targetComp.Type

# 書き込み前に、現在のコードをタイムスタンプ付きでバックアップする（復元用の保険）
# 注意: CodeModule.Lines() はAttribute行を含まないため、バックアップもテキスト内容のみの復元用となる
try {
    $backupDir = Join-Path (Split-Path -Parent $vbproject.FileName) ".excel-vba-sync-backups"
    New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    $cm = $targetComp.CodeModule
    if ($cm.CountOfLines -gt 0) {
        $currentCode = $cm.Lines(1, $cm.CountOfLines)
    } else {
        $currentCode = ""
    }
    $backupExt = if ($compType -eq 2 -or $compType -eq 100) { ".cls" } elseif ($compType -eq 3) { ".frm" } else { ".bas" }
    $backupPath = Join-Path $backupDir ("{0}_{1}{2}" -f $ModuleName, $ts, $backupExt)
    [System.IO.File]::WriteAllText($backupPath, $currentCode, (New-Object System.Text.UTF8Encoding($false)))
} catch {
    # バックアップ失敗は書き込み処理自体を止めない（ベストエフォート）
    $backupPath = $null
}

# --- UserForm(Type=3)のコード上書き ---
# Designer側（コントロールの配置・サイズ、.frxバイナリ）には一切触れず、CodeModuleのコード
# だけを差し替える。Documentモジュール(Type=100)と同じDeleteLines+AddFromString経路。
# VBComponents.Import()を使う既存経路（.frm/.frxのペアが必要）は通さないため、ここで完結させる。
# 実機で確認済み: この経路でもDesignerのコントロール群と、モジュールレベルのAttribute行
# （VB_PredeclaredId=True等、UserFormの既定インスタンスに必須）は保持される。
# なおUserFormにはプロシージャレベルのAttribute行（ショートカットキー割り当て）が
# そもそも存在しないため、Documentモジュールのようなショートカット消失の警告は不要。
if ($compType -eq 3) {
    $formSourceText = Get-Content -LiteralPath $SourceCodePath -Raw -Encoding UTF8
    try {
        $formCm = $targetComp.CodeModule
        if ($formCm.CountOfLines -gt 0) { $formCm.DeleteLines(1, $formCm.CountOfLines) }
        $formCm.AddFromString($formSourceText)
    } catch {
        Write-ResultAndExit @{ ok = $false; error = "import_failed"; module = $ModuleName; workbook = $wb.Name; backupPath = $backupPath; detail = "$($_.Exception.Message)" } 1
    }
    Write-ResultAndExit @{ ok = $true; workbook = $wb.Name; module = $ModuleName; componentType = $compType; backupPath = $backupPath } 0
}

# 対象ワークブック名フォルダ配下に、対象拡張子でコードを配置する
# （Import-ModuleToVBProjectの「親フォルダ名 == ブック名」チェックを満たすため）
$bookName = [System.IO.Path]::GetFileNameWithoutExtension($vbproject.FileName)
$ext = if ($compType -eq 2) { ".cls" } else { ".bas" }   # Type=100(Document)は.clsと同様にDeleteLines経路を通るため.clsで統一

$tmpRoot = Join-Path $env:TEMP ("VBAImport_" + [Guid]::NewGuid())
$tmpBookDir = Join-Path $tmpRoot $bookName
New-Item -ItemType Directory -Force -Path $tmpBookDir | Out-Null
$tmpCodePath = Join-Path $tmpBookDir ($ModuleName + $ext)

$sourceText = Get-Content -LiteralPath $SourceCodePath -Raw -Encoding UTF8
[System.IO.File]::WriteAllText($tmpCodePath, $sourceText, (New-Object System.Text.UTF8Encoding($false)))

$existingModuleNames = @()
foreach ($comp in $vbproject.VBComponents) { $existingModuleNames += $comp.Name }

$fileItem = Get-Item -LiteralPath $tmpCodePath
# Import-ModuleToVBProject は内部で Write-Host による進捗ログを大量に出力するが、
# それが標準出力に混じると呼び出し元(server.ts)のJSON抽出が壊れるため、
# ストリーム6(Information)をここでまとめて捨てる
$ok = & { Import-ModuleToVBProject -vbproject $vbproject -file $fileItem -existingModuleNames $existingModuleNames } 6>$null

if (-not $ok) {
    Write-ResultAndExit @{ ok = $false; error = "import_failed"; module = $ModuleName; workbook = $wb.Name; backupPath = $backupPath } 1
}

$result = @{ ok = $true; workbook = $wb.Name; module = $ModuleName; componentType = $compType; backupPath = $backupPath }
if ($compType -eq 100) {
    $result.warning = "This is a Document module (Sheet/ThisWorkbook code-behind). Per-procedure Attribute lines (e.g. macro shortcut key assignments) cannot be preserved through this write path due to a VBA API constraint (VBComponents.Import is not available for Document-type components)."
}
Write-ResultAndExit $result 0
