<#
    File: export_opened_vba.ps1
    Description: VBAモジュールをVSCodeからVBAにインポートするスクリプト
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

param (
    [string]$InputFile  #オプション（単一ファイル/フォルダ）
)

$OutputEncoding = [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)

write-host "----------------------------------------"
write-host "Input File: $InputFile"
# ロケール取得
$locale = (Get-UICulture).Name.Split('-')[0]
$defaultLocale = "en"
$localePath = Join-Path $PSScriptRoot "..\locales\$locale.json"

if (-Not (Test-Path $localePath)) {
    # 指定ロケールがなければ英語にフォールバック
    $localePath = Join-Path $PSScriptRoot "..\locales\$defaultLocale.json"
}

$messages = Get-Content $localePath -Encoding UTF8 -Raw | ConvertFrom-Json

$InformationPreference = 'Continue'
function Get-Timestamp {
    Get-Date -Format 'yyyy-MM-dd HH:mm:ss'  # → [YYYY-MM-DD HH:mm:ss]
}

if (-not $InputFile) {
    #Write-Host ($messages."import.error.noImportFolder")
    $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $messages.'common.error.noImportFolder'
    Write-host $msg
    exit 1
}

if ($InputFile -and -not (Test-Path $InputFile)) {
    #Write-Host ($messages."import.error.invalidImportFolder" -f $InputFile)
    $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), ($messages."import.error.invalidImportFolder" -f $InputFile)
    Write-host $msg
    exit 2
}

# 単一モジュールのVBProject取り込み処理
function Import-ModuleToVBProject {
    param (
        $vbproject,          #対象のVBProject
        $file,               #インポートするファイル（フルパス *.bas　*.cls）
        $existingModuleNames #インポート先のモジュール名一覧
    )
    $importPath = $file.FullName
    $modName = $file.BaseName

    # モジュールの親ディレクトリ名から対象ブック名を取得
    $parentDirName = Split-Path (Split-Path $importPath -Parent) -Leaf
    $bookName = [System.IO.Path]::GetFileNameWithoutExtension($vbproject.FileName)

    # インポートするファイルの親フォルダ名がブック名と異なるならスキップ（同名モジュール対策）
    if ($parentDirName -ne $bookName) {
        #Write-Host "■ モジュール[$modName] の格納フォルダ[$parentDirName] とインポート先ブック名[$bookName] が異なるためスキップします"
        #Write-Host ($messages."import.warn.skipDifferentFolder" -f $modName, $parentDirName, $bookName)
        $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), ($messages."import.warn.skipDifferentFolder" -f $modName, $parentDirName, $bookName)
        Write-host $msg
        return $false
    }

    # Excelブックに存在しないモジュールはスキップ
    if (-not ($existingModuleNames -contains $modName)) {
        #Write-Host "■ モジュール：$modName は Excel ブック $($vbproject.FileName) に存在しません。スキップします"
        #Write-Host ($messages."import.warn.moduleNotFound" -f $modName, $vbproject.FileName)
        $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), ($messages."import.warn.moduleNotFound" -f $modName, $vbproject.FileName)
        Write-host $msg
        return $false
    }
    else {
        #Write-Host "■ モジュール：$modName を Excel ブック $($vbproject.FileName) へインポートします"
        #Write-Host ($messages."import.info.importModule" -f $modName, $vbproject.FileName)
        $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), ($messages."import.info.importModule" -f $modName, $vbproject.FileName)
        Write-host $msg
    }

    $raw = Get-Content -Path $importPath -Encoding utf8
    $codeLines = $raw | Where-Object {
        ($_ -notmatch "^VERSION") -and
        ($_ -notmatch "^BEGIN") -and
        ($_ -notmatch "^END(\r?\n)?$") -and
        ($_ -notmatch "^Attribute VB_") -and
        ($_ -notmatch "^\s*MultiUse\s*=")
    }
    $code = $codeLines -join "`r`n"

    try {
        $targetComp = $vbproject.VBComponents.Item($modName)
        if ($targetComp.Type -eq 100) {
            $targetComp.CodeModule.DeleteLines(1, $targetComp.CodeModule.CountOfLines)
            $targetComp.CodeModule.AddFromString($code)
            #Write-Host "■ (Doc) $modName 上書き完了しました"
            #Write-Host ($messages."import.info.moduleOverwriteComplete" -f $modName)
            $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), ($messages."import.info.moduleOverwriteComplete" -f $modName)
            Write-host $msg
            return $true
        } else {
            $vbproject.VBComponents.Remove($targetComp)
            #Write-Host "■ $modName を削除して再追加します"
            #Write-Host ($messages."import.info.moduleRemoved" -f $modName)
            $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), ($messages."import.info.moduleRemoved" -f $modName)
            Write-host $msg
        }
    } catch {}

    if ($ext -eq ".cls") {
      $moduleType = 2 # クラスモジュール
    } elseif ($ext -eq ".bas") {
      $moduleType = 1 # 標準モジュール
    } else {
      $moduleType = $null
    }

    if ($moduleType -eq 1 -or $moduleType -eq 2) {

      try {
          #$newComp = $vbproject.VBComponents.Add(1)
          $newComp = $vbproject.VBComponents.Add($moduleType)
          $newComp.Name = $modName
          $newComp.CodeModule.AddFromString($code)
          #Write-Host "■ $modName を追加しました"
          #Write-Host ($messages."import.info.moduleAdded" -f $modName)
          $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), ($messages."import.info.moduleAdded" -f $modName)
          Write-host $msg
          return $true
      } catch {
          #Write-Host "■ $modName の追加に失敗しました: $_"
          #Write-Host ($messages."import.error.moduleAddFailed" -f $modName, $_)
          $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), ($messages."import.error.moduleAddFailed" -f $modName, $_)
          Write-host $msg
          return $false
      }

    }
}

# UTF-8 文字列 → SJIS(CP932) でファイル保存
function Write-SJIS {
    param([string]$Path, [string]$Text)
    $sjis = [System.Text.Encoding]::GetEncoding(932)
    [System.IO.File]::WriteAllText($Path, $Text, $sjis)
}

# .frm を UTF-8 で読み → SJIS に変換して一時パスへ
function Convert-Frm-Utf8-ToSjisTemp {
    param([string]$frmPath)
    $base = [IO.Path]::GetFileNameWithoutExtension($frmPath)
    $frx  = [IO.Path]::ChangeExtension($frmPath, ".frx")
    if (-not (Test-Path $frx)) { throw $messages."import.error.frxNotFound" -f $frx }

    $tmp = Join-Path $env:TEMP ("VBAImport_" + [Guid]::NewGuid())
    New-Item -ItemType Directory -Force -Path $tmp | Out-Null

    $tmpFrm = Join-Path $tmp ($base + ".frm")
    $tmpFrx = Join-Path $tmp ($base + ".frx")

    # ここは“メッセージはUTF-8で読む”方針のまま：.frm だけUTF-8として読み取り
    $text = Get-Content -LiteralPath $frmPath -Raw -Encoding UTF8
    Write-SJIS -Path $tmpFrm -Text $text
    Copy-Item -LiteralPath $frx -Destination $tmpFrx -Force

    return $tmpFrm
}

#フォーム専用の取り込み処理
function Import-FormToVBProject {
    param(
        $filename,
        $vbproject,
        [string]$frmPath,
        $existingModuleNames
    )

    $base = [System.IO.Path]::GetFileNameWithoutExtension($frmPath)
    $frx  = [System.IO.Path]::ChangeExtension($frmPath, ".frx")

    # モジュールの親ディレクトリ名から対象ブック名を取得
    $parentDirName = Split-Path (Split-Path $frmPath -Parent) -Leaf
    $bookName = [System.IO.Path]::GetFileNameWithoutExtension($vbproject.FileName)

    # インポートするファイルの親フォルダ名がブック名と異なるならスキップ（同名モジュール対策）
    if ($parentDirName -ne $bookName) {
        #Write-Host ($messages."import.warn.skipDifferentFolder" -f $base, $parentDirName, $bookName)
        $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), ($messages."import.warn.skipDifferentFolder" -f $base, $parentDirName, $bookName)
        Write-host $msg
        return $false
    }
    
    # Excelブックに存在しないモジュールはスキップ
    if (-not ($existingModuleNames -contains $base)) {
        #Write-Host ($messages."import.warn.moduleNotFound" -f $base, $vbproject.FileName)
        $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), ($messages."import.warn.moduleNotFound" -f $base, $vbproject.FileName)
        Write-host $msg
        return $false
    } else {
        #Write-Host ($messages."import.info.importModule" -f $base, $vbproject.FileName)
        $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), ($messages."import.info.importModule" -f $base, $vbproject.FileName)
        Write-host $msg
    }

    # FormはSJISで扱う
    $tmpFrm = Convert-Frm-Utf8-ToSjisTemp -frmPath $frmPath

    if (-not (Test-Path $frx)) {
        #Write-Host ($messages."import.error.frxNotFound2" -f $base , $frx )
        $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), ($messages."import.error.frxNotFound2" -f $base , $frx )
        Write-host $msg
        return $false
    }

    # 可能ならデザイナを閉じる（開いていると壊れることがある）
    foreach ($w in $vbproject.VBE.Windows) {
        if ($w.Caption -like "*$base*") {
            try { $w.Close() } catch {}
        }
    }

    # 既存同名フォームを削除
    try {
        $existing = $vbproject.VBComponents.Item($base)
        $vbproject.VBComponents.Remove($existing)
        #Write-Host ($messages."import.info.frmModuleRemoved" -f $base)
        $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), ($messages."import.info.frmModuleRemoved" -f $base)
        Write-host $msg

    } catch {}

    # Import（ .frx も自動で取り込まれる）
    try {
        $comp = $vbproject.VBComponents.Import($tmpFrm)

        # 取り込み結果検証（3=UserForm）
        if ($comp.Type -ne 3) {
            throw $messages."import.error.frxImportFailed" -f $($comp.Type)
        }

        try { $comp.Name = $base } catch {}
        #Write-Host ($messages."import.info.frmModuleImportCompleted" -f $base,  $frmPath, $filename, $comp.Type)
        $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), ($messages."import.info.frmModuleImportCompleted" -f $base,  $frmPath, $filename)
        Write-host $msg
        return $true

    } catch {
        #Write-Host ($messages."import.error.frmModuleImportFailed" -f $base, $_)
        $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), ($messages."import.error.frmModuleImportFailed" -f $base, $_)
        Write-host $msg
        return $false
    }
}

function Get-ComProp {
    param([object]$obj, [string]$prop)
    $obj.GetType().InvokeMember($prop, 'GetProperty', $null, $obj, $null)
}

try {
    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
} catch {
    #Write-Host ($messages."common.error.noExcel")
    $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), ($messages."common.error.noExcel")
    Write-host $msg
    exit 3
}

$excel.VBE.MainWindow.Visible = $true
foreach ($window in $excel.VBE.Windows) {
    if ($window.Caption -like "*Project*") {
        $window.SetFocus()
        break
    }
}

$workbooks = @()
for ($i = 1; $i -le $excel.Workbooks.Count; $i++) {
    $wb = $excel.Workbooks.Item($i)
    if ($wb.Path -ne "") {
        $workbooks += $wb
    }
}

if ($workbooks.Count -eq 0) {
    #Write-Host ($messages."common.error.noSavedWorkbook")
    $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), ($messages."common.error.noSavedWorkbook")
    Write-host $msg 
    exit 4
}

$originalAutoSave = @{}
foreach ($wb in $workbooks) {
    try {
        $originalAutoSave[$wb.Name] = $wb.AutoSaveOn
        $wb.AutoSaveOn = $false
    } catch {}
}

# 単一ファイルかどうかを判定
if ($InputFile){
    $item = Get-Item $InputFile
    if ($item.PSIsContainer) {
        #フォルダ
        $IsSingleFile = $false
    } else {
        #ファイルです。
        $IsSingleFile = $true
    }
}
else{
    $IsSingleFile = $false
}

$anySuccess = $false
$i = 1

write-host "----------------------------------------"
foreach ($wb in $workbooks) {
    $vbproject = $wb.VBProject
    if ($vbproject.Protection -ne 0) {
       #Write-Host ($messages."import.warn.protectedWorkbook" -f $i , $wb.Name)
       $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), ($messages."import.warn.protectedWorkbook" -f $i , $wb.Name)
       Write-host $msg
        $i++
        continue
    }
    else{
        #Write-Host ($messages."import.info.operationPossible" -f $i , $wb.Name)
        $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), ($messages."import.info.operationPossible" -f $i , $wb.Name)
    }

    $bookName = [System.IO.Path]::GetFileNameWithoutExtension($wb.Name)
    $bookDir = if ($IsSingleFile) {
        Split-Path $InputFile -Parent
    } else {

        $InputFile
    }

    if (-not (Test-Path $bookDir)) {
        #Write-Host ($messages."import.warn.importDirNotFound" -f $i, $bookDir)
        $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), ($messages."import.warn.importDirNotFound" -f $i, $bookDir)
        Write-host $msg
        $i++
        continue
    }else {
        #Write-Host ($messages."import.info.importDirChecked" -f $i, $bookDir)
        $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), ($messages."import.info.importDirChecked" -f $i, $bookDir)
        Write-host $msg
    }

    #ファイルの場合とフォルダの場合でファイル名取得処理を分ける
    $targetFiles = if ($IsSingleFile) {
        Get-Item $InputFile
    } else {
        Get-ChildItem -Path $InputFile -Include *.bas, *.cls, *.frm  -Recurse
    }

    #Write-Host ($messages."import.info.targetFiles" -f $i, ($targetFiles -join ", "))
    $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), ($messages."import.info.targetFiles" -f $i, ($targetFiles -join ", "))
    Write-host $msg

    # インポート先のEXCEL-VBAモジュール名一覧を事前に取得
    $existingModuleNames = @()
    foreach ($comp in $vbproject.VBComponents) {
      $existingModuleNames += $comp.Name
    }
    #Write-Host ($messages."import.info.existingModules" -f $i, ($existingModuleNames -join ", "))
    $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), ($messages."import.info.existingModules" -f $i, ($existingModuleNames -join ", "))
    Write-host $msg

    # インポート元のファイル単位でループ
    foreach ($file in $targetFiles) {
        $ext = [System.IO.Path]::GetExtension($file.FullName).ToLowerInvariant()

        #.frx/.frm
        if ($ext -eq ".frm") {
          $ok = Import-FormToVBProject -filename $wb.Name -vbproject $vbproject -frmPath $file.FullName -existingModuleNames $existingModuleNames
          if ($ok) {
            $anySuccess = $true
          }
          continue
        }

        #.bas/.cls/ThisWorkbook/Sheet
        if ($ext -in ".bas", ".cls") {
          $result = Import-ModuleToVBProject -vbproject $vbproject -file $file -existingModuleNames $existingModuleNames
          if ($result) {
            $anySuccess = $true
          }
          continue
        }
    }

    write-host "----------------------------------------"

    # 自動保存を復元
    if ($originalAutoSave.ContainsKey($wb.Name)) {
        try {
            $wb.AutoSaveOn = $originalAutoSave[$wb.Name]
        } catch {}
    }
    $i++
}

if (-not $anySuccess) {
    #Write-Host ($messages."import.error.importFailedOrNoTarget")
    $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), ($messages."import.error.importFailedOrNoTarget")
    Write-host $msg
    write-host "----------------------------------------"
    exit 5
}
#Write-Host ($messages."commoninfo.importCompleted")
$msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), ($messages."import.info.importCompleted")
Write-host $msg
write-host "----------------------------------------"
exit 0
