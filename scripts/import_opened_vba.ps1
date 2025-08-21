<#
    File: export_opened_vba.ps1
    Description: VBA���W���[����VSCode����VBA�ɃC���|�[�g����X�N���v�g
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
    [string]$InputFile  #�I�v�V�����i�P��t�@�C��/�t�H���_�j
)

# ���P�[���擾
$locale = (Get-UICulture).Name.Split('-')[0]
$defaultLocale = "en"
$localePath = Join-Path $PSScriptRoot "..\locales\$locale.json"

if (-Not (Test-Path $localePath)) {
    # �w�胍�P�[�����Ȃ���Ήp��Ƀt�H�[���o�b�N
    $localePath = Join-Path $PSScriptRoot "..\locales\$defaultLocale.json"
}

$messages = Get-Content $localePath -Encoding UTF8 -Raw | ConvertFrom-Json

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

if (-not $InputFile) {
    Write-Host ($messages."import.error.noImportFolder")
    exit 1
}

if ($InputFile -and -not (Test-Path $InputFile)) {
    Write-Host ($messages."import.error.invalidImportFolder" -f $InputFile)
    exit 2
}

# �P�ꃂ�W���[����VBProject��荞�ݏ���
function Import-ModuleToVBProject {
    param (
        $vbproject,          #�Ώۂ�VBProject
        $file,               #�C���|�[�g����t�@�C���i�t���p�X *.bas�@*.cls�j
        $existingModuleNames #�C���|�[�g��̃��W���[�����ꗗ
    )
    $importPath = $file.FullName
    $modName = $file.BaseName

    # ���W���[���̐e�f�B���N�g��������Ώۃu�b�N�����擾
    $parentDirName = Split-Path (Split-Path $importPath -Parent) -Leaf
    $bookName = [System.IO.Path]::GetFileNameWithoutExtension($vbproject.FileName)

    # �C���|�[�g����t�@�C���̐e�t�H���_�����u�b�N���ƈقȂ�Ȃ�X�L�b�v�i�������W���[���΍�j
    if ($parentDirName -ne $bookName) {
        #Write-Host "�� ���W���[��[$modName] �̊i�[�t�H���_[$parentDirName] �ƃC���|�[�g��u�b�N��[$bookName] ���قȂ邽�߃X�L�b�v���܂�"
        Write-Host ($messages."import.warn.skipDifferentFolder" -f $modName, $parentDirName, $bookName)
        return $false
    }

    # Excel�u�b�N�ɑ��݂��Ȃ����W���[���̓X�L�b�v
    if (-not ($existingModuleNames -contains $modName)) {
        #Write-Host "�� ���W���[���F$modName �� Excel �u�b�N $($vbproject.FileName) �ɑ��݂��܂���B�X�L�b�v���܂�"
        Write-Host ($messages."import.warn.moduleNotFound" -f $modName, $vbproject.FileName)
        return
    }
    else {
        #Write-Host "�� ���W���[���F$modName �� Excel �u�b�N $($vbproject.FileName) �փC���|�[�g���܂�"
        Write-Host ($messages."import.info.importModule" -f $modName, $vbproject.FileName)
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
            #Write-Host "�� (Doc) $modName �㏑���������܂���"
            Write-Host ($messages."import.info.moduleOverwriteComplete" -f $modName)
            return
        } else {
            $vbproject.VBComponents.Remove($targetComp)
            #Write-Host "�� $modName ���폜���čĒǉ����܂�"
            Write-Host ($messages."import.info.moduleRemoved" -f $modName)
        }
    } catch {}

    try {
        $newComp = $vbproject.VBComponents.Add(1)
        $newComp.Name = $modName
        $newComp.CodeModule.AddFromString($code)
        #Write-Host "�� $modName ��ǉ����܂���"
        Write-Host ($messages."import.info.moduleAdded" -f $modName)
        return $true
    } catch {
        #Write-Host "�� $modName �̒ǉ��Ɏ��s���܂���: $_"
        Write-Host ($messages."import.error.moduleAddFailed" -f $modName, $_)
        return $false
    }
}

# UTF-8 ������ �� SJIS(CP932) �Ńt�@�C���ۑ�
function Write-SJIS {
    param([string]$Path, [string]$Text)
    $sjis = [System.Text.Encoding]::GetEncoding(932)
    [System.IO.File]::WriteAllText($Path, $Text, $sjis)
}

# .frm �� UTF-8 �œǂ� �� SJIS �ɕϊ����Ĉꎞ�p�X��
function Convert-Frm-Utf8-ToSjisTemp {
    param([string]$frmPath)
    $base = [IO.Path]::GetFileNameWithoutExtension($frmPath)
    $frx  = [IO.Path]::ChangeExtension($frmPath, ".frx")
    if (-not (Test-Path $frx)) { throw $messages."import.error.frxNotFound" -f $frx }

    $tmp = Join-Path $env:TEMP ("VBAImport_" + [Guid]::NewGuid())
    New-Item -ItemType Directory -Force -Path $tmp | Out-Null

    $tmpFrm = Join-Path $tmp ($base + ".frm")
    $tmpFrx = Join-Path $tmp ($base + ".frx")

    # �����́g���b�Z�[�W��UTF-8�œǂށh���j�̂܂܁F.frm ����UTF-8�Ƃ��ēǂݎ��
    $text = Get-Content -LiteralPath $frmPath -Raw -Encoding UTF8
    Write-SJIS -Path $tmpFrm -Text $text
    Copy-Item -LiteralPath $frx -Destination $tmpFrx -Force

    return $tmpFrm
}

#�t�H�[����p�̎�荞�ݏ���
function Import-FormToVBProject {
    param(
        $filename,
        $vbproject,
        [string]$frmPath
    )

    $base = [System.IO.Path]::GetFileNameWithoutExtension($frmPath)
    $frx  = [System.IO.Path]::ChangeExtension($frmPath, ".frx")

    # Form��SJIS�ň���
    $tmpFrm = Convert-Frm-Utf8-ToSjisTemp -frmPath $frmPath

    if (-not (Test-Path $frx)) {
        Write-Host ($messages."import.error.frxNotFound2" -f $base , $frx )
        return $false
    }

    # �\�Ȃ�f�U�C�i�����i�J���Ă���Ɖ��邱�Ƃ�����j
    foreach ($w in $vbproject.VBE.Windows) {
        if ($w.Caption -like "*$base*") {
            try { $w.Close() } catch {}
        }
    }

    # ���������t�H�[�����폜
    try {
        $existing = $vbproject.VBComponents.Item($base)
        $vbproject.VBComponents.Remove($existing)
        Write-Host ($messages."import.info.frmModuleRemoved" -f $base)
    } catch {}

    # Import�i .frx �������Ŏ�荞�܂��j
    try {
        $comp = $vbproject.VBComponents.Import($tmpFrm)

        # ��荞�݌��ʌ��؁i3=UserForm�j
        if ($comp.Type -ne 3) {
            throw $messages."import.error.frxImportFailed" -f $($comp.Type)
        }

        try { $comp.Name = $base } catch {}
        Write-Host ($messages."import.info.frmModuleImportCompleted" -f $base,  $frmPath, $filename, $comp.Type)
        return $true

    } catch {
        Write-Host ($messages."import.error.frmModuleImportFailed" -f $base, $_)
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
    Write-Host ($messages."common.error.noExcel")
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
    Write-Host ($messages."common.error.noSavedWorkbook")
    exit 4
}

$originalAutoSave = @{}
foreach ($wb in $workbooks) {
    try {
        $originalAutoSave[$wb.Name] = $wb.AutoSaveOn
        $wb.AutoSaveOn = $false
    } catch {}
}

# �P��t�@�C�����ǂ����𔻒�
if ($InputFile){
    $item = Get-Item $InputFile
    if ($item.PSIsContainer) {
        #�t�H���_
        $IsSingleFile = $false
    } else {
        #�t�@�C���ł��B
        $IsSingleFile = $true
    }
}
else{
    $IsSingleFile = $false
}

$anySuccess = $false
$i = 1

foreach ($wb in $workbooks) {
    $vbproject = $wb.VBProject
    if ($vbproject.Protection -ne 0) {
       Write-Host ($messages."import.warn.protectedWorkbook" -f $i , $wb.Name)
        $i++
        continue
    }
    else{
        Write-Host ($messages."import.info.operationPossible" -f $i , $wb.Name)
    }

    $bookName = [System.IO.Path]::GetFileNameWithoutExtension($wb.Name)
    $bookDir = if ($IsSingleFile) {
        Split-Path $InputFile -Parent
    } else {

        $InputFile
    }

    if (-not (Test-Path $bookDir)) {
        Write-Host ($messages."import.warn.importDirNotFound" -f $i, $bookDir)
        $i++
        continue
    }else {
        Write-Host ($messages."import.info.importDirChecked" -f $i, $bookDir)
    }

    #�t�@�C���̏ꍇ�ƃt�H���_�̏ꍇ�Ńt�@�C�����擾�����𕪂���
    $targetFiles = if ($IsSingleFile) {
        Get-Item $InputFile
    } else {
        Get-ChildItem -Path $InputFile -Include *.bas, *.cls, *.frm  -Recurse
    }

    Write-Host ($messages."import.info.targetFiles" -f $i, ($targetFiles -join ", "))

    # �C���|�[�g���EXCEL-VBA���W���[�����ꗗ�����O�Ɏ擾
    $existingModuleNames = @()
    foreach ($comp in $vbproject.VBComponents) {
      $existingModuleNames += $comp.Name
    }
    Write-Host ($messages."import.info.existingModules" -f $i, ($existingModuleNames -join ", "))

    # �C���|�[�g���̃t�@�C���P�ʂŃ��[�v
    foreach ($file in $targetFiles) {
        $ext = [System.IO.Path]::GetExtension($file.FullName).ToLowerInvariant()

        #.frx/.frm
        if ($ext -eq ".frm") {
          $ok = Import-FormToVBProject -filename $wb.Name -vbproject $vbproject -frmPath $file.FullName
          if ($ok) { $anySuccess = $true }
          continue
        }

        #.bas/.cls/ThisWorkbook/Sheet
        if ($ext -in ".bas", ".cls") {
          $result = Import-ModuleToVBProject -vbproject $vbproject -file $file -existingModuleNames $existingModuleNames
          if ($result) {
            $anySuccess = $true
          }
        }
    }

    # �����ۑ��𕜌�
    if ($originalAutoSave.ContainsKey($wb.Name)) {
        try {
            $wb.AutoSaveOn = $originalAutoSave[$wb.Name]
        } catch {}
    }
    $i++
}

if (-not $anySuccess) {
    Write-Host ($messages."import.error.importFailedOrNoTarget")
    exit 5
}

Write-Host ($messages."commoninfo.importCompleted")
exit 0
