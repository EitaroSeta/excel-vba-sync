<#
    File: export_opened_vba.ps1
    Description: VBA���W���[�����t�@�C���ɃG�N�X�|�[�g����X�N���v�g
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
    [string]$OutputDir
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

# UTF-8�ŏo�͂����悤�ݒ�
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# �����`�F�b�N
if (-not $OutputDir) {
    Write-Host $messages."common.error.noPath"
    exit 1
}

# OneDrive �z���Ȃ璆�~
if ($OutputDir -like "$env:OneDrive*") {
    Write-host ($messages."common.error.oneDriveFolder")
    exit 2
}

# �G�N�X�|�[�g��t�H���_
Write-Host ($messages."export.info.exportFolderName" -f $OutputDir)

# �o�̓t�H���_�̑��݊m�F
if (-not (Test-Path $OutputDir)) {
    #New-Item -ItemType Directory -Path $OutputDir | Out-Null
    Write-Host ($messages."export.error.invalidFolder" -f $OutputDir)
    exit 5
}

# Excel�̊����C���X�^���X���擾�iExcel���J����Ă��Ȃ��Ǝ��s����j
try {
    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
} catch {
    Write-Host $messages."common.error.noExcel"
    exit 3
}

# Excel�E�B���h�E�𖾎��I�ɃA�N�e�B�u�ɂ���
$excel.Visible = $true
$excel.Windows.Item(1).Activate()
Start-Sleep -Milliseconds 300

# VBE �������E�v���W�F�N�g�E�B���h�E�Ƀt�H�[�J�X
$excel.VBE.MainWindow.Visible = $true
foreach ($window in $excel.VBE.Windows) {
    if ($window.Caption -like "*Project*") {
        $window.SetFocus()
        Write-Host $messages."export.info.excelFocus"
        break
    }
}

# �ۑ��ς݃u�b�N�̂ݑΏ�
$workbooks = @()
for ($i = 1; $i -le $excel.Workbooks.Count; $i++) {
    $wb = $excel.Workbooks.Item($i)
    if ($wb.Path -ne "") {
        $workbooks += $wb
    }
}

if ($workbooks.Count -eq 0) {
    Write-host $messages."common.error.noSavedWorkbook"
    exit 4
}

# �����ۑ��̏�Ԃ�ۑ����A�ꎞ�I�ɖ�����
$originalAutoSave = @{}
foreach ($wb in $workbooks) {
    try {
        $originalAutoSave[$wb.Name] = $wb.AutoSaveOn
        if ($wb.AutoSaveOn) {
            $wb.AutoSaveOn = $false
            Write-Host ($messages."export.info.autoSaveCanceled" -f $wb.Name)
        }
    } catch {
        Write-Host ($messages."export.error.autoSaveCancelFailed" -f $wb.Name)
    }
}

# �e���[�N�u�b�N���烂�W���[�����G�N�X�|�[�g
foreach ($wb in $workbooks) {
    $project = $wb.VBProject
    $bookName = [System.IO.Path]::GetFileNameWithoutExtension($wb.Name)
    $bookDir = Join-Path $OutputDir $bookName
    if (-not (Test-Path $bookDir)) {
        New-Item -ItemType Directory -Path $bookDir | Out-Null
    }

    if ($project.Protection -ne 0) {
        Write-Host ($messages."export.warn.protectedVBProject" -f $wb.Name)
        continue
    }

    foreach ($component in $project.VBComponents) {
        $name = $component.Name
        switch ($component.Type) {
            1 { $ext = ".bas" }   # �W�����W���[��
            2 { $ext = ".cls" }   # �N���X���W���[��
            3 { $ext = ".frm" }   # ���[�U�[�t�H�[��
            100 { $ext = ".bas" } # ThisWorkbook / Sheet ���W���[��
            default { $ext = ".txt" }
        }

        $filename = Join-Path $bookDir "$name$ext"

        # Type=100�iThisWorkbook��Sheet�j�� .Export() �s����ȈׁALines�ŏ�������continue
        if ($component.Type -eq 100) {
            try {
                $codeModule = $component.CodeModule
                $lineCount = $codeModule.CountOfLines
                if ($lineCount -gt 0) {
                    $codeText = $codeModule.Lines(1, $lineCount)
                    Set-Content -Path $filename -Value $codeText -Encoding Default
                    Write-Host ($messages."export.info.exportFallbackSuccess100" -f $filename)
                } else {
                    Write-Host ($messages."export.warn.exportEmptyCode" -f $name)
                }
            } catch {
                Write-Host ($messages."export.error.exportFailed100" -f $filename ,$_)
            }
            continue
        }

        # �R�[�h�L���`�F�b�N
        try {
            $codeModule = $component.CodeModule
            $lineCount = $codeModule.CountOfLines
            if ($lineCount -eq 0) {
                Write-Host ($messages."export.warn.exportEmptyModule" -f $name)
                continue
            }
            $codeText = $codeModule.Lines(1, $lineCount)
            if ($codeText -notmatch '\b(Sub|Function|Property)\b') {
                Write-Host ($messages."export.warn.noCodeToExport" -f $name)
                continue
            }
        } catch {
            Write-Host ($messages."export.error.codeFetchFailed" -f $name, $_)
            continue
        }

        # .Export() ���s�i�ő�3��j
        $success = $false
        for ($i = 1; $i -le 3; $i++) {
            try {
                $component.Activate() | Out-Null
                $component.CodeModule.CodePane.Show()
                Start-Sleep -Milliseconds 300
      
                $component.Export($filename)
                $success = $true

                # �� �s�v�� Attribute �s���폜
                $filtered = Get-Content $filename | Where-Object { $_ -notmatch '^Attribute VB_' }
                Set-Content -Encoding UTF8 $filename -Value $filtered

                break
            } catch {
                Start-Sleep -Milliseconds 200
                Write-Host ($messages."export.error.exportFailedModule" -f $i, $filename, $_)
            }
        }

        if ($success) {
            Write-Host ($messages."export.info.exportSuccess" -f $filename)
            continue
        }

        # �t�H�[���o�b�N: .Lines() �ɂ��蓮�ۑ�
        try {
            $codeModule = $component.CodeModule
            $lineCount = $codeModule.CountOfLines
            if ($lineCount -eq 0) {
                Write-Host ($messages."export.warn.exportEmptyModule" -f $name)
                continue
            }
            $codeText = $codeModule.Lines(1, $lineCount)
            Set-Content -Path $filename -Value $codeText -Encoding UTF8
            Write-Host ($messages."export.info.exportFallbackSuccess" -f $filename)

        } catch {
            Write-Host ($messages."export.error.exportFinalFailed" -f $filename)
        }
    }
}

# �����ۑ��̐ݒ�����ɖ߂�
foreach ($wb in $workbooks) {
    if ($originalAutoSave.ContainsKey($wb.Name)) {
        $desiredState = $originalAutoSave[$wb.Name]
        try {
            if ($wb.AutoSaveOn -ne $desiredState) {
                $wb.AutoSaveOn = $desiredState
                Write-Host ($messages."info.autoSaveRestored" -f $wb.Name, $desiredState)
            } else {
                Write-Host ($messages."export.info.autoSaveAlreadyRestored" -f $wb.Name)
            }
        } catch {
            Write-Host ($messages."export.error.autoSaveRestoreFailed" -f $wb.Name, $_)
        }
    }
}

Write-Host ($messages."export.info.exportModuleComplete" -f $OutputDir)

exit 0
