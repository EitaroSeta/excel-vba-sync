<#
    File: export_opened_vba.ps1
    Description: VBAモジュールをファイルにエクスポートするスクリプト
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
    [string]$OutputDir,
    [string]$BookName,
    [string]$ModuleName
)
write-host "----------------------------------------"
write-host "Output Directory: $OutputDir"
write-host "Book Name: $BookName"
write-host "Module Name: $ModuleName"
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
#[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$InformationPreference = 'Continue'
function Get-Timestamp {
    Get-Date -Format 'yyyy-MM-dd HH:mm:ss'  # → [YYYY-MM-DD HH:mm:ss]
}

# 引数チェック
if (-not $OutputDir) {
    #Write-Host ($messages."common.error.noPath")
    # タイムスタンプ付きでエラー出力（拡張側で *>&1 していれば拾う）
    $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $messages.'common.error.noPath'
    Write-host $msg
    exit 1
}

# OneDrive 配下なら中止
if ($OutputDir -like "$env:OneDrive*") {
    #Write-host ($messages."common.error.oneDriveFolder")
    $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $messages.'common.error.oneDriveFolder'
    Write-host $msg
    exit 2
}

# エクスポート先フォルダ
#Write-Host ($messages."export.info.exportFolderName" -f $OutputDir)
$msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $messages.'export.info.exportFolderName' -f $OutputDir
Write-host $msg

# 出力フォルダの存在確認
if (-not (Test-Path $OutputDir)) {
    #New-Item -ItemType Directory -Path $OutputDir | Out-Null
    #Write-Host ($messages."export.error.invalidFolder" -f $OutputDir)
    $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $messages.'export.error.invalidFolder' -f $OutputDir
    Write-host $msg
    exit 5
}

# Excelの既存インスタンスを取得（Excelが開かれていないと失敗する）
try {
    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
} catch {
    #Write-Host $messages."common.error.noExcel"
    $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $messages.'common.error.noExcel'
    Write-host $msg
    exit 3
}

# Excelウィンドウを明示的にアクティブにする
$excel.Visible = $true
[void] $excel.Windows.Item(1).Activate() 
Start-Sleep -Milliseconds 300

# VBE を可視化・プロジェクトウィンドウにフォーカス
$excel.VBE.MainWindow.Visible = $true
foreach ($window in $excel.VBE.Windows) {
    if ($window.Caption -like "*Project*") {
        $window.SetFocus() 
        #Write-Host $messages."export.info.excelFocus"
        $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $messages.'export.info.excelFocus'
        Write-host $msg
        break
    }
}

# 保存済みブックのみ対象
$workbooks = @()
for ($i = 1; $i -le $excel.Workbooks.Count; $i++) {
    $wb = $excel.Workbooks.Item($i)
    if ($wb.Path -ne "") {
        $ext = [System.IO.Path]::GetExtension($wb.Name).ToLower()
        if ($ext -eq ".xlsm" -or $ext -eq ".xlsb") {
          $workbooks += $wb
        } else {
            # Skip non-macro files
            $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $messages.'export.warn.nonMacroFileSkipped' -f $wb.Name
            Write-host $msg
        }
    }
}

if ($workbooks.Count -eq 0) {
    #Write-host $messages."common.error.noSavedWorkbook"
    $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $messages.'common.error.noSavedWorkbook'
    Write-host $msg
    exit 4
}

# 自動保存の状態を保存し、一時的に無効化
$originalAutoSave = @{}
foreach ($wb in $workbooks) {
    try {
        $originalAutoSave[$wb.Name] = $wb.AutoSaveOn
        if ($wb.AutoSaveOn) {
            $wb.AutoSaveOn = $false
            Write-Host ($messages."export.info.autoSaveCanceled" -f $wb.Name)
        }
    } catch {
        #Write-Host ($messages."export.error.autoSaveCancelFailed" -f $wb.Name)
        $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $messages.'export.error.autoSaveCancelFailed' -f $wb.Name
        Write-host $msg
    }
}

# 各ワークブックからモジュールをエクスポート
function ExportModule {
    param (
        [string]$OutputDir,
        [string]$BookName,
        [string]$ModuleName,
        [object]$workbooks
    )
    
    $success = $false # 全体の成功状態を管理

    foreach ($wb in $workbooks) {
        $project = $wb.VBProject
        $currentBookName = [System.IO.Path]::GetFileNameWithoutExtension($wb.Name)

        # ブック名が指定されている場合、対象のブックを絞り込む
        if ($BookName -and $currentBookName -ne $BookName) {
            continue
        }

        $bookDir = Join-Path $OutputDir $currentBookName
        if (-not (Test-Path $bookDir)) {
            New-Item -ItemType Directory -Path $bookDir | Out-Null
        }

        if ($project.Protection -ne 0) {
            $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $messages.'export.warn.protectedVBProject' -f $wb.Name
            Write-host $msg
            continue
        }

        $moduleProcessed = $false  # モジュールが処理されたかを管理するフラグ

        foreach ($component in $project.VBComponents) {
            $name = $component.Name
            # モジュール名が指定されている場合、対象のモジュールを絞り込む
            if ($ModuleName -and $name -ne $ModuleName) {
                continue
            }

            # モジュールが処理される場合にフラグを更新
            $moduleProcessed = $true

            switch ($component.Type) {
                1 { $ext = ".bas" }   # 標準モジュール
                2 { $ext = ".cls" }   # クラスモジュール
                3 { $ext = ".frm" }   # ユーザーフォーム
                100 { $ext = ".bas" } # ThisWorkbook / Sheet モジュール
                default { $ext = ".txt" }
            }

            $filename = Join-Path $bookDir "$name$ext"

            # export 
            $componentSuccess = ExportComponent -component $component -filename $filename
            if ($componentSuccess) {
              $success = $true  # 1件でも成功があれば全体を成功とする
            } #else {
              #$success = $false  # 失敗した場合は全体を失敗とする
              #$msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $messages.'export.error.exportFailedModule' 
              #Write-Host $msg
            #}

        }
        # すべてのモジュールがスキップされた場合の処理
        if (-not $moduleProcessed) {
            $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $messages.'export.warn.noModulesProcessed'
            Write-Host $msg
            #$success = $false
        }

    }
    return $success
}

function ExportComponent {
    param (
        [object]$component,
        [string]$filename
    )
    $success = $false

    # Type=100（ThisWorkbookやSheet）は .Export() 不安定な為、Linesで処理してcontinue
    if ($component.Type -eq 100) {
        try {
            $codeModule = $component.CodeModule
            $lineCount = $codeModule.CountOfLines
            if ($lineCount -gt 0) {
                $codeText = $codeModule.Lines(1, $lineCount)
                Set-Content -Path $filename -Value $codeText -Encoding UTF8
                $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $messages.'export.info.exportFallbackSuccess100' -f $filename
                Write-host $msg
                $success = $true
            } else {
                $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $messages.'export.warn.exportEmptyCode' -f $component.Name
                $success = $false
                Write-host $msg
            }
        } catch {
            $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $messages.'export.error.exportFailed100' -f $filename, $_
            $success = $false
            Write-host $msg
        }
        return $success
    }

    # コード有無チェック
    try {
        $codeModule = $component.CodeModule
        $lineCount = $codeModule.CountOfLines
        if ($lineCount -eq 0) {
            $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $messages.'export.warn.exportEmptyModule' -f $component.Name
            Write-host $msg
            $success = $false
            return $success
        }
        $codeText = $codeModule.Lines(1, $lineCount)
        if ($codeText -notmatch '\b(Sub|Function|Property)\b') {
            $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $messages.'export.warn.noCodeToExport' -f $component.Name
            Write-host $msg
            $success = $false
            return $success
        }
    } catch {
        $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $messages.'export.error.codeFetchFailed' -f $component.Name, $_
        Write-host $msg
        $success = $false
        return $success 
    }

    # .Export() 試行（最大3回）
    $success = $false
    for ($i = 1; $i -le 3; $i++) {

        try {
            $component.Activate() | Out-Null
            $component.CodeModule.CodePane.Show()
            Start-Sleep -Milliseconds 300

            $component.Export($filename)
            $success = $true

            # ★ 不要な Attribute 行を削除
            #$filtered = Get-Content $filename | Where-Object { $_ -notmatch '^Attribute VB_' }
            #Set-Content -Encoding UTF8 $filename -Value $filtered

            break
        } catch {
            Start-Sleep -Milliseconds 200
            $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $messages.'export.error.exportFailedModule' -f $i, $filename, $_
            $success = $false
            Write-host $msg
        }
    }

    if ($success) {
        $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $messages.'export.info.exportSuccess' -f $filename
        Write-host $msg
        return $success
    }

    # フォールバック: .Lines() による手動保存
    try {
        $codeModule = $component.CodeModule
        $lineCount = $codeModule.CountOfLines
        if ($lineCount -eq 0) {
            $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $messages.'export.warn.exportEmptyModule' -f $component.Name
            Write-host $msg
            $success = $false
            return $success
        }
        $codeText = $codeModule.Lines(1, $lineCount)
        Set-Content -Path $filename -Value $codeText -Encoding UTF8
        $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $messages.'export.info.exportFallbackSuccess' -f $filename
        Write-host $msg
        $success = $true
        return $success
    } catch {
        $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $messages.'export.error.exportFinalFailed' -f $filename
        Write-host $msg
    }
    return $success
}

write-host "----------------------------------------"

# 各ワークブックからモジュールをエクスポート
$anysuccess = ExportModule -OutputDir $OutputDir -BookName $BookName -ModuleName $ModuleName -workbooks $workbooks
write-host "----------------------------------------"

# 自動保存の設定を元に戻す
foreach ($wb in $workbooks) {
    if ($originalAutoSave.ContainsKey($wb.Name)) {
        $desiredState = $originalAutoSave[$wb.Name]
        try {
            if ($wb.AutoSaveOn -ne $desiredState) {
                $wb.AutoSaveOn = $desiredState
                #Write-Host ($messages."info.autoSaveRestored" -f $wb.Name, $desiredState)
                $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $messages.'info.autoSaveRestored' -f $wb.Name, $desiredState
                Write-host $msg
            } else {
                #Write-Host ($messages."export.info.autoSaveAlreadyRestored" -f $wb.Name)
                $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $messages.'export.info.autoSaveAlreadyRestored' -f $wb.Name
                Write-host $msg
            }
        } catch {
            #Write-Host ($messages."export.error.autoSaveRestoreFailed" -f $wb.Name, $_)
            $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $messages.'export.error.autoSaveRestoreFailed' -f $wb.Name, $_
            Write-host $msg
        }
    }
    
}

if (-not $anySuccess) {
    #Write-Host ($messages."import.error.exportFinalFailed")
    $msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), ($messages."export.error.exportFinalFailed")
    Write-host $msg
    write-host "----------------------------------------"
    exit 6
}
#Write-Host ($messages."export.info.exportModuleComplete" -f $OutputDir)
$msg = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $messages.'export.info.exportModuleComplete' -f $OutputDir
Write-Host  $msg
write-host "----------------------------------------"
exit 0
