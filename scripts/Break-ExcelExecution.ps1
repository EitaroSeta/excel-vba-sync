<#
    File: Break-ExcelExecution.ps1
    Description: Interrupt a running/stuck VBA macro by sending Ctrl+Break to Excel's
                 window at the Windows INPUT level (user32 keybd_event), deliberately
                 NOT through COM -- while Excel is executing VBA it refuses COM calls,
                 which is exactly the situation this script exists for.

                 Deliberately does NOT report "Excel is fine now". Measured behaviour:
                 a successful Ctrl+Break leaves VBA's modal "Code execution has been
                 interrupted" dialog on screen, and until a human clicks End the VBA
                 project stays in break mode -- every Application.Run fails with
                 0x800ADF09 and project edits raise "project will be reset". Read-only
                 COM calls keep answering throughout, so probing COM says nothing about
                 whether the macro stopped. An earlier version probed COM and reported
                 excelResponsive=true while VBA was wedged in exactly that state.
    Author: Eitaro SETA
    License: MIT License
    Copyright (c) 2025 Eitaro SETA
#>

param(
    [int]$TargetPid = 0,
    [switch]$AlsoSendEsc
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)

function Write-Json {
    param([hashtable]$Obj, [int]$ExitCode = 0)
    $Obj | ConvertTo-Json -Compress -Depth 5 | Write-Output
    exit $ExitCode
}

Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class BreakNative {
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, IntPtr lpdwProcessId);
    [DllImport("user32.dll")] public static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);
    [DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();
    [DllImport("user32.dll")] public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
}
"@

# --- locate the target Excel window ---
$procs = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 })
if ($TargetPid -gt 0) {
    $procs = @($procs | Where-Object { $_.Id -eq $TargetPid })
}
if ($procs.Count -eq 0) {
    Write-Json @{ ok = $false; error = 'no_excel_window'; detail = 'No EXCEL.EXE with a main window was found. Excel may not be running, or it has no visible window to receive input.' } 3
}
# Several Excel processes can coexist (see the multi-process caveat in AI_USAGE); without an
# explicit pid there is no way to tell which one is stuck, so report them and let the caller pick.
if ($procs.Count -gt 1 -and $TargetPid -le 0) {
    Write-Json @{ ok = $false; error = 'multiple_excel_processes'; pids = @($procs | ForEach-Object { $_.Id }); detail = 'More than one EXCEL.EXE is running. Re-call with processId set to the stuck one.' } 4
}

$proc = $procs[0]
$hwnd = $proc.MainWindowHandle

# --- bring Excel to the foreground ---
# keybd_event injects into the FOREGROUND window's input queue, so activation is mandatory.
# A bare SetForegroundWindow from a background process is refused by Windows' foreground lock
# (measured: windowActivated=false, the keystroke landed on the wrong window, and the macro
# kept running). Attaching this thread's input state to the current foreground thread first
# makes this process a legitimate caller, which fixed it (measured: windowActivated=true).
$fg = [BreakNative]::GetForegroundWindow()
$fgThread = [BreakNative]::GetWindowThreadProcessId($fg, [IntPtr]::Zero)
$myThread = [BreakNative]::GetCurrentThreadId()
$attached = $false
if ($fgThread -ne 0 -and $fgThread -ne $myThread) {
    $attached = [BreakNative]::AttachThreadInput($myThread, $fgThread, $true)
}
try {
    if ([BreakNative]::IsIconic($hwnd)) { [void][BreakNative]::ShowWindow($hwnd, 9) }  # SW_RESTORE
    [void][BreakNative]::BringWindowToTop($hwnd)
    [void][BreakNative]::SetForegroundWindow($hwnd)
} finally {
    if ($attached) { [void][BreakNative]::AttachThreadInput($myThread, $fgThread, $false) }
}
Start-Sleep -Milliseconds 400
$activated = ([BreakNative]::GetForegroundWindow() -eq $hwnd)

# --- send Ctrl+Break (VK_CONTROL 0x11 + VK_CANCEL 0x03) ---
# Sent repeatedly: VBA only samples the cancel key between statements, so a single press can
# fall in a gap. Cheap redundancy on a rescue path.
$VK_CONTROL = 0x11
$VK_CANCEL  = 0x03
$VK_ESCAPE  = 0x1B
$KEYUP      = 0x2
$sent = @()

for ($k = 0; $k -lt 3; $k++) {
    [BreakNative]::keybd_event($VK_CONTROL, 0, 0,      [UIntPtr]::Zero)
    [BreakNative]::keybd_event($VK_CANCEL,  0, 0,      [UIntPtr]::Zero)
    Start-Sleep -Milliseconds 80
    [BreakNative]::keybd_event($VK_CANCEL,  0, $KEYUP, [UIntPtr]::Zero)
    [BreakNative]::keybd_event($VK_CONTROL, 0, $KEYUP, [UIntPtr]::Zero)
    Start-Sleep -Milliseconds 150
}
$sent += 'Ctrl+Break x3'

if ($AlsoSendEsc) {
    Start-Sleep -Milliseconds 300
    [BreakNative]::keybd_event($VK_ESCAPE, 0, 0,      [UIntPtr]::Zero)
    [BreakNative]::keybd_event($VK_ESCAPE, 0, $KEYUP, [UIntPtr]::Zero)
    $sent += 'Esc'
}

if (-not $activated) {
    Write-Json @{
        ok             = $false
        error          = 'excel_not_activated'
        targetPid      = $proc.Id
        threadAttached = $attached
        sent           = $sent
        detail         = 'Excel could not be brought to the foreground, so the keystroke went to whatever window was focused and the macro is almost certainly still running. Windows blocks focus changes from background processes under some conditions (a full-screen app, a screen lock, an elevated window, or an active remote session). Ask the user to click on the Excel window and press Ctrl+Break themselves.'
    } 5
}

Write-Json @{
    ok                 = $true
    targetPid          = $proc.Id
    windowActivated    = $true
    threadAttached     = $attached
    sent               = $sent
    macroStopped       = 'unknown'
    userActionRequired = 'Excel is now almost certainly showing VBA''s modal dialog "Code execution has been interrupted". Tell the user to click End -- NOT Continue, which resumes the macro. Until they click End the VBA project stays in break mode: every macro run fails with 0x800ADF09 and any module write prompts "this action will reset the project". Once they click End, Excel is fully usable again and they can SAVE their work.'
    note               = 'This tool cannot confirm the macro actually stopped. Read-only COM calls answer normally even while VBA is wedged, so there is nothing here worth probing; ask the user what Excel is showing instead.'
} 0
