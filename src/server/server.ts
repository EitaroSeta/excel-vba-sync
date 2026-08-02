import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import * as fs from "node:fs";
import * as path from "node:path";
import * as os from "node:os";
import { createHash } from "node:crypto";
import { scanModuleForDependencies, ModuleDependencyScan } from "./dependencyScan.js";
import { scanModuleForReferences, ModuleReferenceScan } from "./referenceScan.js";
import { scanModuleForVariableScopes, ModuleVariableScopeScan, resolveVariableUsages } from "./variableScopeScan.js";
const execFileAsyncRaw = promisify(execFile);
// Serializes all Excel-COM-touching PowerShell invocations so concurrent MCP tool
// calls (e.g. an agent firing off excel_list_modules/excel_list_macros/excel_read_range
// "in parallel" against the same workbook) can't race each other into each
// independently deciding Excel isn't running yet and launching its own redundant
// Excel.exe process -- Get-OrStartExcelApplication's GetActiveObject check has no
// visibility into another in-flight launch by a sibling call. Queuing every
// execFileAsync call here means only one PowerShell/COM operation is ever actually
// running at a time, regardless of how many tool calls arrive concurrently.
let excelOpQueue: Promise<unknown> = Promise.resolve();
function execFileAsync(...args: any[]): Promise<any> {
  const run = excelOpQueue.then(
    () => (execFileAsyncRaw as any)(...args),
    () => (execFileAsyncRaw as any)(...args)
  );
  excelOpQueue = run.then(() => undefined, () => undefined);
  return run;
}

console.log("# vba-excel-mcp server: booting...");

const server = new McpServer(
  { name: "vba-excel-mcp", title: "Excel VBA Sync", version: "0.1.0" },
  {
    instructions:
      "Read, search, run macros in, and write to VBA modules of Excel workbooks on this Windows machine, via COM automation. " +
      "Prefer workbookPath (full file path) over workbook (display name) -- it auto-launches Excel and auto-opens the file if needed. " +
      "excel_update_module_code requires dryRun:true first to preview and get a confirmToken, then a second call with that token to actually write; " +
      "the token is rejected with ERR_MODULE_CHANGED_SINCE_DRYRUN if the module changed in the meantime. " +
      "A backup of the replaced code is always written before any change.",
  }
);
server.tool("ping", "Health check for the excel-vba-sync MCP server. Returns the literal string 'pong' if the server process is reachable. Does not touch Excel.", {}, async () => ({ content: [{ type: "text", text: "pong" }] }));

const transport = new StdioServerTransport();
server.connect(transport);

// 文字列の ' をエスケープ
function psq(s: string) { return s.replace(/'/g, "''"); }

// scripts/ フォルダの絶対パス（MCP_SCRIPTS_DIR優先、無ければ MCP_PS_LIST から逆算）
function getScriptsDir(): string | undefined {
  if (process.env.MCP_SCRIPTS_DIR) { return process.env.MCP_SCRIPTS_DIR; }
  if (process.env.MCP_PS_LIST) { return path.dirname(process.env.MCP_PS_LIST); }
  return undefined;
}

// ExcelUtil.ps1 を dot-source する行を生成（見つからない場合は空文字＝スキップ）
function dotSourceExcelUtil(): string {
  const dir = getScriptsDir();
  if (!dir) { return ""; }
  const utilPath = path.join(dir, "ExcelUtil.ps1");
  if (!fs.existsSync(utilPath)) { return ""; }
  return `. '${psq(utilPath)}'`;
}

// PowerShell からの JSON 出力を見て、エラー（ERR_ プレフィックス、または ok:false）なら
// isError:true として返す。注意: isError が立っていない（ok:true）からといって、
// マクロが「意図した通りに」動作したことまでは保証しない（例外を投げずに完走した、という意味に過ぎない）。
function classifyResult(outText: string): { content: { type: "text"; text: string }[]; isError?: boolean } {
  try {
    const start = Math.min(
      ...['{', '['].map(ch => { const i = outText.indexOf(ch); return i === -1 ? Number.POSITIVE_INFINITY : i; })
    );
    if (Number.isFinite(start)) {
      const payload = JSON.parse(outText.slice(start));
      if (payload && typeof payload.error === "string" && payload.error.startsWith("ERR_")) {
        return { content: [{ type: "text", text: outText }], isError: true };
      }
      if (payload && payload.ok === false) {
        return { content: [{ type: "text", text: outText }], isError: true };
      }
    }
  } catch { /* JSON以外の出力はそのまま返す */ }
  return { content: [{ type: "text", text: outText }] };
}

// execFileAsyncが非ゼロ終了コードで失敗した場合でも、e.stdout に実際のJSON出力
// （例: {ok:false, error:"macro not found", ...}）が残っていることが多いため、
// "ps failed" という汎用メッセージで実際のエラー内容を握りつぶさないようにする
function extractFailureResult(e: any): { content: { type: "text"; text: string }[]; isError: boolean } {
  const stdout = e?.stdout;
  if (stdout) {
    const outText = Buffer.isBuffer(stdout) ? stdout.toString("utf8") : String(stdout);
    if (outText.trim().length > 0) {
      const classified = classifyResult(outText);
      return { content: classified.content, isError: true };
    }
  }
  return { content: [{ type: "text", text: JSON.stringify({ ok: false, error: "ps_failed", detail: String(e?.message ?? e) }) }], isError: true };
}

// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ excel_get_module_code ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
server.tool(
  "excel_get_module_code",
  "Read the full source code of a VBA module (all Sub/Function bodies as VBE would show them; module- and procedure-level Attribute lines, e.g. macro shortcut key bindings, are NOT included -- this reads via CodeModule.Lines()). If the workbook is already open in Excel, 'workbook' (its display name) is enough. Otherwise pass 'workbookPath' (full file path): Excel will be launched if not running, and the file opened if not already open. Fails with ERR_VBOM_TRUST_DISABLED if Excel's 'Trust access to the VBA project object model' setting is off (cannot be enabled programmatically).",
  {
    workbook: z.string().describe("Workbook display name, e.g. 'Book1.xlsm'. Must match an already-open workbook unless workbookPath is also given."),
    module: z.string().describe("VBA module name (e.g. 'Module1', 'Sheet1', 'ThisWorkbook')."),
    workbookPath: z.string().optional().describe("Full path to the workbook file. If set, Excel is auto-launched and the file auto-opened when needed, instead of requiring it to already be open."),
  },
  async (params) => {
    const wb = psq(params.workbook);
    const mod = psq(params.module);
    const wbPath = psq(params.workbookPath ?? "");
    const dotSource = dotSourceExcelUtil();

    // PowerShell ワンライナーで COM 経由取得
    const psScript = `
$ErrorActionPreference='Stop'
# --- Force UTF-8 output (no BOM) ---
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$OutputEncoding           = [Console]::OutputEncoding

${dotSource}

try { $r = Get-OrStartExcelApplication; $excel = $r.App }
catch { @{ ok=$false; error='excel_not_found' } | ConvertTo-Json ; exit }

try { $wb = Resolve-TargetWorkbook -App $excel -WorkbookPath '${wbPath}' -WorkbookName '${wb}' }
catch { @{ ok=$false; error="$($_.Exception.Message)" } | ConvertTo-Json ; exit }

try { Test-VbaTrustAccess -Workbook $wb | Out-Null }
catch { @{ ok=$false; error="$($_.Exception.Message)" } | ConvertTo-Json ; exit }

try { $vbc=$wb.VBProject.VBComponents.Item('${mod}') }
catch { @{ ok=$false; error='module_not_found'; module='${mod}' } | ConvertTo-Json ; exit }

try {
  $cm=$vbc.CodeModule
  $code=$cm.Lines(1, $cm.CountOfLines)
  $res = @{ ok=$true; workbook=$wb.Name; module=$vbc.Name; lines=$cm.CountOfLines; code=$code }
  if ($r.LaunchedProcessId) { $res.launchedExcelPid = $r.LaunchedProcessId }
  $res | ConvertTo-Json -Depth 6
} catch {
  @{ ok=$false; error='read_failed'; detail="$($_.Exception.Message)" } | ConvertTo-Json
}
`.trim();

    try {
      const { stdout } = await execFileAsync(
        "powershell.exe",
        ["-NoLogo","-NoProfile","-NonInteractive","-STA","-ExecutionPolicy","Bypass","-Command", psScript],
        {
          windowsHide: true,
          encoding: "buffer",
          timeout: 20000,
          maxBuffer: 2 * 1024 * 1024,
        }
      );
      const outText  = Buffer.isBuffer(stdout) ? stdout.toString("utf8") : String(stdout);
      return classifyResult(outText);
    } catch (e: any) {
      return { content: [{ type: "text", text: JSON.stringify({ ok:false, error:"ps_failed", detail:String(e?.message ?? e) }) }] };
    }
  }
);

server.tool(
  "excel_list_modules",
  "List the VBA modules in a workbook (name, component type, and line count) without scanning any code content -- cheap and fast compared to vba_search_code. Use this instead of vba_search_code when you just need to know what modules exist, before deciding what to read/search/run. If the workbook is already open in Excel, 'workbook' (its display name) is enough. Otherwise pass 'workbookPath' (full file path): Excel will be launched if not running, and the file opened if not already open.",
  {
    workbook: z.string().optional().describe("Workbook display name. Either this or workbookPath is required; workbookPath is preferred since it also auto-launches/opens Excel if needed."),
    workbookPath: z.string().optional().describe("Full path to the workbook file. If set, Excel is auto-launched and the file auto-opened when needed."),
  },
  async (params) => {
    if (!params.workbook && !params.workbookPath) {
      return { content: [{ type: "text", text: JSON.stringify({ ok: false, error: "workbook or workbookPath is required" }) }], isError: true };
    }
    const wb = psq(params.workbook ?? "");
    const wbPath = psq(params.workbookPath ?? "");
    const dotSource = dotSourceExcelUtil();

    const psScript = `
$ErrorActionPreference='Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$OutputEncoding           = [Console]::OutputEncoding

${dotSource}

try { $r = Get-OrStartExcelApplication; $excel = $r.App }
catch { @{ ok=$false; error='excel_not_found' } | ConvertTo-Json ; exit }

try { $wb = Resolve-TargetWorkbook -App $excel -WorkbookPath '${wbPath}' -WorkbookName '${wb}' }
catch { @{ ok=$false; error="$($_.Exception.Message)" } | ConvertTo-Json ; exit }

try { Test-VbaTrustAccess -Workbook $wb | Out-Null }
catch { @{ ok=$false; error="$($_.Exception.Message)" } | ConvertTo-Json ; exit }

try {
  $mods = @()
  foreach ($c in @($wb.VBProject.VBComponents)) {
    $vbType = $c.Type   # 1:StdModule, 2:Class, 3:MSForm, 100:Document(Worksheet/ThisWorkbook)
    $ext = switch ($vbType) {
      1 { 'bas' }
      3 { 'frm' }
      default { 'cls' }
    }
    $lineCount = 0
    try { $lineCount = $c.CodeModule.CountOfLines } catch {}
    $mods += [pscustomobject]@{ name=$c.Name; componentType=$vbType; exportExt=$ext; lines=$lineCount }
  }
  $res = @{ ok=$true; workbook=$wb.Name; modules=$mods; count=$mods.Count }
  if ($r.LaunchedProcessId) { $res.launchedExcelPid = $r.LaunchedProcessId }
  $res | ConvertTo-Json -Depth 6
} catch {
  @{ ok=$false; error='list_failed'; detail="$($_.Exception.Message)" } | ConvertTo-Json
}
`.trim();

    try {
      const { stdout } = await execFileAsync(
        "powershell.exe",
        ["-NoLogo", "-NoProfile", "-NonInteractive", "-STA", "-ExecutionPolicy", "Bypass", "-Command", psScript],
        { windowsHide: true, encoding: "buffer", timeout: 20000, maxBuffer: 2 * 1024 * 1024 }
      );
      const outText = Buffer.isBuffer(stdout) ? stdout.toString("utf8") : String(stdout);
      return classifyResult(outText);
    } catch (e: any) {
      return extractFailureResult(e);
    }
  }
);

// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ excel_list_macros ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
server.tool(
  "excel_read_range",
  "Read cell values from a range in an Excel worksheet -- e.g. to verify what excel_run_macro actually did to the spreadsheet, since that tool alone cannot confirm the macro's effect. Does NOT require the VBA Trust Center setting (unlike the VBA code tools): this only touches the normal Excel object model (Range/Worksheet), not VBProject. Returns a row-major 2D array of values using Range.Value2 (avoids the Date/Currency wrapping that Range.Value can introduce). Large ranges (e.g. whole columns/rows) can be slow -- prefer a bounded address like 'A1:C10'. If the workbook is already open in Excel, 'workbook' (its display name) is enough. Otherwise pass 'workbookPath' (full file path): Excel will be launched if not running, and the file opened if not already open.",
  {
    workbook: z.string().optional().describe("Workbook display name. Either this or workbookPath is required; workbookPath is preferred since it also auto-launches/opens Excel if needed."),
    workbookPath: z.string().optional().describe("Full path to the workbook file. If set, Excel is auto-launched and the file auto-opened when needed."),
    sheet: z.string().describe("Worksheet name to read from."),
    range: z.string().describe("Cell range address, e.g. 'A1', 'A1:C10', or a defined name."),
  },
  async (params) => {
    if (!params.workbook && !params.workbookPath) {
      return { content: [{ type: "text", text: JSON.stringify({ ok: false, error: "workbook or workbookPath is required" }) }], isError: true };
    }
    const wb = psq(params.workbook ?? "");
    const wbPath = psq(params.workbookPath ?? "");
    const sheet = psq(params.sheet);
    const range = psq(params.range);
    const dotSource = dotSourceExcelUtil();

    const psScript = `
$ErrorActionPreference='Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$OutputEncoding           = [Console]::OutputEncoding

${dotSource}

try { $r = Get-OrStartExcelApplication; $excel = $r.App }
catch { @{ ok=$false; error='excel_not_found' } | ConvertTo-Json ; exit }

try { $wb = Resolve-TargetWorkbook -App $excel -WorkbookPath '${wbPath}' -WorkbookName '${wb}' }
catch { @{ ok=$false; error="$($_.Exception.Message)" } | ConvertTo-Json ; exit }

try { $ws = $wb.Worksheets.Item('${sheet}') }
catch { @{ ok=$false; error='sheet_not_found'; sheet='${sheet}' } | ConvertTo-Json ; exit }

try { $rng = $ws.Range('${range}') }
catch { @{ ok=$false; error='invalid_range'; range='${range}' } | ConvertTo-Json ; exit }

try {
  $rowCount = $rng.Rows.Count
  $colCount = $rng.Columns.Count
  $raw = $rng.Value2
  $data = New-Object 'System.Collections.ArrayList'
  for ($rIdx = 1; $rIdx -le $rowCount; $rIdx++) {
    $rowData = New-Object 'System.Collections.ArrayList'
    for ($cIdx = 1; $cIdx -le $colCount; $cIdx++) {
      if ($rowCount -eq 1 -and $colCount -eq 1) {
        $cellVal = $raw
      } else {
        $cellVal = $raw[$rIdx, $cIdx]
      }
      [void]$rowData.Add($cellVal)
    }
    [void]$data.Add($rowData)
  }
  $res = @{ ok=$true; workbook=$wb.Name; sheet=$ws.Name; address=$rng.Address($false, $false); rowCount=$rowCount; columnCount=$colCount; values=$data }
  if ($r.LaunchedProcessId) { $res.launchedExcelPid = $r.LaunchedProcessId }
  $res | ConvertTo-Json -Depth 8
} catch {
  @{ ok=$false; error='read_failed'; detail="$($_.Exception.Message)" } | ConvertTo-Json
}
`.trim();

    try {
      const { stdout } = await execFileAsync(
        "powershell.exe",
        ["-NoLogo", "-NoProfile", "-NonInteractive", "-STA", "-ExecutionPolicy", "Bypass", "-Command", psScript],
        { windowsHide: true, encoding: "buffer", timeout: 20000, maxBuffer: 2 * 1024 * 1024 }
      );
      const outText = Buffer.isBuffer(stdout) ? stdout.toString("utf8") : String(stdout);
      return classifyResult(outText);
    } catch (e: any) {
      return extractFailureResult(e);
    }
  }
);
server.tool(
  "excel_list_macros",
  "List runnable procedures (Subs; module-level Public or implicitly-public -- Private/Friend are excluded, and Functions are not listed) in one VBA module, or in every module of the workbook at once if moduleName is omitted -- prefer omitting moduleName over calling this once per module when you need the whole workbook's macros, since each call re-resolves Excel/the workbook. Each result includes a fully-qualified name usable directly as the 'qualified' argument to excel_run_macro. Scans all currently open workbooks for a matching module unless workbookPath narrows it to one specific file (auto-launching/opening it if needed).",
  {
    moduleName: z.string().optional().describe("VBA module name to enumerate procedures in. Omit to list macros from every module in the target workbook in a single call."),
    basPath: z.string().optional().describe("Optional: full path to a previously-exported .bas file for this module; if given, its content hash is used to disambiguate which open workbook to target when multiple books have a same-named module."),
    workbookPath: z.string().optional().describe("Full path to the workbook file. If set, Excel is auto-launched and the file auto-opened when needed, instead of requiring it to already be open."),
  },
  async (params) => {
    const ps = process.env.MCP_PS_LIST;
    if (!ps) {
      return { content: [{ type: "text", text: JSON.stringify({ error: "MCP_PS_LIST not set" }) }] };
    }
    if (!fs.existsSync(ps)) {
      return { content: [{ type: "text", text: JSON.stringify({ error: `ps1 not found: ${ps}` }) }] };
    }

    let args: string[] = [
      "-NoLogo",
      "-NoProfile",
      "-NonInteractive",
      "-STA",
      "-ExecutionPolicy", "Bypass",
      "-File", ps,
      "-ListOutput","JSON"
    ];
    if (params.moduleName) {
        args.push("-ModuleName", params.moduleName);
    }
    if (params.basPath) {
        args.push("-BasPath", params.basPath);
    }
    if (params.workbookPath) {
        args.push("-WorkbookPath", params.workbookPath);
    }

    try {
      const { stdout } = await execFileAsync("powershell.exe", args, {
        windowsHide: true,
        encoding: "buffer",      // Buffer で受け取ってから UTF-8 に変換
        cwd: path.dirname(ps),   // ps1 のあるフォルダをカレントに
        timeout: 20000,          // ★ 20 秒で強制終了
        maxBuffer: 2 * 1024 * 1024
      });
      const outText  = Buffer.isBuffer(stdout) ? stdout.toString("utf8") : String(stdout);
      return classifyResult(outText);
    } catch (e: any) {
      return { content: [{ type: "text", text: JSON.stringify({ error: "ps failed", detail: String(e?.message ?? e) }) }] };
    }
  }
);

// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ excel_run_macros ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
server.tool(
  "excel_run_macro",
  "Run a VBA macro via Application.Run. WARNING: if the macro shows a dialog (MsgBox, InputBox) or a UserForm, or otherwise waits for user interaction, this call will hang until timeoutMs is reached; the timeout only stops this tool's own wait -- it does NOT close the dialog or unstick Excel, so check Excel directly afterward if you hit ERR_TIMEOUT. IMPORTANT: a successful response only means Application.Run completed without throwing an exception -- it does NOT confirm the macro did the intended thing (cell writes, files written, etc. are not verified by this tool). Prefer excel_list_macros first to get an exact 'qualified' name rather than guessing moduleName/procName.",
  {
    qualified: z.string().optional().describe("Fully-qualified macro name, e.g. \"'Book1.xlsm'!Module1.DoWork\" (as returned by excel_list_macros). Takes priority over moduleName/procName if both are given."),
    moduleName: z.string().optional().describe("Module name. Required together with procName if 'qualified' is not given."),
    procName: z.string().optional().describe("Procedure (Sub) name within moduleName. Required together with moduleName if 'qualified' is not given."),
    workbookName: z.string().optional().describe("Optional: display name of the workbook to disambiguate when the same module/proc name exists in multiple open workbooks."),
    basPath: z.string().optional().describe("Optional: full path to a previously-exported .bas file; its content hash disambiguates which open workbook to target."),
    workbookPath: z.string().optional().describe("Full path to the workbook file. If set, Excel is auto-launched and the file auto-opened when needed, instead of requiring it to already be open."),
    ActivateExcel: z.boolean().optional().describe("Bring the Excel window to the foreground before running."),
    ShowStatus: z.boolean().optional().describe("Show a transient message in Excel's status bar while/after running."),
    timeoutMs: z.number().optional().describe("Milliseconds to wait before giving up and returning ERR_TIMEOUT. Default 30000. Does not stop Excel itself if it's blocked on a dialog."),
  },
  async (params) => {
    const ps = process.env.MCP_PS_RUN || process.env.MCP_PS_LIST;
    if (!ps) {
      return { content: [{ type: "text", text: JSON.stringify({ error: "MCP_PS_RUN/MCP_PS_LIST not set" }) }] };
    }
    if (!fs.existsSync(ps)) {
      return { content: [{ type: "text", text: JSON.stringify({ error: `ps1 not found: ${ps}` }) }] };
    }

    // ← ここがポイント：一度だけ宣言してから push する
    let args: string[] = [
      "-NoLogo",
      "-NoProfile",
      "-NonInteractive",
      "-STA",
      "-ExecutionPolicy", "Bypass",
      "-File", ps
    ];

    if (params.qualified && params.qualified.trim().length > 0) {
      // 完全修飾が来たら最優先（.ps1 側に -Qualified 対応を実装済みであること）
      args.push("-Qualified", params.qualified);

    } else {
      if (!params.moduleName || !params.procName) {
        return { content: [{ type: "text", text: JSON.stringify({ error: "moduleName/procName or qualified required" }) }] };
      }
      args.push("-ModuleName", params.moduleName, "-ProcName", params.procName);
      if (params.workbookName) {
        args.push("-WorkbookName", params.workbookName);
      }
      if (params.basPath) {
        args.push("-BasPath", params.basPath);
      }
    }

    if (params.workbookPath) {
      args.push("-WorkbookPath", params.workbookPath);
    }
    if (params.ActivateExcel) {
      args.push("-ActivateExcel");
    }
    if (params.ShowStatus) {
      args.push("-ShowStatus");
    }

    const timeoutMs = params.timeoutMs ?? 30000;
    try {
      const { stdout } = await execFileAsync("powershell.exe", args, {
        windowsHide: true ,
        encoding: "buffer",
        maxBuffer: 2 * 1024 * 1024,
        cwd: path.dirname(ps),
        timeout: timeoutMs,
    });
      const outText  = Buffer.isBuffer(stdout) ? stdout.toString("utf8") : String(stdout);
      return classifyResult(outText);
    } catch (e: any) {
      // execFileAsyncのtimeoutで強制終了された場合。ただしこちら側のPowerShellプロセスを
      // 止めるだけで、Excel自体やダイアログで止まっている状態は解消されない点に注意。
      if (e?.killed) {
        return {
          content: [{ type: "text", text: JSON.stringify({
            ok: false,
            error: "ERR_TIMEOUT",
            detail: `Macro execution timed out after ${timeoutMs}ms. Excel may be blocked on a dialog (MsgBox/InputBox) or still running -- please check Excel directly.`,
          }) }],
          isError: true,
        };
      }
      return extractFailureResult(e);
    }
  }
);

// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ vba_search_code ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
server.tool(
  "vba_search_code",
  "Search VBA source code for a literal string or regex across all currently open workbooks (or one specific workbook via workbookPath / workbookFilter). Returns matching LINES with context (workbook/module/proc/line number), not full module code -- use excel_get_module_code to read a whole module. Results are capped at maxResults (default 50); if there were more matches, the response sets truncated:true and totalMatchCount so you know to narrow the query rather than assuming there were no more hits.",
  {
    query: z.string().describe("Search text. Plain substring by default, or a .NET regex pattern if useRegex is true. Case-insensitive."),
    moduleFilter: z.string().optional().describe("Restrict the search to a single module name."),
    workbookFilter: z.string().optional().describe("Restrict the search to a single open workbook's display name."),
    useRegex: z.boolean().optional().describe("Treat 'query' as a .NET regular expression instead of a literal substring."),
    workbookPath: z.string().optional().describe("Full path to a workbook to include in the search. If set and not already open, Excel is auto-launched and the file auto-opened before searching."),
    maxResults: z.number().optional().describe("Maximum number of hits to return in one call. Default 50. Excess hits are dropped, with truncated:true and totalMatchCount reported instead."),
  },
  async (params) => {
    // PowerShellワンライナーで開いている全ブックの全モジュールを走査
    // ・TrustOM 必須（VBAプロジェクトOMへのアクセスを信頼）
    // ・全コンポーネント種別を対象 vbext_ct_StdModule(1), Class(2), Document(100)
    const wbPath = psq(params.workbookPath ?? "");
    const maxResults = params.maxResults ?? 50;
    const dotSource = dotSourceExcelUtil();
    const psScript = `
# --- Force UTF-8 (no BOM) for stdout/stderr ---
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$OutputEncoding           = [Console]::OutputEncoding

${dotSource}

$ErrorActionPreference='Stop'
try{
  $r = Get-OrStartExcelApplication
  $excel = $r.App
}catch{
  Write-Output (@{ ok=$false; error='excel_not_found' } | ConvertTo-Json); exit
}

# workbookPath指定時は、未オープンなら自動で開く（対象を明示検索対象に含めるため）
$workbookPathParam = '${wbPath}'
if ($workbookPathParam -and $workbookPathParam.Trim().Length -gt 0) {
  try {
    $targetWb = Resolve-TargetWorkbook -App $excel -WorkbookPath $workbookPathParam
    Test-VbaTrustAccess -Workbook $targetWb | Out-Null
  } catch {
    Write-Output (@{ ok=$false; error="$($_.Exception.Message)" } | ConvertTo-Json); exit
  }
}

$hits=@()
$reRaw='${psq(params.query)}'
$useRe=${params.useRegex ? '$true' : '$false'}
$moduleFilter=${params.moduleFilter ? `'${params.moduleFilter.replace(/'/g,"''")}'` : '$null'}
$workbookFilter=${params.workbookFilter ? `'${params.workbookFilter.replace(/'/g,"''")}'` : '$null'}

# 大文字小文字無視のため (?i) を前置
if($useRe){ $re='(?i)'+$reRaw } else { $re=[regex]::Escape($reRaw); $re='(?i)'+$re }
$rx = [regex]::new($re)  # ★ 事前コンパイル

foreach($wb in @($excel.Workbooks)){
  if($workbookFilter -and $wb.Name -ne $workbookFilter){ continue }
  try{ $vbp=$wb.VBProject }catch{ continue }

  foreach($c in @($vbp.VBComponents)){
    # 種別フィルタ不要：全部対象
    $modName=$c.Name
    if($moduleFilter -and $modName -ne $moduleFilter){ continue }
    try{
      $cm=$c.CodeModule

      #$procKind = $null
      #$procName = $null

      # 走査ループ内のヒット生成部を置換
      $vbType = $c.Type   # 1:StdModule, 2:Class, 3:MSForm, 100:Document(Worksheet/ThisWorkbook)
      $ext = switch ($vbType) {
        1 { 'bas' }      # 標準モジュール
        3 { 'frm' }      # ユーザーフォーム（.frm + .frx）
        default { 'cls' }# クラス/シート/ThisWorkbook は .cls
      }
      $text=$cm.Lines(1,$cm.CountOfLines)
      $i=0

      #try { $procName = $cm.ProcOfLine([int]$i, [ref]$procKind) } catch {}
      #if (-not $procName) {
      #  $declRe = [regex]'(?im)^\s*Public\s+(Sub|Function)\s+([A-Za-z_]\w*)\b'
      #  for ($j = [Math]::Min($i, $cm.CountOfLines); $j -ge 1; $j--) {
      #    try {
      #      $decl = $cm.Lines($j, 1)
      #      $m = $declRe.Match($decl)
      #      if ($m.Success) { $procName = $m.Groups[2].Value; break }
      #    } catch {}
      #  }
      #}
      #$text=$cm.Lines(1,$cm.CountOfLines)
      #$i=0
      foreach($line in $text -split "\\r?\\n"){
        $i++
        #if([regex]::IsMatch($line,$re)){

        if($rx.IsMatch($line)){
          $procKind = $null
          $procName = $null
          try { $procName = $cm.ProcOfLine([int]$i, [ref]$procKind) } catch {}
          if (-not $procName) {
            $declRe = [regex]'(?im)^\\s*Public\\s+(Sub|Function)\\s+([A-Za-z_]\\w*)\\b'
            for ($j=[Math]::Min($i,$cm.CountOfLines); $j -ge 1; $j--) {
              try {
                $m = $declRe.Match($cm.Lines($j,1))
                if ($m.Success) { $procName = $m.Groups[2].Value; break }
              } catch {}
            }
          }

          $hits += [pscustomobject]@{
            workbook  = $wb.Name
            module    = $modName
            proc      = $procName
            line      = $i
            snippet   = $line.Trim()
            qualified = if ($procName) { "'$($wb.Name)'!$modName.$procName" } else { "'$($wb.Name)'!$modName" }  # ★ 修正
            compType  = $vbType
            exportExt = $ext                 
            }
        }
      }
    }catch{}
  }
}
$totalMatchCount = $hits.Count
$truncated = $false
$maxResultsParam = ${maxResults}
if ($totalMatchCount -gt $maxResultsParam) {
  $hits = $hits[0..($maxResultsParam - 1)]
  $truncated = $true
}
$searchRes = @{ ok=$true; query=$reRaw; hits=$hits; count=$hits.Count; totalMatchCount=$totalMatchCount; truncated=$truncated }
if ($r.LaunchedProcessId) { $searchRes.launchedExcelPid = $r.LaunchedProcessId }
$searchRes | ConvertTo-Json -Depth 6
`;

    try {
      const { stdout } = await execFileAsync(
        "powershell.exe",
        ["-NoLogo","-NoProfile","-NonInteractive","-STA","-ExecutionPolicy","Bypass","-Command", psScript],
        { windowsHide: true, encoding: "buffer", timeout: 20000, maxBuffer: 2*1024*1024 }
      );
      const outText  = Buffer.isBuffer(stdout) ? stdout.toString("utf8") : String(stdout);
      return classifyResult(outText);
    } catch (e:any) {
      return { content: [{ type: "text", text: JSON.stringify({ ok:false, error:"ps_failed", detail:String(e?.message ?? e) }) }], isError: true };
    }
  }
);

// ------------------------------------------------------------ vba_analyze_flow ------------------------------------------------------------
// Reuses scripts/VBA-FlowJson.ps1 (also used by the extension's manual "Generate VBA
// Flow Chart" command) as an external process against a live COM snapshot of ALL the
// workbook's modules, written to sibling temp files -- the script's symbol table
// (BuildSymbolTable) scans every .bas/.cls/.frm it finds in the folder, so every module
// must be present as a file there for cross-module call resolution to work. No disk
// writes to the workbook's own folder (result is always returned inline, never saved).
function computeNormalizedTextHash(text: string): string {
  const norm = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  return createHash("sha256").update(norm, "utf8").digest("hex");
}

// Resolves the workbook once via COM, then snapshots every VBA module's current code
// into <tempDir>/<ModuleName>.<ext> in a single PowerShell/COM session (avoiding one
// launch per module) so VBA-FlowJson.ps1's symbol table can see the whole project.
// Returns the target module's own code/type/name, same shape as a single-module read.
async function writeAllModulesToTemp(
  wbEscaped: string,
  wbPathEscaped: string,
  modEscaped: string,
  tempDirEscaped: string
): Promise<
  | { ok: true; workbook: string; module: string; componentType: number; currentCode: string; launchedExcelPid?: number }
  | { ok: false; content: { type: "text"; text: string }[]; isError: true }
> {
  const dotSource = dotSourceExcelUtil();
  const psScript = `
$ErrorActionPreference='Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$OutputEncoding           = [Console]::OutputEncoding

${dotSource}

try { $r = Get-OrStartExcelApplication; $excel = $r.App }
catch { @{ ok=$false; error='excel_not_found' } | ConvertTo-Json ; exit }

try { $wb = Resolve-TargetWorkbook -App $excel -WorkbookPath '${wbPathEscaped}' -WorkbookName '${wbEscaped}' }
catch { @{ ok=$false; error="$($_.Exception.Message)" } | ConvertTo-Json ; exit }

try { Test-VbaTrustAccess -Workbook $wb | Out-Null }
catch { @{ ok=$false; error="$($_.Exception.Message)" } | ConvertTo-Json ; exit }

$extByType = @{ 1='bas'; 2='cls'; 100='cls'; 3='frm' }
$targetCode = $null
$targetType = $null
$targetName = $null

try {
  foreach ($vbc in $wb.VBProject.VBComponents) {
    $cm = $vbc.CodeModule
    $code = if ($cm.CountOfLines -gt 0) { $cm.Lines(1, $cm.CountOfLines) } else { "" }
    if ($vbc.Name -eq '${modEscaped}') {
      $targetCode = $code
      $targetType = [int]$vbc.Type
      $targetName = $vbc.Name
    }
    $ext = $extByType[[int]$vbc.Type]
    if (-not $ext) { continue }
    $outPath = Join-Path '${tempDirEscaped}' "$($vbc.Name).$ext"
    [System.IO.File]::WriteAllText($outPath, $code, (New-Object System.Text.UTF8Encoding($false)))
  }
} catch {
  @{ ok=$false; error='read_failed'; detail="$($_.Exception.Message)" } | ConvertTo-Json
  exit
}

if ($null -eq $targetName) {
  @{ ok=$false; error='module_not_found'; module='${modEscaped}' } | ConvertTo-Json
  exit
}

$res = @{ ok=$true; workbook=$wb.Name; module=$targetName; componentType=$targetType; currentCode=$targetCode }
if ($r.LaunchedProcessId) { $res.launchedExcelPid = $r.LaunchedProcessId }
$res | ConvertTo-Json -Depth 6
`.trim();

  try {
    const { stdout } = await execFileAsync(
      "powershell.exe",
      ["-NoLogo", "-NoProfile", "-NonInteractive", "-STA", "-ExecutionPolicy", "Bypass", "-Command", psScript],
      { windowsHide: true, encoding: "buffer", timeout: 30000, maxBuffer: 4 * 1024 * 1024 }
    );
    const outText = Buffer.isBuffer(stdout) ? stdout.toString("utf8") : String(stdout);
    const classified = classifyResult(outText);
    if (classified.isError) {
      return { ok: false, content: classified.content, isError: true };
    }

    let payload: any = null;
    try {
      const start = Math.min(...['{', '['].map(ch => { const i = outText.indexOf(ch); return i === -1 ? Number.POSITIVE_INFINITY : i; }));
      payload = Number.isFinite(start) ? JSON.parse(outText.slice(start)) : null;
    } catch { /* noop */ }

    if (!payload?.ok) {
      return { ok: false, content: [{ type: "text", text: outText }], isError: true };
    }

    return {
      ok: true,
      workbook: payload.workbook,
      module: payload.module,
      componentType: payload.componentType,
      currentCode: payload.currentCode,
      launchedExcelPid: payload.launchedExcelPid,
    };
  } catch (e: any) {
    return { ok: false, content: [{ type: "text", text: JSON.stringify({ ok: false, error: "ps_failed", detail: String(e?.message ?? e) }) }], isError: true };
  }
}

server.tool(
  "vba_analyze_flow",
  "Analyze the control-flow structure (branches, loops, GoTo/labels, calls) of a VBA procedure, or list procedures in a module, as structured JSON -- for answering questions like 'where does this GoTo jump to' or 'is there an unreachable branch' directly from data, without re-reading and mentally parsing raw code. Reuses the same analyzer as the extension's manual 'Generate VBA Flow Chart' command (If/ElseIf/Else, Do/Loop, For/Next, Select Case, With, GoTo/labels, Exit/Return, Err.Raise), run as an external process against a live snapshot of the module's current code (not the exported .bas/.cls/.frm on disk). Omit 'procedure' to get a lightweight list of {name, kind, startLine, endLine} for every Sub/Function/Property in the module -- prefer this over full analysis when you just need to know what procedures exist. Cross-module calls ARE resolved (every module in the workbook is snapshotted alongside the target one): resolved:false on a call means the callee genuinely could not be found anywhere in this workbook, not merely that this tool didn't check other modules. Never writes anything to disk. If the workbook is already open in Excel, 'workbook' (its display name) is enough. Otherwise pass 'workbookPath' (full file path): Excel will be launched if not running, and the file opened if not already open. Fails with ERR_VBOM_TRUST_DISABLED if Excel's 'Trust access to the VBA project object model' setting is off.",
  {
    workbook: z.string().optional().describe("Workbook display name. Either this or workbookPath is required; workbookPath is preferred since it also auto-launches/opens Excel if needed."),
    workbookPath: z.string().optional().describe("Full path to the workbook file. If set, Excel is auto-launched and the file auto-opened when needed."),
    module: z.string().describe("VBA module name to analyze."),
    procedure: z.string().optional().describe("Procedure (Sub/Function/Property) name within module to get full flow detail for. Omit to instead get a lightweight list of every procedure in the module ({name, kind, startLine, endLine}) without flow detail -- cheaper, use this first if you don't already know the exact procedure name."),
  },
  async (params) => {
    if (!params.workbook && !params.workbookPath) {
      return { content: [{ type: "text", text: JSON.stringify({ ok: false, error: "workbook or workbookPath is required" }) }], isError: true };
    }
    const wb = psq(params.workbook ?? "");
    const wbPath = psq(params.workbookPath ?? "");
    const mod = psq(params.module);

    const tempDir = path.join(os.tmpdir(), `VBAFlow_${Date.now()}_${Math.random().toString(36).slice(2)}`);
    const resultJsonFile = path.join(tempDir, "result.flow.json");

    try {
      fs.mkdirSync(tempDir, { recursive: true });

      const readResult = await writeAllModulesToTemp(wb, wbPath, mod, psq(tempDir));
      if (!readResult.ok) { return { content: readResult.content, isError: readResult.isError }; }

      const extByType: Record<number, string> = { 1: "bas", 2: "cls", 100: "cls", 3: "frm" };
      const ext = extByType[readResult.componentType];
      if (!ext) {
        return {
          content: [{ type: "text", text: JSON.stringify({ ok: false, error: "unsupported_component_type", componentType: readResult.componentType }) }],
          isError: true,
        };
      }
      const sourceFile = path.join(tempDir, `${readResult.module}.${ext}`);
      const sourceHash = computeNormalizedTextHash(readResult.currentCode);

      const scriptsDir = getScriptsDir();
      if (!scriptsDir) {
        return { content: [{ type: "text", text: JSON.stringify({ ok: false, error: "MCP_SCRIPTS_DIR/MCP_PS_LIST not set" }) }], isError: true };
      }
      const flowScript = path.join(scriptsDir, "VBA-FlowJson.ps1");
      if (!fs.existsSync(flowScript)) {
        return { content: [{ type: "text", text: JSON.stringify({ ok: false, error: `script not found: ${flowScript}` }) }], isError: true };
      }

      const args = [
        "-NoLogo", "-NoProfile", "-NonInteractive", "-STA", "-ExecutionPolicy", "Bypass",
        "-File", flowScript,
        "-FolderPath", tempDir,
        "-FilePath", sourceFile,
        "-OutputPath", resultJsonFile,
        "-Encoding", "UTF8",
      ];
      try {
        await execFileAsync("powershell.exe", args, {
          windowsHide: true,
          encoding: "buffer",
          timeout: 30000,
          maxBuffer: 2 * 1024 * 1024,
        });
      } catch (e: any) {
        return {
          content: [{ type: "text", text: JSON.stringify({ ok: false, error: "flow_analysis_failed", detail: String(e?.message ?? e) }) }],
          isError: true,
        };
      }

      if (!fs.existsSync(resultJsonFile)) {
        return { content: [{ type: "text", text: JSON.stringify({ ok: false, error: "flow_analysis_failed", detail: "result JSON was not produced" }) }], isError: true };
      }

      const doc: any = JSON.parse(fs.readFileSync(resultJsonFile, "utf8"));
      const procedures: any[] = Array.isArray(doc.procedures) ? doc.procedures : [];

      if (params.procedure) {
        const found = procedures.find((p: any) => p.name === params.procedure);
        if (!found) {
          return {
            content: [{
              type: "text",
              text: JSON.stringify({
                ok: false,
                error: "procedure_not_found",
                module: readResult.module,
                procedure: params.procedure,
                availableProcedures: procedures.map((p: any) => p.name),
              }),
            }],
            isError: true,
          };
        }
        const res: Record<string, unknown> = {
          ok: true,
          workbook: readResult.workbook,
          module: readResult.module,
          componentType: readResult.componentType,
          mode: "procedure",
          procedure: found,
          sourceHash,
        };
        if (readResult.launchedExcelPid) { res.launchedExcelPid = readResult.launchedExcelPid; }
        return { content: [{ type: "text", text: JSON.stringify(res, null, 2) }] };
      }

      const list = procedures.map((p: any) => ({ name: p.name, kind: p.kind, startLine: p.startLine, endLine: p.endLine }));
      const res: Record<string, unknown> = {
        ok: true,
        workbook: readResult.workbook,
        module: readResult.module,
        componentType: readResult.componentType,
        mode: "list",
        procedures: list,
        sourceHash,
      };
      if (readResult.launchedExcelPid) { res.launchedExcelPid = readResult.launchedExcelPid; }
      return { content: [{ type: "text", text: JSON.stringify(res, null, 2) }] };
    } finally {
      try { fs.rmSync(tempDir, { recursive: true, force: true }); } catch { /* noop */ }
    }
  }
);

// ------------------------------------------------------------ vba_render_flowchart ------------------------------------------------------------
// Reuses scripts/VBA-FlowJson.ps1 + scripts/Convert-FlowJsonToMermaid.ps1 (also used by the
// extension's manual "Generate VBA Flow Chart" command) as external processes against a live
// COM snapshot of ALL the workbook's modules (see writeAllModulesToTemp above -- needed for
// cross-module call resolution in the call graph). Never writes to the workbook's own folder --
// all Mermaid output is generated in a throwaway temp folder and returned inline, then removed.
server.tool(
  "vba_render_flowchart",
  "Render a VBA procedure's control flow as Mermaid flowchart text (flowchart TD), or the whole module's call graph if 'procedure' is omitted -- for pasting into a Markdown preview, mermaid.live, or a Mermaid-rendering chat client to see a diagram instead of reasoning over vba_analyze_flow's raw JSON. Uses the exact same rendering logic as the extension's manual 'Generate VBA Flow Chart' command (If/ElseIf/Else, Do/Loop, For/Next, Select Case, GoTo/labels rendered as Mermaid node/edge shapes), run as an external process against a live snapshot of the module's current code (not the exported .bas/.cls/.frm on disk). Cross-module calls in the call graph ARE resolved (every module in the workbook is snapshotted alongside the target one). Never writes anything to disk (no vbaExport/.mmd files are created; the workbook's own folder is untouched). If the workbook is already open in Excel, 'workbook' (its display name) is enough. Otherwise pass 'workbookPath' (full file path): Excel will be launched if not running, and the file opened if not already open. Fails with ERR_VBOM_TRUST_DISABLED if Excel's 'Trust access to the VBA project object model' setting is off.",
  {
    workbook: z.string().optional().describe("Workbook display name. Either this or workbookPath is required; workbookPath is preferred since it also auto-launches/opens Excel if needed."),
    workbookPath: z.string().optional().describe("Full path to the workbook file. If set, Excel is auto-launched and the file auto-opened when needed."),
    module: z.string().describe("VBA module name to render."),
    procedure: z.string().optional().describe("Procedure (Sub/Function/Property) name within module to render a detailed flowchart for. Omit to instead render the whole module's call graph (which procedure calls which) as a single diagram."),
  },
  async (params) => {
    if (!params.workbook && !params.workbookPath) {
      return { content: [{ type: "text", text: JSON.stringify({ ok: false, error: "workbook or workbookPath is required" }) }], isError: true };
    }
    const wb = psq(params.workbook ?? "");
    const wbPath = psq(params.workbookPath ?? "");
    const mod = psq(params.module);

    const tempDir = path.join(os.tmpdir(), `VBAFlow_${Date.now()}_${Math.random().toString(36).slice(2)}`);
    const flowJsonFile = path.join(tempDir, "result.flow.json");

    try {
      fs.mkdirSync(tempDir, { recursive: true });

      const readResult = await writeAllModulesToTemp(wb, wbPath, mod, psq(tempDir));
      if (!readResult.ok) { return { content: readResult.content, isError: readResult.isError }; }

      const extByType: Record<number, string> = { 1: "bas", 2: "cls", 100: "cls", 3: "frm" };
      const ext = extByType[readResult.componentType];
      if (!ext) {
        return {
          content: [{ type: "text", text: JSON.stringify({ ok: false, error: "unsupported_component_type", componentType: readResult.componentType }) }],
          isError: true,
        };
      }
      const sourceFile = path.join(tempDir, `${readResult.module}.${ext}`);
      const sourceHash = computeNormalizedTextHash(readResult.currentCode);

      const scriptsDir = getScriptsDir();
      if (!scriptsDir) {
        return { content: [{ type: "text", text: JSON.stringify({ ok: false, error: "MCP_SCRIPTS_DIR/MCP_PS_LIST not set" }) }], isError: true };
      }
      const flowScript = path.join(scriptsDir, "VBA-FlowJson.ps1");
      const mermaidScript = path.join(scriptsDir, "Convert-FlowJsonToMermaid.ps1");
      if (!fs.existsSync(flowScript)) {
        return { content: [{ type: "text", text: JSON.stringify({ ok: false, error: `script not found: ${flowScript}` }) }], isError: true };
      }
      if (!fs.existsSync(mermaidScript)) {
        return { content: [{ type: "text", text: JSON.stringify({ ok: false, error: `script not found: ${mermaidScript}` }) }], isError: true };
      }

      try {
        await execFileAsync("powershell.exe", [
          "-NoLogo", "-NoProfile", "-NonInteractive", "-STA", "-ExecutionPolicy", "Bypass",
          "-File", flowScript,
          "-FolderPath", tempDir,
          "-FilePath", sourceFile,
          "-OutputPath", flowJsonFile,
          "-Encoding", "UTF8",
        ], { windowsHide: true, encoding: "buffer", timeout: 30000, maxBuffer: 2 * 1024 * 1024 });
      } catch (e: any) {
        return { content: [{ type: "text", text: JSON.stringify({ ok: false, error: "flow_analysis_failed", detail: String(e?.message ?? e) }) }], isError: true };
      }
      if (!fs.existsSync(flowJsonFile)) {
        return { content: [{ type: "text", text: JSON.stringify({ ok: false, error: "flow_analysis_failed", detail: "result JSON was not produced" }) }], isError: true };
      }

      try {
        await execFileAsync("powershell.exe", [
          "-NoLogo", "-NoProfile", "-NonInteractive", "-STA", "-ExecutionPolicy", "Bypass",
          "-File", mermaidScript,
          "-JsonPath", flowJsonFile,
          "-OutDir", tempDir,
        ], { windowsHide: true, encoding: "buffer", timeout: 30000, maxBuffer: 2 * 1024 * 1024 });
      } catch (e: any) {
        return { content: [{ type: "text", text: JSON.stringify({ ok: false, error: "mermaid_render_failed", detail: String(e?.message ?? e) }) }], isError: true };
      }

      const doc: any = JSON.parse(fs.readFileSync(flowJsonFile, "utf8"));
      const procedures: any[] = Array.isArray(doc.procedures) ? doc.procedures : [];

      if (params.procedure) {
        const found = procedures.find((p: any) => p.name === params.procedure);
        if (!found) {
          return {
            content: [{
              type: "text",
              text: JSON.stringify({
                ok: false,
                error: "procedure_not_found",
                module: readResult.module,
                procedure: params.procedure,
                availableProcedures: procedures.map((p: any) => p.name),
              }),
            }],
            isError: true,
          };
        }
        const mmdFile = path.join(tempDir, `${readResult.module}.${params.procedure}.mmd`);
        if (!fs.existsSync(mmdFile)) {
          return {
            content: [{ type: "text", text: JSON.stringify({ ok: false, error: "mermaid_render_failed", detail: "expected .mmd file was not produced", expectedPath: mmdFile }) }],
            isError: true,
          };
        }
        const mermaid = fs.readFileSync(mmdFile, "utf8");
        const res: Record<string, unknown> = {
          ok: true,
          workbook: readResult.workbook,
          module: readResult.module,
          componentType: readResult.componentType,
          mode: "procedure",
          procedure: params.procedure,
          mermaid,
          sourceHash,
        };
        if (readResult.launchedExcelPid) { res.launchedExcelPid = readResult.launchedExcelPid; }
        return { content: [{ type: "text", text: JSON.stringify(res, null, 2) }] };
      }

      const callgraphFile = path.join(tempDir, `${readResult.module}.callgraph.mmd`);
      if (!fs.existsSync(callgraphFile)) {
        return {
          content: [{ type: "text", text: JSON.stringify({ ok: false, error: "mermaid_render_failed", detail: "expected callgraph .mmd file was not produced", expectedPath: callgraphFile }) }],
          isError: true,
        };
      }
      const mermaid = fs.readFileSync(callgraphFile, "utf8");
      const res: Record<string, unknown> = {
        ok: true,
        workbook: readResult.workbook,
        module: readResult.module,
        componentType: readResult.componentType,
        mode: "callgraph",
        mermaid,
        sourceHash,
      };
      if (readResult.launchedExcelPid) { res.launchedExcelPid = readResult.launchedExcelPid; }
      return { content: [{ type: "text", text: JSON.stringify(res, null, 2) }] };
    } finally {
      try { fs.rmSync(tempDir, { recursive: true, force: true }); } catch { /* noop */ }
    }
  }
);

// ------------------------------------------------------------ vba_list_dependencies ------------------------------------------------------------
// Resolves the workbook once via COM, then reads every VBA module's current code into
// memory in a single PowerShell session (no temp files -- unlike vba_analyze_flow/
// vba_render_flowchart, this tool never invokes an external script; the matching is
// pure regex over the text, done in dependencyScan.ts). Read-only, advisory only.
async function readAllModulesCode(
  wbEscaped: string,
  wbPathEscaped: string
): Promise<
  | { ok: true; workbook: string; modules: { name: string; componentType: number; code: string }[]; launchedExcelPid?: number }
  | { ok: false; content: { type: "text"; text: string }[]; isError: true }
> {
  const dotSource = dotSourceExcelUtil();
  const psScript = `
$ErrorActionPreference='Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$OutputEncoding           = [Console]::OutputEncoding

${dotSource}

try { $r = Get-OrStartExcelApplication; $excel = $r.App }
catch { @{ ok=$false; error='excel_not_found' } | ConvertTo-Json ; exit }

try { $wb = Resolve-TargetWorkbook -App $excel -WorkbookPath '${wbPathEscaped}' -WorkbookName '${wbEscaped}' }
catch { @{ ok=$false; error="$($_.Exception.Message)" } | ConvertTo-Json ; exit }

try { Test-VbaTrustAccess -Workbook $wb | Out-Null }
catch { @{ ok=$false; error="$($_.Exception.Message)" } | ConvertTo-Json ; exit }

$mods = New-Object System.Collections.Generic.List[object]
try {
  foreach ($vbc in $wb.VBProject.VBComponents) {
    $cm = $vbc.CodeModule
    $code = if ($cm.CountOfLines -gt 0) { $cm.Lines(1, $cm.CountOfLines) } else { "" }
    $mods.Add(@{ name = $vbc.Name; componentType = [int]$vbc.Type; code = $code })
  }
} catch {
  @{ ok=$false; error='read_failed'; detail="$($_.Exception.Message)" } | ConvertTo-Json
  exit
}

$res = @{ ok=$true; workbook=$wb.Name; modules=$mods.ToArray() }
if ($r.LaunchedProcessId) { $res.launchedExcelPid = $r.LaunchedProcessId }
$res | ConvertTo-Json -Depth 6
`.trim();

  try {
    const { stdout } = await execFileAsync(
      "powershell.exe",
      ["-NoLogo", "-NoProfile", "-NonInteractive", "-STA", "-ExecutionPolicy", "Bypass", "-Command", psScript],
      { windowsHide: true, encoding: "buffer", timeout: 30000, maxBuffer: 8 * 1024 * 1024 }
    );
    const outText = Buffer.isBuffer(stdout) ? stdout.toString("utf8") : String(stdout);
    const classified = classifyResult(outText);
    if (classified.isError) {
      return { ok: false, content: classified.content, isError: true };
    }

    let payload: any = null;
    try {
      const start = Math.min(...['{', '['].map(ch => { const i = outText.indexOf(ch); return i === -1 ? Number.POSITIVE_INFINITY : i; }));
      payload = Number.isFinite(start) ? JSON.parse(outText.slice(start)) : null;
    } catch { /* noop */ }

    if (!payload?.ok) {
      return { ok: false, content: [{ type: "text", text: outText }], isError: true };
    }

    return {
      ok: true,
      workbook: payload.workbook,
      modules: payload.modules ?? [],
      launchedExcelPid: payload.launchedExcelPid,
    };
  } catch (e: any) {
    return { ok: false, content: [{ type: "text", text: JSON.stringify({ ok: false, error: "ps_failed", detail: String(e?.message ?? e) }) }], isError: true };
  }
}

server.tool(
  "vba_list_dependencies",
  "Scan VBA module source for external/platform dependencies via lightweight regex matching -- Windows API 'Declare' statements, CreateObject/GetObject COM automation calls, VBA's built-in Shell function calls, Application.Run (dynamic macro dispatch, including cross-workbook targets), native VBA file I/O (Open/Kill/FileCopy/MkDir/RmDir), Scripting.FileSystemObject file/folder methods (OpenTextFile/CreateTextFile/CopyFile/DeleteFile/MoveFile/CreateFolder/DeleteFolder/MoveFolder -- matched by method name alone, so a fileIo entry with methodNameOnly:true means the call target's actual type was not verified and some other object with a same-named method would also match), and Workbooks.Open (external workbook references). Read-only and advisory: this is best-effort text matching (not a real VBA parser), at the same rigor level as excel_update_module_code's lintWarnings -- it can miss dynamic or commented-out cases, and can rarely false-positive on text that merely resembles the pattern inside an unrelated string literal. A procedure with no incoming calls found by vba_analyze_flow is not necessarily unused -- check here for an Application.Run dispatching to it by name before concluding that. Useful for scoping migration work (e.g. Office Scripts has no equivalent for any of these), auditing what a workbook automates outside VBA itself, or finding which other files must travel together with this workbook (Workbooks.Open targets). Omit 'module' to scan every module in the workbook in a single COM session; pass it to scan only that module. Modules with zero findings are omitted from the response. If the workbook is already open in Excel, 'workbook' (its display name) is enough. Otherwise pass 'workbookPath' (full file path): Excel will be launched if not running, and the file opened if not already open. Fails with ERR_VBOM_TRUST_DISABLED if Excel's 'Trust access to the VBA project object model' setting is off.",
  {
    workbook: z.string().optional().describe("Workbook display name. Either this or workbookPath is required; workbookPath is preferred since it also auto-launches/opens Excel if needed."),
    workbookPath: z.string().optional().describe("Full path to the workbook file. If set, Excel is auto-launched and the file auto-opened when needed."),
    module: z.string().optional().describe("VBA module name to scan. Omit to scan every module in the workbook in one call."),
  },
  async (params) => {
    if (!params.workbook && !params.workbookPath) {
      return { content: [{ type: "text", text: JSON.stringify({ ok: false, error: "workbook or workbookPath is required" }) }], isError: true };
    }
    const wb = psq(params.workbook ?? "");
    const wbPath = psq(params.workbookPath ?? "");

    const readResult = await readAllModulesCode(wb, wbPath);
    if (!readResult.ok) { return { content: readResult.content, isError: readResult.isError }; }

    let targetModules = readResult.modules;
    if (params.module) {
      targetModules = targetModules.filter((m) => m.name === params.module);
      if (targetModules.length === 0) {
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              ok: false,
              error: "module_not_found",
              module: params.module,
              availableModules: readResult.modules.map((m) => m.name),
            }),
          }],
          isError: true,
        };
      }
    }

    const scans: ModuleDependencyScan[] = targetModules.map((m) => scanModuleForDependencies(m.name, m.code));
    const nonEmptyScans = scans.filter((s) =>
      s.apiDeclares.length > 0 ||
      s.comObjects.length > 0 ||
      s.shellCalls.length > 0 ||
      s.applicationRunCalls.length > 0 ||
      s.fileIo.length > 0 ||
      s.externalWorkbooks.length > 0
    );

    const summary = {
      modulesScanned: targetModules.length,
      modulesWithFindings: nonEmptyScans.length,
      apiDeclares: scans.reduce((sum, s) => sum + s.apiDeclares.length, 0),
      comObjects: scans.reduce((sum, s) => sum + s.comObjects.length, 0),
      shellCalls: scans.reduce((sum, s) => sum + s.shellCalls.length, 0),
      applicationRunCalls: scans.reduce((sum, s) => sum + s.applicationRunCalls.length, 0),
      fileIo: scans.reduce((sum, s) => sum + s.fileIo.length, 0),
      externalWorkbooks: scans.reduce((sum, s) => sum + s.externalWorkbooks.length, 0),
    };

    const res: Record<string, unknown> = {
      ok: true,
      workbook: readResult.workbook,
      summary,
      modules: nonEmptyScans,
    };
    if (readResult.launchedExcelPid) { res.launchedExcelPid = readResult.launchedExcelPid; }
    return { content: [{ type: "text", text: JSON.stringify(res, null, 2) }] };
  }
);

// ------------------------------------------------------------ vba_list_references ------------------------------------------------------------
// Reuses readAllModulesCode (defined above, for vba_list_dependencies) -- same
// single-COM-session snapshot of every module, no new COM/PowerShell logic needed
// here. The regex matching itself lives in referenceScan.ts, pure and COM-free.
server.tool(
  "vba_list_references",
  "List event-procedure entry points and internal Excel object references via lightweight regex matching. Event procedures: Workbook_*/Worksheet_*/UserForm_* (matched by VBA's own naming convention, so no event-name list is needed) plus legacy Auto_Open/Auto_Close -- these are candidate triggers for code the user never calls directly, so a procedure with no incoming calls in vba_analyze_flow is not necessarily unused if it's one of these (or is dispatched by vba_list_dependencies' Application.Run). Embedded ActiveX control events (e.g. CommandButton1_Click) are NOT detected -- telling those apart from an ordinary Sub needs the sheet/form's control names, which this tool does not read. References: Worksheets(...)/Sheets(...) by name (sheetName is the literal when static, null with dynamic:true otherwise), and likely named-range references via Range(...)/Names(...) -- Range(\"...\") is only reported when the literal does NOT look like a plain cell address (e.g. \"A1\", \"B2:C10\"), since those make up the overwhelming majority of ordinary Range(...) calls and would otherwise drown out genuine named-range references; dynamic Range(variable) calls are not reported at all for the same reason (most are computed cell addresses, not named-range access), while Names(...) is reported even when dynamic since any use of the Names collection is inherently about a named range. Useful for impact analysis: what triggers exist in this workbook, and what would break if a given sheet were renamed/deleted or a named range were removed. Read-only and advisory -- best-effort text matching, not a real VBA parser, at the same rigor level as vba_list_dependencies and excel_update_module_code's lintWarnings. Omit 'module' to scan every module in the workbook in a single COM session; modules with no findings in any category are omitted from the response. If the workbook is already open in Excel, 'workbook' (its display name) is enough. Otherwise pass 'workbookPath' (full file path): Excel will be launched if not running, and the file opened if not already open. Fails with ERR_VBOM_TRUST_DISABLED if Excel's 'Trust access to the VBA project object model' setting is off.",
  {
    workbook: z.string().optional().describe("Workbook display name. Either this or workbookPath is required; workbookPath is preferred since it also auto-launches/opens Excel if needed."),
    workbookPath: z.string().optional().describe("Full path to the workbook file. If set, Excel is auto-launched and the file auto-opened when needed."),
    module: z.string().optional().describe("VBA module name to scan. Omit to scan every module in the workbook in one call."),
  },
  async (params) => {
    if (!params.workbook && !params.workbookPath) {
      return { content: [{ type: "text", text: JSON.stringify({ ok: false, error: "workbook or workbookPath is required" }) }], isError: true };
    }
    const wb = psq(params.workbook ?? "");
    const wbPath = psq(params.workbookPath ?? "");

    const readResult = await readAllModulesCode(wb, wbPath);
    if (!readResult.ok) { return { content: readResult.content, isError: readResult.isError }; }

    let targetModules = readResult.modules;
    if (params.module) {
      targetModules = targetModules.filter((m) => m.name === params.module);
      if (targetModules.length === 0) {
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              ok: false,
              error: "module_not_found",
              module: params.module,
              availableModules: readResult.modules.map((m) => m.name),
            }),
          }],
          isError: true,
        };
      }
    }

    const scans: ModuleReferenceScan[] = targetModules.map((m) => scanModuleForReferences(m.name, m.code));
    const nonEmptyScans = scans.filter((s) =>
      s.eventProcedures.length > 0 ||
      s.sheetReferences.length > 0 ||
      s.namedRangeReferences.length > 0
    );

    const summary = {
      modulesScanned: targetModules.length,
      modulesWithFindings: nonEmptyScans.length,
      eventProcedures: scans.reduce((sum, s) => sum + s.eventProcedures.length, 0),
      sheetReferences: scans.reduce((sum, s) => sum + s.sheetReferences.length, 0),
      namedRangeReferences: scans.reduce((sum, s) => sum + s.namedRangeReferences.length, 0),
    };

    const res: Record<string, unknown> = {
      ok: true,
      workbook: readResult.workbook,
      summary,
      modules: nonEmptyScans,
    };
    if (readResult.launchedExcelPid) { res.launchedExcelPid = readResult.launchedExcelPid; }
    return { content: [{ type: "text", text: JSON.stringify(res, null, 2) }] };
  }
);

// ------------------------------------------------------------ vba_list_variable_scopes ------------------------------------------------------------
// Reuses readAllModulesCode (defined above, for vba_list_dependencies) -- same
// single-COM-session snapshot of every module, no new COM/PowerShell logic needed
// here. The regex matching itself lives in variableScopeScan.ts, pure and COM-free.
server.tool(
  "vba_list_variable_scopes",
  "List variable and constant declarations (Dim/Private/Public/Static/Const) classified by scope: 'procedure' (local to one Sub/Function/Property -- declaredIn names it), 'module' (Private, or unmarked Dim/Const at module level -- visible module-wide but not from other modules), or 'public' (Public at module level -- visible from anywhere in the project). This exists because VBE's own Find & Replace ('Search In: Current Project') is a blind text substitution with no concept of scope: it will happily replace an unrelated local variable in a completely different procedure just because it shares the same name. Omit 'variableName' to list every declaration in 'module' (or the whole workbook if 'module' is also omitted) -- this establishes a declaration's true boundary. Provide 'variableName' (module becomes required) to instead find every usage of that one declaration within its correct boundary: module-/public-scoped lookups automatically skip any procedure that shadows the name with its own local declaration, so an unrelated same-named local elsewhere is never mixed in. If the name matches more than one declaration in 'module' (several procedures each with their own same-named local, or a module-level declaration itself shadowed by a same-named local somewhere), the response is ambiguous_declaration listing every candidate (scope, declaredIn, line) -- pass 'procedure' (matching a candidate's declaredIn) to pick one. Each usage is classified 'write' (matched via a Set-optional assignment pattern near the line start) or 'reference' (everything else -- not distinguishing a read from e.g. an 'If name = x Then' comparison beyond that). The declaration's own line is never included as a usage. Read-only and advisory throughout: best-effort text matching, not a real VBA parser, at the same rigor level as vba_list_dependencies/vba_list_references (e.g. 'Static Sub Foo()' is correctly excluded as a procedure header rather than a Static variable, but no attempt is made to parse array bounds, string lengths, or line-continuation edge cases beyond the ordinary 'Dim x As Long, y As String' and 'Dim arr(1 To 10, 1 To 5) As Variant'-style comma splitting). If the workbook is already open in Excel, 'workbook' (its display name) is enough. Otherwise pass 'workbookPath' (full file path): Excel will be launched if not running, and the file opened if not already open. Fails with ERR_VBOM_TRUST_DISABLED if Excel's 'Trust access to the VBA project object model' setting is off.",
  {
    workbook: z.string().optional().describe("Workbook display name. Either this or workbookPath is required; workbookPath is preferred since it also auto-launches/opens Excel if needed."),
    workbookPath: z.string().optional().describe("Full path to the workbook file. If set, Excel is auto-launched and the file auto-opened when needed."),
    module: z.string().optional().describe("VBA module name to scan. Omit to scan every module in the workbook in one call (list-declarations mode only -- required when variableName is given)."),
    variableName: z.string().optional().describe("Variable or constant name to find usages of. Provide this to switch from listing declarations to finding usage sites; when given, 'module' becomes required."),
    procedure: z.string().optional().describe("Disambiguates which declaration 'variableName' refers to, when more than one exists in 'module' -- match it against a candidate's declaredIn from the ambiguous_declaration error. Only meaningful together with 'variableName'."),
  },
  async (params) => {
    if (!params.workbook && !params.workbookPath) {
      return { content: [{ type: "text", text: JSON.stringify({ ok: false, error: "workbook or workbookPath is required" }) }], isError: true };
    }
    const wb = psq(params.workbook ?? "");
    const wbPath = psq(params.workbookPath ?? "");

    const readResult = await readAllModulesCode(wb, wbPath);
    if (!readResult.ok) { return { content: readResult.content, isError: readResult.isError }; }

    if (params.variableName) {
      if (!params.module) {
        return { content: [{ type: "text", text: JSON.stringify({ ok: false, error: "module is required when variableName is given" }) }], isError: true };
      }
      const usageResult = resolveVariableUsages(params.module, params.variableName, params.procedure ?? null, readResult.modules);
      if (!usageResult.ok) {
        const errRes: Record<string, unknown> = { ok: false, error: usageResult.error, variableName: usageResult.variableName };
        if (usageResult.error === "ambiguous_declaration") { errRes.candidates = usageResult.candidates; }
        if (usageResult.error === "declaration_not_found") { errRes.module = params.module; }
        return { content: [{ type: "text", text: JSON.stringify(errRes) }], isError: true };
      }
      const writes = usageResult.usages.filter((u) => u.kind === "write").length;
      const usageRes: Record<string, unknown> = {
        ok: true,
        workbook: readResult.workbook,
        mode: "usages",
        declaration: usageResult.declaration,
        usages: usageResult.usages,
        summary: { total: usageResult.usages.length, writes, references: usageResult.usages.length - writes },
      };
      if (readResult.launchedExcelPid) { usageRes.launchedExcelPid = readResult.launchedExcelPid; }
      return { content: [{ type: "text", text: JSON.stringify(usageRes, null, 2) }] };
    }

    let targetModules = readResult.modules;
    if (params.module) {
      targetModules = targetModules.filter((m) => m.name === params.module);
      if (targetModules.length === 0) {
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              ok: false,
              error: "module_not_found",
              module: params.module,
              availableModules: readResult.modules.map((m) => m.name),
            }),
          }],
          isError: true,
        };
      }
    }

    const scans: ModuleVariableScopeScan[] = targetModules.map((m) => scanModuleForVariableScopes(m.name, m.code));
    const nonEmptyScans = scans.filter((s) => s.declarations.length > 0);

    const summary = {
      modulesScanned: targetModules.length,
      modulesWithDeclarations: nonEmptyScans.length,
      procedureScoped: scans.reduce((sum, s) => sum + s.declarations.filter((d) => d.scope === "procedure").length, 0),
      moduleScoped: scans.reduce((sum, s) => sum + s.declarations.filter((d) => d.scope === "module").length, 0),
      publicScoped: scans.reduce((sum, s) => sum + s.declarations.filter((d) => d.scope === "public").length, 0),
    };

    const res: Record<string, unknown> = {
      ok: true,
      workbook: readResult.workbook,
      mode: "declarations",
      summary,
      modules: nonEmptyScans,
    };
    if (readResult.launchedExcelPid) { res.launchedExcelPid = readResult.launchedExcelPid; }
    return { content: [{ type: "text", text: JSON.stringify(res, null, 2) }] };
  }
);

// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ excel_update_module_code ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
// dry-run（プレビュー）→ confirmToken付き呼び出し（実書き込み）の2段階フロー。
// 実際の書き込みは import_single_module.ps1 経由（= import_opened_vba.ps1 の
// VBComponents.Import()ベースのロジックを再利用）で行い、CodeModule.AddFromString()を
// 直接呼ぶことは絶対に行わない（Attribute行を含むコードでコンパイルエラーになるため）。
function computeConfirmToken(module: string, currentCode: string, newCode: string): string {
  return createHash("sha256").update(`${module}\u0000${currentCode}\u0000${newCode}`).digest("hex").slice(0, 16);
}

// Reads the current source of a VBA module. Shared by excel_update_module_code's
// dry-run path and by its confirm/write path (re-read immediately before writing,
// to detect concurrent modification since the dry-run -- see below).
async function readCurrentModuleCode(
  wbEscaped: string,
  wbPathEscaped: string,
  modEscaped: string
): Promise<
  | { ok: true; workbook: string; module: string; componentType: number; currentCode: string; launchedExcelPid?: number }
  | { ok: false; content: { type: "text"; text: string }[]; isError: true }
> {
  const dotSource = dotSourceExcelUtil();
  const psScript = `
$ErrorActionPreference='Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$OutputEncoding           = [Console]::OutputEncoding

${dotSource}

try { $r = Get-OrStartExcelApplication; $excel = $r.App }
catch { @{ ok=$false; error='excel_not_found' } | ConvertTo-Json ; exit }

try { $wb = Resolve-TargetWorkbook -App $excel -WorkbookPath '${wbPathEscaped}' -WorkbookName '${wbEscaped}' }
catch { @{ ok=$false; error="$($_.Exception.Message)" } | ConvertTo-Json ; exit }

try { Test-VbaTrustAccess -Workbook $wb | Out-Null }
catch { @{ ok=$false; error="$($_.Exception.Message)" } | ConvertTo-Json ; exit }

try { $vbc=$wb.VBProject.VBComponents.Item('${modEscaped}') }
catch { @{ ok=$false; error='module_not_found'; module='${modEscaped}' } | ConvertTo-Json ; exit }

try {
  $cm=$vbc.CodeModule
  $code = if ($cm.CountOfLines -gt 0) { $cm.Lines(1, $cm.CountOfLines) } else { "" }
  $readRes = @{ ok=$true; workbook=$wb.Name; module=$vbc.Name; componentType=$vbc.Type; currentCode=$code }
  if ($r.LaunchedProcessId) { $readRes.launchedExcelPid = $r.LaunchedProcessId }
  $readRes | ConvertTo-Json -Depth 6
} catch {
  @{ ok=$false; error='read_failed'; detail="$($_.Exception.Message)" } | ConvertTo-Json
}
`.trim();

  try {
    const { stdout } = await execFileAsync(
      "powershell.exe",
      ["-NoLogo", "-NoProfile", "-NonInteractive", "-STA", "-ExecutionPolicy", "Bypass", "-Command", psScript],
      { windowsHide: true, encoding: "buffer", timeout: 20000, maxBuffer: 2 * 1024 * 1024 }
    );
    const outText = Buffer.isBuffer(stdout) ? stdout.toString("utf8") : String(stdout);
    const classified = classifyResult(outText);
    if (classified.isError) {
      return { ok: false, content: classified.content, isError: true };
    }

    let payload: any = null;
    try {
      const start = Math.min(...['{', '['].map(ch => { const i = outText.indexOf(ch); return i === -1 ? Number.POSITIVE_INFINITY : i; }));
      payload = Number.isFinite(start) ? JSON.parse(outText.slice(start)) : null;
    } catch { /* noop */ }

    if (!payload?.ok) {
      return { ok: false, content: [{ type: "text", text: outText }], isError: true };
    }

    return {
      ok: true,
      workbook: payload.workbook,
      module: payload.module,
      componentType: payload.componentType,
      currentCode: payload.currentCode,
      launchedExcelPid: payload.launchedExcelPid,
    };
  } catch (e: any) {
    return { ok: false, content: [{ type: "text", text: JSON.stringify({ ok: false, error: "ps_failed", detail: String(e?.message ?? e) }) }], isError: true };
  }
}

// Existence check only, used by excel_update_module_code create-mode (moduleType set):
// distinguishes "module not found" (expected -- safe to create) from any other kind of
// failure (Excel not found, workbook not resolved, trust access denied), which must NOT
// be treated as safe-to-create and should propagate as a real error instead.
async function checkModuleExists(
  wbEscaped: string,
  wbPathEscaped: string,
  modEscaped: string
): Promise<
  | { ok: true; workbook: string; exists: boolean; launchedExcelPid?: number }
  | { ok: false; content: { type: "text"; text: string }[]; isError: true }
> {
  const dotSource = dotSourceExcelUtil();
  const psScript = `
$ErrorActionPreference='Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$OutputEncoding           = [Console]::OutputEncoding

${dotSource}

try { $r = Get-OrStartExcelApplication; $excel = $r.App }
catch { @{ ok=$false; error='excel_not_found' } | ConvertTo-Json ; exit }

try { $wb = Resolve-TargetWorkbook -App $excel -WorkbookPath '${wbPathEscaped}' -WorkbookName '${wbEscaped}' }
catch { @{ ok=$false; error="$($_.Exception.Message)" } | ConvertTo-Json ; exit }

try { Test-VbaTrustAccess -Workbook $wb | Out-Null }
catch { @{ ok=$false; error="$($_.Exception.Message)" } | ConvertTo-Json ; exit }

$exists = $true
try { $wb.VBProject.VBComponents.Item('${modEscaped}') | Out-Null } catch { $exists = $false }

$existsRes = @{ ok=$true; workbook=$wb.Name; exists=$exists }
if ($r.LaunchedProcessId) { $existsRes.launchedExcelPid = $r.LaunchedProcessId }
$existsRes | ConvertTo-Json -Depth 6
`.trim();

  try {
    const { stdout } = await execFileAsync(
      "powershell.exe",
      ["-NoLogo", "-NoProfile", "-NonInteractive", "-STA", "-ExecutionPolicy", "Bypass", "-Command", psScript],
      { windowsHide: true, encoding: "buffer", timeout: 20000, maxBuffer: 2 * 1024 * 1024 }
    );
    const outText = Buffer.isBuffer(stdout) ? stdout.toString("utf8") : String(stdout);
    const classified = classifyResult(outText);
    if (classified.isError) {
      return { ok: false, content: classified.content, isError: true };
    }

    let payload: any = null;
    try {
      const start = Math.min(...['{', '['].map(ch => { const i = outText.indexOf(ch); return i === -1 ? Number.POSITIVE_INFINITY : i; }));
      payload = Number.isFinite(start) ? JSON.parse(outText.slice(start)) : null;
    } catch { /* noop */ }

    if (!payload?.ok) {
      return { ok: false, content: [{ type: "text", text: outText }], isError: true };
    }

    return { ok: true, workbook: payload.workbook, exists: !!payload.exists, launchedExcelPid: payload.launchedExcelPid };
  } catch (e: any) {
    return { ok: false, content: [{ type: "text", text: JSON.stringify({ ok: false, error: "ps_failed", detail: String(e?.message ?? e) }) }], isError: true };
  }
}

// Tier 1 static checks only (regex/text-based, no real VBA parser). See the
// excel-vba-sync-dev skill's references/vba-lint-tiers.md for the full 40-rule
// catalog this is drawn from, and why Tier 2/3 rules are deferred (they need
// per-procedure boundary tracking or real dataflow analysis to stay low-noise;
// this simpler pass would false-positive/negative too often on those).
interface VbaLintWarning {
  ruleId: string;
  severity: "Warning" | "Info";
  line: number;
  message: string;
}

function lintVbaCode(code: string): VbaLintWarning[] {
  const warnings: VbaLintWarning[] = [];
  const lines = code.split(/\r\n|\r|\n/);

  const simplePatterns: { ruleId: string; severity: "Warning" | "Info"; regex: RegExp; message: string }[] = [
    { ruleId: "VBA001", severity: "Warning", regex: /\.Select\b/i, message: "Select relies on selection state and can make behavior unstable -- operate on the target object directly instead." },
    { ruleId: "VBA002", severity: "Warning", regex: /\.Activate\b/i, message: "Activate depends on the active sheet/workbook -- reference the target via a Worksheet variable instead." },
    { ruleId: "VBA003", severity: "Warning", regex: /\bSelection\b/i, message: "Selection is not guaranteed to be the intended target -- operate on an explicit Range instead." },
    { ruleId: "VBA004", severity: "Warning", regex: /\bActiveSheet\b/i, message: "ActiveSheet can change due to user action or other code -- reference the target sheet explicitly instead." },
    { ruleId: "VBA005", severity: "Warning", regex: /\bActiveWorkbook\b/i, message: "ActiveWorkbook may not be the intended workbook -- use ThisWorkbook or an explicit workbook variable instead." },
    { ruleId: "VBA021", severity: "Warning", regex: /\bUsedRange\b/i, message: "UsedRange can include stale formatting or deleted cells -- compute the last row/column from a specific column instead." },
    { ruleId: "VBA028", severity: "Warning", regex: /^\s*End\s*(?:'.*)?$/i, message: "A bare End statement skips cleanup/restore code -- use Exit Sub or a unified cleanup routine instead." },
    { ruleId: "VBA035", severity: "Info", regex: /^\s*Call\s+/i, message: "Call is legacy VBA style -- calling the procedure directly is more idiomatic." },
    { ruleId: "VBA039", severity: "Warning", regex: /\bAs\s+#\s*\d+/i, message: "A hardcoded file number can collide with other open files -- use FreeFile instead." },
  ];

  lines.forEach((line, idx) => {
    const lineNo = idx + 1;

    for (const p of simplePatterns) {
      if (p.regex.test(line)) {
        warnings.push({ ruleId: p.ruleId, severity: p.severity, line: lineNo, message: p.message });
      }
    }

    // VBA037: Declare ... without PtrSafe (won't compile under 64-bit Office)
    if (/\bDeclare\s+(Function|Sub)\b/i.test(line) && !/\bPtrSafe\b/i.test(line)) {
      warnings.push({ ruleId: "VBA037", severity: "Warning", line: lineNo, message: "A Declare statement without PtrSafe won't compile under 64-bit Office -- add PtrSafe (and LongPtr where needed)." });
    }

    // VBA034: assignment to CreateObject/New/GetObject without a leading Set
    if (/^\s*(?!Set\b)[A-Za-z_][\w.]*\s*=\s*(New\s+[A-Za-z_]\w*|CreateObject\s*\(|GetObject\s*\()/i.test(line)) {
      warnings.push({ ruleId: "VBA034", severity: "Warning", line: lineNo, message: "Assigning an object (CreateObject/New/GetObject) without Set is invalid -- add Set." });
    }
  });

  // VBA009: Option Explicit missing anywhere in the module
  if (!/^\s*Option\s+Explicit\b/im.test(code)) {
    warnings.push({ ruleId: "VBA009", severity: "Warning", line: 1, message: "Option Explicit is missing -- typos in variable names won't be caught at compile time." });
  }

  // VBA030: procedures longer than 200 lines
  const procStart = /^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)/i;
  const procEnd = /^\s*End\s+(?:Sub|Function|Property)\b/i;
  let openLine: number | null = null;
  let openName = "";
  lines.forEach((line, idx) => {
    const lineNo = idx + 1;
    if (openLine === null) {
      const m = procStart.exec(line);
      if (m) { openLine = lineNo; openName = m[1]; }
    } else if (procEnd.test(line)) {
      const len = lineNo - openLine + 1;
      if (len > 200) {
        warnings.push({ ruleId: "VBA030", severity: "Info", line: openLine, message: `Procedure "${openName}" is ${len} lines long -- consider splitting it into smaller procedures for maintainability.` });
      }
      openLine = null;
      openName = "";
    }
  });

  return warnings;
}

server.tool(
  "excel_update_module_code",
  "Overwrite the code of an EXISTING VBA module, or create a brand-new one when moduleType is set (cannot target .frm UserForm modules either way -- overwriting one fails with ERR_UNSUPPORTED_MODULE_TYPE, and UserForms cannot be created via this tool at all). REQUIRED two-step flow: (1) call once with dryRun:true to preview a diff against the current code and receive a confirmToken; (2) call again with the exact same workbook/workbookPath, module and newCode plus that confirmToken to actually write -- calling without a valid confirmToken is rejected. The confirmToken is bound to the module's code as it was at dry-run time: immediately before writing, the tool re-reads the module and recomputes the token -- if the code changed since the dry-run (e.g. another client wrote to it first), the write is rejected with ERR_MODULE_CHANGED_SINCE_DRYRUN instead of silently overwriting that change. A timestamped backup of the code being replaced is always written to '<workbook folder>/.excel-vba-sync-backups' before the write happens. If the target is a Sheet/ThisWorkbook code-behind module (componentType 100), per-procedure Attribute lines such as an assigned macro shortcut key CANNOT be preserved (VBA API limitation, not a bug) -- check willLoseShortcutAttributes in the dry-run response before proceeding on such modules. The dry-run response also includes lintWarnings: a best-effort, regex-based static check (not a real VBA parser) for a small set of common issues -- Select/Activate/Selection/ActiveSheet/ActiveWorkbook usage, missing Option Explicit, UsedRange, bare End statements, overly long procedures, missing Set before an object assignment, Declare without PtrSafe, and hardcoded file numbers. These are advisory only and never block the write. To create a new module instead of overwriting one, pass moduleType ('standard' or 'class') together with a module name that does not yet exist. The dry-run response then reports mode as 'create', and the confirmToken is bound to (module, moduleType, newCode) rather than to any existing code, since none exists yet. If a module with that name is created by someone else between the dry-run and the confirming call, the write is rejected with ERR_MODULE_ALREADY_EXISTS_SINCE_DRYRUN instead of colliding with it. No backup is written when creating a new module (there is nothing to back up), and the response's componentType is 1 for a standard module or 2 for a class module.",
  {
    workbook: z.string().optional().describe("Workbook display name. Either this or workbookPath is required; workbookPath is preferred since it also auto-launches/opens Excel if needed."),
    workbookPath: z.string().optional().describe("Full path to the workbook file. If set, Excel is auto-launched and the file auto-opened when needed."),
    module: z.string().describe("Name of a VBA module. When moduleType is omitted, this must be an EXISTING module to overwrite. When moduleType is set, this is the name of a NEW module to create -- it must not already exist."),
    moduleType: z.enum(["standard", "class"]).optional().describe("Set this to create module NAME as a brand-new module instead of overwriting an existing one. 'standard' = .bas (StdModule), 'class' = .cls (Class module). UserForm modules cannot be created this way. Omit this parameter entirely to overwrite an existing module (unchanged default behavior)."),
    newCode: z.string().describe("Full replacement source code for the module (procedure bodies only -- do not include Attribute lines)."),
    dryRun: z.boolean().optional().describe("If true, only preview the change (or, in create mode, the module that would be created) and return a confirmToken; does not write anything."),
    confirmToken: z.string().optional().describe("Token obtained from a prior dryRun:true call with the identical workbook/module/moduleType/newCode. Required to actually perform the write. Rejected with ERR_MODULE_CHANGED_SINCE_DRYRUN in overwrite mode, or ERR_MODULE_ALREADY_EXISTS_SINCE_DRYRUN in create mode, if the target changed since the dry-run."),
  },
  async (params) => {
    const wb = psq(params.workbook ?? "");
    const wbPath = psq(params.workbookPath ?? "");
    const mod = psq(params.module);

    if (!params.workbook && !params.workbookPath) {
      return { content: [{ type: "text", text: JSON.stringify({ ok: false, error: "workbook or workbookPath is required" }) }], isError: true };
    }

    // --- dry-run: read the current code, compute a confirmToken bound to it, do not write anything ---
    if (params.dryRun) {
      if (params.moduleType) {
        // --- create-mode dry-run: the module must NOT already exist ---
        const existsResult = await checkModuleExists(wb, wbPath, mod);
        if (!existsResult.ok) { return { content: existsResult.content, isError: existsResult.isError }; }
        if (existsResult.exists) {
          return { content: [{ type: "text", text: JSON.stringify({ ok: false, error: "module_already_exists", module: params.module }) }], isError: true };
        }
        const createToken = computeConfirmToken(params.module, "__CREATE__", params.newCode);
        const createPreview: Record<string, unknown> = {
          ok: true,
          mode: "create",
          workbook: existsResult.workbook,
          module: params.module,
          moduleType: params.moduleType,
          newCode: params.newCode,
          lintWarnings: lintVbaCode(params.newCode),
          confirmToken: createToken,
          note: "Call this tool again with the same workbook/module/moduleType/newCode and this confirmToken to create the module. If a module with this name gets created by someone else before that call, the write will be rejected with ERR_MODULE_ALREADY_EXISTS_SINCE_DRYRUN instead of colliding with it.",
        };
        if (existsResult.launchedExcelPid) { createPreview.launchedExcelPid = existsResult.launchedExcelPid; }
        return { content: [{ type: "text", text: JSON.stringify(createPreview, null, 2) }] };
      }

      const readResult = await readCurrentModuleCode(wb, wbPath, mod);
      if (!readResult.ok) { return { content: readResult.content, isError: readResult.isError }; }

      const expectedToken = computeConfirmToken(params.module, readResult.currentCode, params.newCode);
      const preview: Record<string, unknown> = {
        ok: true,
        workbook: readResult.workbook,
        module: readResult.module,
        componentType: readResult.componentType,
        currentCode: readResult.currentCode,
        newCode: params.newCode,
        willLoseShortcutAttributes: readResult.componentType === 100,
        lintWarnings: lintVbaCode(params.newCode),
        confirmToken: expectedToken,
        note: "Call this tool again with the same workbook/module/newCode and this confirmToken to apply the write. If the module's code changes before that call (e.g. another client writes to it first), the token will no longer match and the write will be rejected rather than silently overwriting that change.",
      };
      if (readResult.launchedExcelPid) { preview.launchedExcelPid = readResult.launchedExcelPid; }
      return { content: [{ type: "text", text: JSON.stringify(preview, null, 2) }] };
    }

    // --- confirmToken required (prevents a blind write without having previewed via dry-run) ---
    if (!params.confirmToken) {
      return {
        content: [{ type: "text", text: JSON.stringify({ ok: false, error: "confirmToken is required. Call this tool with dryRun:true first to preview the change and obtain a confirmToken." }) }],
        isError: true,
      };
    }

    // --- re-verify immediately before writing (optimistic concurrency check), branching on
    // whether this is a create (moduleType set) or an overwrite (moduleType omitted). ---
    if (params.moduleType) {
      const existsResult = await checkModuleExists(wb, wbPath, mod);
      if (!existsResult.ok) { return { content: existsResult.content, isError: existsResult.isError }; }
      const createToken = computeConfirmToken(params.module, "__CREATE__", params.newCode);
      if (existsResult.exists || params.confirmToken !== createToken) {
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              ok: false,
              error: "ERR_MODULE_ALREADY_EXISTS_SINCE_DRYRUN",
              detail: "A module with this name already exists (created since the dry-run, e.g. by another client), or the confirmToken does not match this workbook/module/moduleType/newCode. Re-run with dryRun:true to get a fresh token before retrying, or omit moduleType to overwrite the existing module instead.",
            }),
          }],
          isError: true,
        };
      }
    } else {
      // if it no longer matches what the dry-run saw, someone else changed it in the meantime --
      // reject instead of silently overwriting that change. ---
      const freshRead = await readCurrentModuleCode(wb, wbPath, mod);
      if (!freshRead.ok) { return { content: freshRead.content, isError: freshRead.isError }; }

      const expectedToken = computeConfirmToken(params.module, freshRead.currentCode, params.newCode);
      if (params.confirmToken !== expectedToken) {
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              ok: false,
              error: "ERR_MODULE_CHANGED_SINCE_DRYRUN",
              detail: "The module's code has changed since the dry-run (or the confirmToken does not match this workbook/module/newCode). Re-run with dryRun:true to get a fresh token before retrying, to avoid overwriting a change made by another client in the meantime.",
            }),
          }],
          isError: true,
        };
      }
    }

    // --- perform the write via import_single_module.ps1 (reuses Import-ModuleToVBProject; never
    // call CodeModule.AddFromString() directly here -- it rejects Attribute lines, see issue #3) ---
    const scriptsDir = getScriptsDir();
    if (!scriptsDir) {
      return { content: [{ type: "text", text: JSON.stringify({ ok: false, error: "MCP_SCRIPTS_DIR/MCP_PS_LIST not set" }) }], isError: true };
    }
    const singleModuleScript = path.join(scriptsDir, "import_single_module.ps1");
    if (!fs.existsSync(singleModuleScript)) {
      return { content: [{ type: "text", text: JSON.stringify({ ok: false, error: `script not found: ${singleModuleScript}` }) }], isError: true };
    }

    const tmpFile = path.join(os.tmpdir(), `vba_mcp_write_${Date.now()}_${Math.random().toString(36).slice(2)}.txt`);
    fs.writeFileSync(tmpFile, params.newCode, { encoding: "utf8" });

    try {
      const args = [
        "-NoLogo", "-NoProfile", "-NonInteractive", "-STA", "-ExecutionPolicy", "Bypass",
        "-File", singleModuleScript,
        "-WorkbookPath", params.workbookPath || "",
        "-WorkbookName", params.workbook || "",
        "-ModuleName", params.module,
        "-SourceCodePath", tmpFile,
        "-ScriptsDir", scriptsDir,
      ];
      if (params.moduleType) { args.push("-ModuleType", params.moduleType); }
      const { stdout } = await execFileAsync("powershell.exe", args, {
        windowsHide: true,
        encoding: "buffer",
        timeout: 30000,
        maxBuffer: 2 * 1024 * 1024,
      });
      const outText = Buffer.isBuffer(stdout) ? stdout.toString("utf8") : String(stdout);
      return classifyResult(outText);
    } catch (e: any) {
      return extractFailureResult(e);
    } finally {
      try { fs.unlinkSync(tmpFile); } catch { /* noop */ }
    }
  }
);
console.log("# vba-excel-mcp server: ready");
