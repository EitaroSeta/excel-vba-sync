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
import { findCrossModuleDuplicates, CrossModuleDuplicate } from "./duplicateProcedureScan.js";
import { redactSecrets, redactCodeText } from "./secretRedaction.js";
import { classifyResult, classifyResultWithRedaction } from "./responseClassification.js";
import { evaluateExportGuard } from "./exportGuard.js";
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

// serverInfo.version をハードコードすると更新を忘れる（実際、package.json が
// 0.0.7x まで進む間ずっと "0.1.0" のまま乖離していた）。package.json から読んで
// 自動追従させる。読めなかった場合だけ "0.0.0" にフォールバックする。
function readPackageVersion(): string {
  try {
    const pkg = JSON.parse(fs.readFileSync(path.join(__dirname, "..", "package.json"), "utf8"));
    if (typeof pkg.version === "string" && pkg.version.length > 0) { return pkg.version; }
  } catch { /* fall through to the placeholder below */ }
  return "0.0.0";
}

console.log("# vba-excel-mcp server: booting...");

const server = new McpServer(
  { name: "vba-excel-mcp", title: "Excel VBA Sync", version: readPackageVersion() },
  {
    instructions:
      "Read, search, analyze, run macros in, and write to Excel workbooks on this Windows machine, via COM automation. " +
      "Prefer workbookPath (full file path) over workbook (display name) -- it auto-launches Excel and auto-opens the file if needed. " +
      "Coverage is not limited to VBA source: the workbook's own logic can be inspected too -- cell formulas (excel_list_formulas), " +
      "conditional formatting (excel_list_conditional_formats) and data validation incl. dropdown choices (excel_list_data_validations). " +
      "When investigating or migrating an Excel application, check these as well as the VBA code: a macro's output is often further " +
      "processed by spreadsheet formulas and formatting rules that appear nowhere in the VBA, and excel_read_range shows only the " +
      "resulting values, never the formula behind them. " +
      "Before generating code that references a sheet, defined name or form control by name, confirm it actually exists " +
      "(excel_list_worksheets / excel_list_defined_names / excel_list_form_controls) -- a wrong name fails at RUNTIME, not at write time. " +
      "excel_update_module_code requires dryRun:true first to preview and get a confirmToken, then a second call with that token to actually write; " +
      "the token is rejected with ERR_MODULE_CHANGED_SINCE_DRYRUN if the module changed in the meantime. " +
      "A backup of the replaced code is always written before any change, and the workbook is never saved to disk automatically -- " +
      "tell the user their change is not yet persisted. UserForms can have their code-behind overwritten but their layout is never touched, " +
      "and new UserForms cannot be created. " +
      "If the user keeps exported .bas/.cls/.frm files (for editing in VS Code or git), call excel_export_module right after each confirmed write -- " +
      "otherwise their on-disk copy is stale and a later manual Import will silently revert your change. " +
      "If a written change is tested and rejected: since nothing was auto-saved, closing Excel WITHOUT saving and reopening restores the workbook's last " +
      "saved state; alternatively rewrite the module from the pre-change backup in .excel-vba-sync-backups. Either way, call excel_export_module again " +
      "afterwards so the exported file matches the restored code. " +
      "If a workbook or module the user named does not exist (ERR_WORKBOOK_NOT_FOUND, or missing from excel_list_modules), do NOT silently substitute " +
      "a similar-looking one you found yourself -- a one-character difference can be a different file, not a typo. Say what you found and get the user's " +
      "confirmation BEFORE any write. Reading to help identify the right target is fine; writing to a guessed target is not. " +
      "Values that look like hardcoded passwords or API keys are always masked as [REDACTED] in any tool output that returns code text. " +
      "This server intentionally enforces no coding conventions (naming, error handling, form structure) -- follow whatever conventions " +
      "the caller's own instructions define.",
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

// FindAndRun-ExcelMacroByModule.ps1 の絶対パスを解決する。
// MCP_PS_LIST / MCP_PS_RUN は「スクリプトファイルそのもの」を指す旧来の環境変数で、
// excel_list_macros と excel_run_macro だけがこれらを直接読んでいた。そのため
// MCP_SCRIPTS_DIR しか渡されない起動経路では、この2ツールだけが動かない
// （他のツールは getScriptsDir() 経由なので正常）という、気づきにくい不具合が
// 起きていた（v0.0.41で実際に発生）。ここで getScriptsDir() を最終フォールバックに
// 加え、全ツールで解決方法を揃える。
function resolveMacroScript(preferRun: boolean): { ok: true; path: string } | { ok: false; error: string } {
  const candidates: string[] = [];
  if (preferRun && process.env.MCP_PS_RUN) { candidates.push(process.env.MCP_PS_RUN); }
  if (process.env.MCP_PS_LIST) { candidates.push(process.env.MCP_PS_LIST); }
  const dir = getScriptsDir();
  if (dir) { candidates.push(path.join(dir, "FindAndRun-ExcelMacroByModule.ps1")); }
  for (const c of candidates) {
    if (fs.existsSync(c)) { return { ok: true, path: c }; }
  }
  return {
    ok: false,
    error: candidates.length
      ? `ps1 not found: ${candidates.join(" | ")}`
      : "MCP_SCRIPTS_DIR / MCP_PS_LIST not set",
  };
}

// ExcelUtil.ps1 を dot-source する行を生成（見つからない場合は空文字＝スキップ）
function dotSourceExcelUtil(): string {
  const dir = getScriptsDir();
  if (!dir) { return ""; }
  const utilPath = path.join(dir, "ExcelUtil.ps1");
  if (!fs.existsSync(utilPath)) { return ""; }
  return `. '${psq(utilPath)}'`;
}

// classifyResult / classifyResultWithRedaction moved to responseClassification.ts (v0.0.73)
// so their branching logic can be covered by the persistent node:test suite.


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
  "Read a VBA module's full source, as the VBE shows it. Attribute lines (module- and procedure-level, e.g. macro shortcut key bindings) are NOT included -- this reads via CodeModule.Lines(). " +
  "Values resembling a hardcoded password/API key/Authorization header are masked as [REDACTED] -- always on, best-effort. " +
  "Pass workbookPath (full path) unless the workbook is already open. Fails with ERR_VBOM_TRUST_DISABLED if Excel's 'Trust access to the VBA project object model' is off (it cannot be enabled programmatically).",
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
      return classifyResultWithRedaction(outText, { stringFields: ["code"] });
    } catch (e: any) {
      return { content: [{ type: "text", text: JSON.stringify({ ok:false, error:"ps_failed", detail:String(e?.message ?? e) }) }] };
    }
  }
);

server.tool(
  "excel_list_modules",
  "List a workbook's VBA modules (name, component type, line count) without reading any code -- cheap and fast. " +
  "Use this, not vba_search_code, when you only need to know what modules exist before deciding what to read, search or run.",
  {
    workbook: z.string().optional().describe("Workbook display name. Give this or workbookPath; workbookPath is preferred (it can auto-launch Excel and open the file)."),
    workbookPath: z.string().optional().describe("Full path to the workbook. Auto-launches Excel and opens the file if needed."),
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

// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ excel_list_worksheets ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
server.tool(
  "excel_list_worksheets",
  "List a workbook's actual worksheets (display name, VBA CodeName, index, visibility) -- not the code modules; use excel_list_modules for those. " +
  "A sheet has two names: the display name (used by Worksheets('...'), renamable) and the CodeName (used by direct Sheet1.Range(...) references, fixed at creation, and what excel_list_modules shows for componentType 100 entries). " +
  "Call this before writing code that names a sheet: a mismatch fails at RUNTIME, not at write time, and nothing else catches it in advance. " +
  "If a sheet is missing, do not try to create or rename one -- ask the user to do it in Excel. " +
  "Does NOT require the VBA Trust Center setting.",
  {
    workbook: z.string().optional().describe("Workbook display name. Give this or workbookPath; workbookPath is preferred (it can auto-launch Excel and open the file)."),
    workbookPath: z.string().optional().describe("Full path to the workbook. Auto-launches Excel and opens the file if needed."),
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

try {
  $sheets = @()
  foreach ($ws in @($wb.Worksheets)) {
    $vis = switch ($ws.Visible) {
      -1 { 'visible' }
      0 { 'hidden' }
      2 { 'veryHidden' }
      default { 'unknown' }
    }
    $sheets += [pscustomobject]@{ name=$ws.Name; codeName=$ws.CodeName; index=$ws.Index; visible=$vis }
  }
  $res = @{ ok=$true; workbook=$wb.Name; worksheets=$sheets; count=$sheets.Count }
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

// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ excel_list_defined_names ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
server.tool(
  "excel_list_defined_names",
  "List a workbook's actual defined names (named ranges), both workbook- and sheet-scoped. " +
  "Call this before writing code that references one (e.g. Range('MyNamedRange')): vba_list_references only reports likely references found in VBA code text, never the workbook's real Names collection. " +
  "Each entry has refersTo and isBroken (true when refersTo contains #REF!, i.e. the name points at something deleted -- these fail silently otherwise). " +
  "If a name is missing, do not try to create it -- ask the user to define it (Formulas > Name Manager). " +
  "Does NOT require the VBA Trust Center setting.",
  {
    workbook: z.string().optional().describe("Workbook display name. Give this or workbookPath; workbookPath is preferred (it can auto-launch Excel and open the file)."),
    workbookPath: z.string().optional().describe("Full path to the workbook. Auto-launches Excel and opens the file if needed."),
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

try {
  $names = @()
  foreach ($n in @($wb.Names)) {
    $scope = 'workbook'
    $scopeSheet = $null
    try { if ($n.Parent.Name -ne $wb.Name) { $scope = 'sheet'; $scopeSheet = $n.Parent.Name } } catch {}
    $refersTo = $null
    $isBroken = $false
    try { $refersTo = $n.RefersTo } catch { $isBroken = $true }
    if ($refersTo -and $refersTo -like '*#REF!*') { $isBroken = $true }
    $vis = $true
    try { $vis = [bool]$n.Visible } catch {}
    $names += [pscustomobject]@{ name=$n.Name; refersTo=$refersTo; scope=$scope; scopeSheet=$scopeSheet; visible=$vis; isBroken=$isBroken }
  }
  $res = @{ ok=$true; workbook=$wb.Name; definedNames=$names; count=$names.Count }
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

// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ excel_list_form_controls ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
server.tool(
  "excel_list_form_controls",
  "List the controls inside one or every UserForm -- name and type only, not layout or position. " +
  "Call this before writing code that references a control (e.g. UserForm1.TextBox1): nothing else can confirm a control exists, since excel_get_module_code returns only the form's code-behind, not its designer layout (that lives in the binary .frx). " +
  "If the form or control is missing, do not try to create it -- ask the user to add it in the VBE form designer. " +
  "type is given the familiar VBA name where known; anything else comes back as its raw COM interface name (e.g. IMdcListBox), which is still correct in substance, just not normalized yet. " +
  "Requires the VBA Trust Center setting.",
  {
    workbook: z.string().optional().describe("Workbook display name. Give this or workbookPath; workbookPath is preferred (it can auto-launch Excel and open the file)."),
    workbookPath: z.string().optional().describe("Full path to the workbook. Auto-launches Excel and opens the file if needed."),
    formName: z.string().optional().describe("Name of one UserForm to inspect. Omit to list controls for every UserForm in the project in a single call."),
  },
  async (params) => {
    if (!params.workbook && !params.workbookPath) {
      return { content: [{ type: "text", text: JSON.stringify({ ok: false, error: "workbook or workbookPath is required" }) }], isError: true };
    }
    const wb = psq(params.workbook ?? "");
    const wbPath = psq(params.workbookPath ?? "");
    const formName = psq(params.formName ?? "");
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

$targetName = '${formName}'
$comps = @($wb.VBProject.VBComponents)

if ($targetName) {
  $match = $comps | Where-Object { $_.Name -eq $targetName }
  if (-not $match) { @{ ok=$false; error='form_not_found'; formName=$targetName } | ConvertTo-Json ; exit }
  if ($match.Type -ne 3) { @{ ok=$false; error='ERR_NOT_A_USERFORM'; formName=$targetName; componentType=$match.Type } | ConvertTo-Json ; exit }
  $forms = @($match)
} else {
  $forms = @($comps | Where-Object { $_.Type -eq 3 })
}

try {
  $result = @()
  foreach ($f in $forms) {
    $ctrls = @()
    try {
      foreach ($ctrl in @($f.Designer.Controls)) {
        $ctrlType = [Microsoft.VisualBasic.Information]::TypeName($ctrl)
        # TypeName() on a late-bound MSForms control returns its raw internal COM interface/
        # coclass name (e.g. "IMdcCombo"), NOT the friendly VBA name ("ComboBox") that VBA's
        # own TypeName() shows -- confirmed live 2026-08-05. Normalize the handful of control
        # types actually confirmed against a real form; anything else is left as the raw name
        # rather than guessed, since the naming pattern is not consistent enough to extrapolate
        # safely (ILabelControl / IMdcCombo / IMdcOptionButton / ImageClass are three different
        # naming conventions for four controls).
        $knownTypeMap = @{
          'ILabelControl'    = 'Label'
          'IMdcCombo'        = 'ComboBox'
          'IMdcOptionButton' = 'OptionButton'
          'ImageClass'       = 'Image'
          'IMdcText'         = 'TextBox'
          'ITabStrip'        = 'TabStrip'
          'IMultiPage'       = 'MultiPage'
          'IScrollbar'       = 'ScrollBar'
        }
        if ($knownTypeMap.ContainsKey($ctrlType)) { $ctrlType = $knownTypeMap[$ctrlType] }
        $ctrls += [pscustomobject]@{ name=$ctrl.Name; type=$ctrlType }
      }
    } catch {}
    $result += [pscustomobject]@{ form=$f.Name; controls=$ctrls }
  }
  $res = @{ ok=$true; workbook=$wb.Name; forms=$result }
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
  "Read cell values from a worksheet range -- e.g. to verify what excel_run_macro actually did, since that tool cannot confirm its own effect. " +
  "Returns a row-major 2D array via Range.Value2 (avoiding the Date/Currency wrapping Range.Value introduces). Values only, never the formulas behind them -- use excel_list_formulas for those. " +
  "Whole columns/rows can be slow; prefer a bounded address like 'A1:C10'. Does NOT require the VBA Trust Center setting.",
  {
    workbook: z.string().optional().describe("Workbook display name. Give this or workbookPath; workbookPath is preferred (it can auto-launch Excel and open the file)."),
    workbookPath: z.string().optional().describe("Full path to the workbook. Auto-launches Excel and opens the file if needed."),
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

// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ excel_list_formulas ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
server.tool(
  "excel_list_formulas",
  "List the formulas actually present in a worksheet's cells -- VLOOKUP/INDEX-MATCH and similar logic that turns a macro's output into what the user actually sees. A common migration blind spot: excel_read_range returns computed VALUES only, and vba_list_references scans VBA code text only, so neither reveals that a cell holds a formula at all. " +
  "Cells are grouped by FormulaR1C1 (Excel's position-independent form), so a formula filled down thousands of rows collapses into ONE group -- read cellCount as 'this many cells share this pattern', not as that many formulas to review. exampleFormula is ordinary A1 style; addresses lists up to 20 cells (addressesTruncated:true beyond that). " +
  "No formulas returns an empty formulaGroups array, not an error. Does NOT require the VBA Trust Center setting -- this touches only Worksheet/Range, not VBProject.",
  {
    workbook: z.string().optional().describe("Workbook display name. Give this or workbookPath; workbookPath is preferred (it can auto-launch Excel and open the file)."),
    workbookPath: z.string().optional().describe("Full path to the workbook. Auto-launches Excel and opens the file if needed."),
    sheet: z.string().describe("Worksheet name to scan for formulas."),
    range: z.string().optional().describe("Cell range address, e.g. 'A1:D100'. Omit to scan the sheet's entire UsedRange -- prefer a bounded range on very large sheets for speed."),
  },
  async (params) => {
    if (!params.workbook && !params.workbookPath) {
      return { content: [{ type: "text", text: JSON.stringify({ ok: false, error: "workbook or workbookPath is required" }) }], isError: true };
    }
    const wb = psq(params.workbook ?? "");
    const wbPath = psq(params.workbookPath ?? "");
    const sheet = psq(params.sheet);
    const range = psq(params.range ?? "");
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

$targetRangeAddr = '${range}'
try {
  if ($targetRangeAddr) { $scanRange = $ws.Range($targetRangeAddr) } else { $scanRange = $ws.UsedRange }
} catch {
  @{ ok=$false; error='invalid_range'; range=$targetRangeAddr } | ConvertTo-Json ; exit
}

try {
  $formulaCells = $null
  try { $formulaCells = $scanRange.SpecialCells(-4123) } catch { $formulaCells = $null }

  $groups = [ordered]@{}
  $totalCells = 0
  if ($formulaCells) {
    foreach ($cell in $formulaCells) {
      $totalCells++
      $key = $cell.FormulaR1C1
      if (-not $groups.Contains($key)) {
        $groups[$key] = [pscustomobject]@{
          formulaR1C1 = $key
          exampleFormula = $cell.Formula
          exampleAddress = $cell.Address($false, $false)
          addresses = New-Object System.Collections.ArrayList
          cellCount = 0
        }
      }
      $groups[$key].cellCount++
      if ($groups[$key].addresses.Count -lt 20) {
        [void]$groups[$key].addresses.Add($cell.Address($false, $false))
      }
    }
  }

  $formulaGroups = @()
  $groupCount = 0
  $truncated = $false
  foreach ($key in $groups.Keys) {
    $groupCount++
    if ($groupCount -gt 500) { $truncated = $true; break }
    $g = $groups[$key]
    $formulaGroups += [pscustomobject]@{
      formulaR1C1 = $g.formulaR1C1
      exampleFormula = $g.exampleFormula
      exampleAddress = $g.exampleAddress
      addresses = @($g.addresses)
      addressesTruncated = ($g.cellCount -gt 20)
      cellCount = $g.cellCount
    }
  }

  $res = @{ ok=$true; workbook=$wb.Name; sheet=$ws.Name; formulaGroups=$formulaGroups; totalFormulaCells=$totalCells; groupCount=$formulaGroups.Count; truncated=$truncated }
  if ($r.LaunchedProcessId) { $res.launchedExcelPid = $r.LaunchedProcessId }
  $res | ConvertTo-Json -Depth 8
} catch {
  @{ ok=$false; error='list_failed'; detail="$($_.Exception.Message)" } | ConvertTo-Json
}
`.trim();

    try {
      const { stdout } = await execFileAsync(
        "powershell.exe",
        ["-NoLogo", "-NoProfile", "-NonInteractive", "-STA", "-ExecutionPolicy", "Bypass", "-Command", psScript],
        { windowsHide: true, encoding: "buffer", timeout: 20000, maxBuffer: 4 * 1024 * 1024 }
      );
      const outText = Buffer.isBuffer(stdout) ? stdout.toString("utf8") : String(stdout);
      return classifyResult(outText);
    } catch (e: any) {
      return extractFailureResult(e);
    }
  }
);

// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ excel_list_conditional_formats ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
server.tool(
  "excel_list_conditional_formats",
  "List conditional formatting rules in a worksheet's cells -- a cell's color or style can depend on rule logic ('red if negative') that appears nowhere in VBA code or in the cell's formula. " +
  "Cells are grouped by type, operator and position-normalized Formula1/Formula2, so one rule applied down many rows is a single group with a cellCount rather than one entry per cell. " +
  "type/operator get readable names where known; an unrecognized value comes back as its raw number rather than a guess -- still correct, just unnamed. " +
  "No rules returns an empty formatGroups array, not an error. Does NOT require the VBA Trust Center setting.",
  {
    workbook: z.string().optional().describe("Workbook display name. Give this or workbookPath; workbookPath is preferred (it can auto-launch Excel and open the file)."),
    workbookPath: z.string().optional().describe("Full path to the workbook. Auto-launches Excel and opens the file if needed."),
    sheet: z.string().describe("Worksheet name to scan for conditional formatting rules."),
    range: z.string().optional().describe("Cell range address, e.g. 'A1:D100'. Omit to scan the sheet's entire UsedRange -- prefer a bounded range on very large sheets for speed."),
  },
  async (params) => {
    if (!params.workbook && !params.workbookPath) {
      return { content: [{ type: "text", text: JSON.stringify({ ok: false, error: "workbook or workbookPath is required" }) }], isError: true };
    }
    const wb = psq(params.workbook ?? "");
    const wbPath = psq(params.workbookPath ?? "");
    const sheet = psq(params.sheet);
    const range = psq(params.range ?? "");
    const dotSource = dotSourceExcelUtil();

    const psScript = `
$ErrorActionPreference='Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$OutputEncoding           = [Console]::OutputEncoding

${dotSource}

function Convert-ToR1C1 {
  param($App, $Formula, $Cell)
  if (-not $Formula) { return $null }
  try { return $App.ConvertFormula($Formula, 1, -4150, 4, $Cell) } catch { return $Formula }
}

$typeMap = @{
  1='CellValue'; 2='Expression'; 3='ColorScale'; 4='DataBar'; 5='Top10'; 6='IconSet';
  8='UniqueValues'; 9='TextString'; 10='Blanks'; 11='TimePeriod'; 12='AboveAverage';
  13='NoBlanks'; 16='Errors'; 17='NoErrors'
}
$opMap = @{ 1='Between'; 2='NotBetween'; 3='Equal'; 4='NotEqual'; 5='Greater'; 6='Less'; 7='GreaterEqual'; 8='LessEqual' }

try { $r = Get-OrStartExcelApplication; $excel = $r.App }
catch { @{ ok=$false; error='excel_not_found' } | ConvertTo-Json ; exit }

try { $wb = Resolve-TargetWorkbook -App $excel -WorkbookPath '${wbPath}' -WorkbookName '${wb}' }
catch { @{ ok=$false; error="$($_.Exception.Message)" } | ConvertTo-Json ; exit }

try { $ws = $wb.Worksheets.Item('${sheet}') }
catch { @{ ok=$false; error='sheet_not_found'; sheet='${sheet}' } | ConvertTo-Json ; exit }

$targetRangeAddr = '${range}'
try {
  if ($targetRangeAddr) { $scanRange = $ws.Range($targetRangeAddr) } else { $scanRange = $ws.UsedRange }
} catch {
  @{ ok=$false; error='invalid_range'; range=$targetRangeAddr } | ConvertTo-Json ; exit
}

try {
  $fcCells = $null
  try { $fcCells = $scanRange.SpecialCells(-4172) } catch { $fcCells = $null }

  $groups = [ordered]@{}
  $totalCells = 0
  if ($fcCells) {
    foreach ($cell in $fcCells) {
      foreach ($cond in $cell.FormatConditions) {
        $totalCells++
        $typeNum = $cond.Type
        $typeName = if ($typeMap.Contains($typeNum)) { $typeMap[$typeNum] } else { $typeNum }
        $opNum = $null
        try { $opNum = $cond.Operator } catch {}
        $opName = if ($opNum -and $opMap.Contains($opNum)) { $opMap[$opNum] } elseif ($opNum) { $opNum } else { $null }
        $f1 = $null; $f2 = $null
        try { $f1 = $cond.Formula1 } catch {}
        try { $f2 = $cond.Formula2 } catch {}
        $f1Key = Convert-ToR1C1 -App $excel -Formula $f1 -Cell $cell
        $f2Key = Convert-ToR1C1 -App $excel -Formula $f2 -Cell $cell
        $priority = $null
        try { $priority = $cond.Priority } catch {}
        $stopIfTrue = $null
        try { $stopIfTrue = [bool]$cond.StopIfTrue } catch {}

        $key = "$typeName|$opName|$f1Key|$f2Key"
        if (-not $groups.Contains($key)) {
          $groups[$key] = [pscustomobject]@{
            type = $typeName
            operator = $opName
            exampleFormula1 = $f1
            exampleFormula2 = $f2
            priority = $priority
            stopIfTrue = $stopIfTrue
            exampleAddress = $cell.Address($false, $false)
            addresses = New-Object System.Collections.ArrayList
            cellCount = 0
          }
        }
        $groups[$key].cellCount++
        if ($groups[$key].addresses.Count -lt 20) {
          [void]$groups[$key].addresses.Add($cell.Address($false, $false))
        }
      }
    }
  }

  $formatGroups = @()
  $groupCount = 0
  $truncated = $false
  foreach ($key in $groups.Keys) {
    $groupCount++
    if ($groupCount -gt 500) { $truncated = $true; break }
    $g = $groups[$key]
    $formatGroups += [pscustomobject]@{
      type = $g.type
      operator = $g.operator
      exampleFormula1 = $g.exampleFormula1
      exampleFormula2 = $g.exampleFormula2
      priority = $g.priority
      stopIfTrue = $g.stopIfTrue
      exampleAddress = $g.exampleAddress
      addresses = @($g.addresses)
      addressesTruncated = ($g.cellCount -gt 20)
      cellCount = $g.cellCount
    }
  }

  $res = @{ ok=$true; workbook=$wb.Name; sheet=$ws.Name; formatGroups=$formatGroups; totalCells=$totalCells; groupCount=$formatGroups.Count; truncated=$truncated }
  if ($r.LaunchedProcessId) { $res.launchedExcelPid = $r.LaunchedProcessId }
  $res | ConvertTo-Json -Depth 8
} catch {
  @{ ok=$false; error='list_failed'; detail="$($_.Exception.Message)" } | ConvertTo-Json
}
`.trim();

    try {
      const { stdout } = await execFileAsync(
        "powershell.exe",
        ["-NoLogo", "-NoProfile", "-NonInteractive", "-STA", "-ExecutionPolicy", "Bypass", "-Command", psScript],
        { windowsHide: true, encoding: "buffer", timeout: 20000, maxBuffer: 4 * 1024 * 1024 }
      );
      const outText = Buffer.isBuffer(stdout) ? stdout.toString("utf8") : String(stdout);
      return classifyResult(outText);
    } catch (e: any) {
      return extractFailureResult(e);
    }
  }
);

// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ excel_list_data_validations ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
server.tool(
  "excel_list_data_validations",
  "List data validation rules in a worksheet's cells -- input constraints (dropdown list, number/date range, text length) that appear nowhere in VBA code or in the cell's formula. " +
  "For type 'List', exampleFormula1 IS the dropdown's source: either literal choices ('Yes,No,Maybe') or a range reference ('=$D$1:$D$5'). " +
  "Cells are grouped by type, operator and position-normalized Formula1/Formula2, so one rule applied down many rows is a single group with a cellCount. " +
  "type/operator get readable names where known; an unrecognized value comes back as its raw number rather than a guess -- still correct, just unnamed. " +
  "No rules returns an empty validationGroups array, not an error. Does NOT require the VBA Trust Center setting.",
  {
    workbook: z.string().optional().describe("Workbook display name. Give this or workbookPath; workbookPath is preferred (it can auto-launch Excel and open the file)."),
    workbookPath: z.string().optional().describe("Full path to the workbook. Auto-launches Excel and opens the file if needed."),
    sheet: z.string().describe("Worksheet name to scan for data validation rules."),
    range: z.string().optional().describe("Cell range address, e.g. 'A1:D100'. Omit to scan the sheet's entire UsedRange -- prefer a bounded range on very large sheets for speed."),
  },
  async (params) => {
    if (!params.workbook && !params.workbookPath) {
      return { content: [{ type: "text", text: JSON.stringify({ ok: false, error: "workbook or workbookPath is required" }) }], isError: true };
    }
    const wb = psq(params.workbook ?? "");
    const wbPath = psq(params.workbookPath ?? "");
    const sheet = psq(params.sheet);
    const range = psq(params.range ?? "");
    const dotSource = dotSourceExcelUtil();

    const psScript = `
$ErrorActionPreference='Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$OutputEncoding           = [Console]::OutputEncoding

${dotSource}

function Convert-ToR1C1 {
  param($App, $Formula, $Cell)
  if (-not $Formula) { return $null }
  try { return $App.ConvertFormula($Formula, 1, -4150, 4, $Cell) } catch { return $Formula }
}

$typeMap = @{ 0='InputOnly'; 1='WholeNumber'; 2='Decimal'; 3='List'; 4='Date'; 5='Time'; 6='TextLength'; 7='Custom' }
$opMap = @{ 1='Between'; 2='NotBetween'; 3='Equal'; 4='NotEqual'; 5='Greater'; 6='Less'; 7='GreaterEqual'; 8='LessEqual' }

try { $r = Get-OrStartExcelApplication; $excel = $r.App }
catch { @{ ok=$false; error='excel_not_found' } | ConvertTo-Json ; exit }

try { $wb = Resolve-TargetWorkbook -App $excel -WorkbookPath '${wbPath}' -WorkbookName '${wb}' }
catch { @{ ok=$false; error="$($_.Exception.Message)" } | ConvertTo-Json ; exit }

try { $ws = $wb.Worksheets.Item('${sheet}') }
catch { @{ ok=$false; error='sheet_not_found'; sheet='${sheet}' } | ConvertTo-Json ; exit }

$targetRangeAddr = '${range}'
try {
  if ($targetRangeAddr) { $scanRange = $ws.Range($targetRangeAddr) } else { $scanRange = $ws.UsedRange }
} catch {
  @{ ok=$false; error='invalid_range'; range=$targetRangeAddr } | ConvertTo-Json ; exit
}

try {
  $valCells = $null
  try { $valCells = $scanRange.SpecialCells(-4174) } catch { $valCells = $null }

  $groups = [ordered]@{}
  $totalCells = 0
  if ($valCells) {
    foreach ($cell in $valCells) {
      $totalCells++
      $val = $cell.Validation
      $typeNum = $null
      try { $typeNum = $val.Type } catch {}
      $typeName = if ($null -ne $typeNum -and $typeMap.Contains($typeNum)) { $typeMap[$typeNum] } else { $typeNum }
      $opNum = $null
      try { $opNum = $val.Operator } catch {}
      $opName = if ($opNum -and $opMap.Contains($opNum)) { $opMap[$opNum] } elseif ($opNum) { $opNum } else { $null }
      $f1 = $null; $f2 = $null
      try { $f1 = $val.Formula1 } catch {}
      try { $f2 = $val.Formula2 } catch {}
      $f1Key = Convert-ToR1C1 -App $excel -Formula $f1 -Cell $cell
      $f2Key = Convert-ToR1C1 -App $excel -Formula $f2 -Cell $cell
      $inCellDropdown = $null
      try { $inCellDropdown = [bool]$val.InCellDropdown } catch {}
      $inputMessage = $null
      try { $inputMessage = $val.InputMessage } catch {}
      $errorMessage = $null
      try { $errorMessage = $val.ErrorMessage } catch {}

      $key = "$typeName|$opName|$f1Key|$f2Key"
      if (-not $groups.Contains($key)) {
        $groups[$key] = [pscustomobject]@{
          type = $typeName
          operator = $opName
          exampleFormula1 = $f1
          exampleFormula2 = $f2
          inCellDropdown = $inCellDropdown
          inputMessage = $inputMessage
          errorMessage = $errorMessage
          exampleAddress = $cell.Address($false, $false)
          addresses = New-Object System.Collections.ArrayList
          cellCount = 0
        }
      }
      $groups[$key].cellCount++
      if ($groups[$key].addresses.Count -lt 20) {
        [void]$groups[$key].addresses.Add($cell.Address($false, $false))
      }
    }
  }

  $validationGroups = @()
  $groupCount = 0
  $truncated = $false
  foreach ($key in $groups.Keys) {
    $groupCount++
    if ($groupCount -gt 500) { $truncated = $true; break }
    $g = $groups[$key]
    $validationGroups += [pscustomobject]@{
      type = $g.type
      operator = $g.operator
      exampleFormula1 = $g.exampleFormula1
      exampleFormula2 = $g.exampleFormula2
      inCellDropdown = $g.inCellDropdown
      inputMessage = $g.inputMessage
      errorMessage = $g.errorMessage
      exampleAddress = $g.exampleAddress
      addresses = @($g.addresses)
      addressesTruncated = ($g.cellCount -gt 20)
      cellCount = $g.cellCount
    }
  }

  $res = @{ ok=$true; workbook=$wb.Name; sheet=$ws.Name; validationGroups=$validationGroups; totalCells=$totalCells; groupCount=$validationGroups.Count; truncated=$truncated }
  if ($r.LaunchedProcessId) { $res.launchedExcelPid = $r.LaunchedProcessId }
  $res | ConvertTo-Json -Depth 8
} catch {
  @{ ok=$false; error='list_failed'; detail="$($_.Exception.Message)" } | ConvertTo-Json
}
`.trim();

    try {
      const { stdout } = await execFileAsync(
        "powershell.exe",
        ["-NoLogo", "-NoProfile", "-NonInteractive", "-STA", "-ExecutionPolicy", "Bypass", "-Command", psScript],
        { windowsHide: true, encoding: "buffer", timeout: 20000, maxBuffer: 4 * 1024 * 1024 }
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
  "List runnable macros: module-level Public (or implicitly public) Subs. Private/Friend Subs and all Functions are excluded. " +
  "Omit moduleName to cover the whole workbook in one call -- do that rather than calling once per module, since each call re-resolves Excel. " +
  "Each result carries a fully-qualified name to pass straight to excel_run_macro as 'qualified'. " +
  "Scans all open workbooks unless workbookPath narrows it to one file.",
  {
    moduleName: z.string().optional().describe("VBA module name to enumerate procedures in. Omit to list macros from every module in the target workbook in a single call."),
    basPath: z.string().optional().describe("Optional: full path to a previously-exported .bas file for this module; if given, its content hash is used to disambiguate which open workbook to target when multiple books have a same-named module."),
    workbookPath: z.string().optional().describe("Full path to the workbook file. If set, Excel is auto-launched and the file auto-opened when needed, instead of requiring it to already be open."),
  },
  async (params) => {
    const resolved = resolveMacroScript(false);
    if (!resolved.ok) {
      return { content: [{ type: "text", text: JSON.stringify({ error: resolved.error }) }] };
    }
    const ps = resolved.path;

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
  "Run a VBA macro via Application.Run. Get an exact 'qualified' name from excel_list_macros rather than guessing moduleName/procName. " +
  "WARNING: a macro that shows a dialog (MsgBox, InputBox, a modal UserForm) or runs long (infinite loop, runaway recursion) will hang this call until timeoutMs. The timeout only ends this tool's wait -- it does not close the dialog or stop the macro, so Excel stays stuck and every other tool here starts timing out too; use excel_break_execution to interrupt it. " +
  "Success only means Application.Run returned without throwing; it does NOT confirm the macro did what was intended -- verify effects yourself (e.g. excel_read_range).",
  {
    qualified: z.string().optional().describe("Fully-qualified name, e.g. \"'Book1.xlsm'!Module1.DoWork\" (from excel_list_macros). Wins over moduleName/procName."),
    moduleName: z.string().optional().describe("Module name. Needs procName too, if 'qualified' is not given."),
    procName: z.string().optional().describe("Sub name within moduleName. Needs moduleName too, if 'qualified' is not given."),
    workbookName: z.string().optional().describe("Workbook display name, to disambiguate when several open workbooks share the module/proc name."),
    basPath: z.string().optional().describe("Path to a previously-exported .bas; its content hash disambiguates which open workbook to target."),
    workbookPath: z.string().optional().describe("Full path to the workbook. Auto-launches Excel and opens the file if needed."),
    ActivateExcel: z.boolean().optional().describe("Bring the Excel window to the foreground first."),
    ShowStatus: z.boolean().optional().describe("Show a transient message in Excel's status bar."),
    timeoutMs: z.number().optional().describe("Wait before returning ERR_TIMEOUT. Default 30000. Does not unstick Excel itself."),
  },
  async (params) => {
    const resolved = resolveMacroScript(true);
    if (!resolved.ok) {
      return { content: [{ type: "text", text: JSON.stringify({ error: resolved.error }) }] };
    }
    const ps = resolved.path;

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
            detail: `Macro execution timed out after ${timeoutMs}ms. The macro is still running in Excel -- only this tool's wait ended. To stop it: call excel_break_execution (sends Ctrl+Break at the Windows input level, so it works while COM is blocked), or have the user press Ctrl+Break in Excel. Either way the user must then click End in VBA's "Code execution has been interrupted" dialog before any macro will run again. If it is stuck on a dialog (MsgBox/InputBox) instead, the user must dismiss that. Do not retry other tools first -- they queue behind this and time out too.`,
          }) }],
          isError: true,
        };
      }
      return extractFailureResult(e);
    }
  }
);

// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ excel_break_execution ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
// The escape hatch for "an agent wrote a runaway macro, ran it, and now Excel is wedged".
// Two deliberate departures from every other tool here, both forced by that situation:
//   1. It reaches Excel through Windows INPUT (user32 keybd_event), not COM -- a VBA loop
//      makes Excel refuse COM calls, so any COM-based rescue would itself hang.
//   2. It calls execFileAsyncRaw, BYPASSING the excelOpQueue. Queuing it would park the
//      rescue behind the stuck invocation it is meant to rescue -- it must overtake.
// Value beyond unsticking: an interrupted Excel can still SAVE. Force-quitting EXCEL.EXE
// (the only previous option) discards everything written since the last save.
server.tool(
  "excel_break_execution",
  "Interrupt a running or stuck VBA macro (infinite loop, runaway recursion, very long computation) by sending Ctrl+Break to Excel's window at the Windows input level. " +
  "Use this after excel_run_macro returns ERR_TIMEOUT, or when every Excel tool suddenly times out -- it does not use COM, so it still works while Excel is too busy to answer COM calls. " +
  "REQUIRES A HUMAN FOLLOW-UP: a successful break leaves VBA's modal 'Code execution has been interrupted' dialog on screen, and until the user clicks End (not Continue -- that resumes the macro) the project stays in break mode, where every macro run fails with 0x800ADF09 and every module write prompts 'this action will reset the project'. Always relay that instruction; do not retry other tools first. " +
  "Once they click End, Excel is usable again and they can SAVE -- which is the point: force-quitting Excel is the alternative and loses everything since the last save. " +
  "This tool cannot tell you whether the macro actually stopped (read-only COM keeps answering either way), so ask the user what Excel is showing rather than probing. " +
  "Needs an interactive desktop and briefly steals foreground focus; if it reports excel_not_activated the keystroke went to another window and the macro is still running. Cannot interrupt a macro that set Application.EnableCancelKey = xlDisabled, or one blocked inside a Win32/COM call.",
  {
    processId: z.number().optional().describe("PID of the EXCEL.EXE to target. Required only when several Excel processes are running (the tool refuses to guess and returns their PIDs)."),
    alsoSendEsc: z.boolean().optional().describe("Also send Esc ~0.3s after the break. Can dismiss some dialogs; leave off unless Ctrl+Break alone did not free Excel."),
  },
  async ({ processId, alsoSendEsc }) => {
    const scriptsDir = getScriptsDir();
    if (!scriptsDir) {
      return { content: [{ type: "text", text: JSON.stringify({ ok: false, error: "MCP_SCRIPTS_DIR / MCP_PS_LIST not set" }) }], isError: true };
    }
    const ps = path.join(scriptsDir, "Break-ExcelExecution.ps1");
    if (!fs.existsSync(ps)) {
      return { content: [{ type: "text", text: JSON.stringify({ ok: false, error: `ps1 not found: ${ps}` }) }], isError: true };
    }
    const args = [
      "-NoLogo", "-NoProfile", "-NonInteractive", "-STA", "-ExecutionPolicy", "Bypass",
      "-File", ps,
    ];
    if (typeof processId === "number") { args.push("-TargetPid", String(processId)); }
    if (alsoSendEsc) { args.push("-AlsoSendEsc"); }
    try {
      // execFileAsyncRaw, NOT execFileAsync: see the note above -- this must not queue.
      const { stdout } = await execFileAsyncRaw("powershell.exe", args, {
        windowsHide: true,
        encoding: "buffer",
        maxBuffer: 1024 * 1024,
        cwd: path.dirname(ps),
        timeout: 20000,
      } as any);
      const outText = Buffer.isBuffer(stdout) ? stdout.toString("utf8") : String(stdout);
      return classifyResult(outText);
    } catch (e: any) {
      return extractFailureResult(e);
    }
  }
);

// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ excel_export_module ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
// Re-exports one module to the extension's export-folder layout, so the on-disk
// .bas/.cls/.frm a human will edit in VS Code reflects what is actually in Excel.
// Motivation: in the "AI writes via MCP, human polishes in VS Code" workflow, the
// dangerous omission is forgetting to re-export after an MCP write -- the human then
// edits a stale file and a later manual Import silently reverts the AI's change.
// Reuses scripts/export_opened_vba.ps1 (the manual "Export All Modules" path) with
// -BookName/-ModuleName filters and -NoActivate, so Attribute lines, .frm/.frx pairs
// and the Type-100 Lines() fallback behave exactly like a manual export. The exported
// file is written from Excel's CURRENT state via COM .Export(), not from any text this
// server was given -- and its content is deliberately NOT returned (a local file for
// the user's editor must never pass through redaction, and the AI already knows the code).
// Core of excel_export_module, shared with excel_update_module_code's post-write auto-sync
// (MCP_SYNC_MODE=auto): guard check, pre-overwrite backup, script run, sidecar update.
type ModuleExportResult =
  | { ok: true; workbook: string; module: string; exportedPath: string; previousFileBackups?: string[]; note: string }
  | { ok: false; error: string; detail?: string; file?: string; log?: string };

async function performModuleExport(source: string, moduleName: string, exportDir: string, force: boolean): Promise<ModuleExportResult> {
  const scriptsDir = getScriptsDir();
  if (!scriptsDir) {
    return { ok: false, error: "MCP_SCRIPTS_DIR / MCP_PS_LIST not set" };
  }
  const ps = path.join(scriptsDir, "export_opened_vba.ps1");
  if (!fs.existsSync(ps)) {
    return { ok: false, error: `ps1 not found: ${ps}` };
  }
  // The script's -BookName filter compares against the name WITHOUT extension.
  const bookName = path.basename(source, path.extname(source));

  // Guard: refuse to overwrite a file that changed since this tool's last export.
  // The write tool's optimistic lock watches only the VBE side, so a human editing the
  // EXPORTED FILE in VS Code has no protection there -- their saved-but-not-imported
  // edit would be silently replaced. A sidecar records the hash of what this tool last
  // wrote; a mismatch means someone else touched the file since (see exportGuard.ts).
  const bookDirPre = path.join(exportDir, bookName);
  const sidecarPath = path.join(exportDir, ".excel-vba-sync-backups", bookName, `${moduleName}.lastexport.json`);
  const codeFileOnDisk = [".bas", ".cls", ".frm", ".txt"]
    .map((ext) => path.join(bookDirPre, moduleName + ext))
    .find((p) => fs.existsSync(p));
  let lastExportHash: string | null = null;
  try {
    lastExportHash = JSON.parse(fs.readFileSync(sidecarPath, "utf8")).sha256 ?? null;
  } catch { /* no sidecar or unreadable -> provenance unknown, guard stays out of the way */ }
  const currentFileHash = codeFileOnDisk
    ? createHash("sha256").update(fs.readFileSync(codeFileOnDisk)).digest("hex")
    : null;
  const guard = evaluateExportGuard({ currentFileHash, lastExportHash, force });
  if (!guard.ok) {
    return { ok: false, error: guard.error, detail: guard.detail, file: codeFileOnDisk };
  }

  // Back up any existing exported file for this module (and a .frm's paired .frx binary)
  // before the script overwrites it. Two reasons: (1) if the AI-written change is tested
  // and REJECTED, the previous exported content is the fallback; (2) the human may have
  // edited the exported file without importing yet -- this export would silently destroy
  // that work. Same conventions as the write tool's code backup: best-effort (a backup
  // failure never blocks the export), timestamped, under a .excel-vba-sync-backups dir.
  let previousFileBackups: string[] = [];
  try {
    const existing = [".bas", ".cls", ".frm", ".txt", ".frx"]
      .map((ext) => ({ ext, p: path.join(bookDirPre, moduleName + ext) }))
      .filter(({ p }) => fs.existsSync(p));
    if (existing.length > 0) {
      const d = new Date();
      const pad = (n: number) => String(n).padStart(2, "0");
      const ts = `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
      const backupDir = path.join(exportDir, ".excel-vba-sync-backups", bookName);
      fs.mkdirSync(backupDir, { recursive: true });
      for (const { ext, p } of existing) {
        const dest = path.join(backupDir, `${moduleName}_${ts}${ext}`);
        fs.copyFileSync(p, dest);
        previousFileBackups.push(dest);
      }
    }
  } catch {
    previousFileBackups = [];
  }

  const args = [
    "-NoLogo", "-NoProfile", "-NonInteractive", "-STA", "-ExecutionPolicy", "Bypass",
    "-File", ps,
    "-OutputDir", exportDir,
    "-BookName", bookName,
    "-ModuleName", moduleName,
    "-NoActivate",
  ];
  // Exit codes of export_opened_vba.ps1 (localized log text goes to stdout, not JSON).
  const exitCodeMessages: Record<number, string> = {
    1: "exportDir argument missing",
    2: "exportDir is under OneDrive -- refused, same rule as the manual export command. Use a folder outside OneDrive.",
    3: "Excel is not running. This tool does not auto-launch Excel; open the workbook first (any read tool with workbookPath does that).",
    4: "No saved macro-enabled workbook (.xlsm/.xlsb) is open in Excel.",
    5: `exportDir does not exist: ${exportDir}. Create it first (or fix a typo) -- it is not auto-created.`,
    6: `Nothing was exported -- module '${moduleName}' not found in workbook '${bookName}', or its VBA project is protected.`,
  };

  let stdoutText = "";
  try {
    const { stdout } = await execFileAsync("powershell.exe", args, {
      windowsHide: true,
      encoding: "buffer",
      maxBuffer: 2 * 1024 * 1024,
      cwd: path.dirname(ps),
      timeout: 60000,
    });
    stdoutText = Buffer.isBuffer(stdout) ? stdout.toString("utf8") : String(stdout);
  } catch (e: any) {
    const code = typeof e?.code === "number" ? e.code : undefined;
    const tail = (Buffer.isBuffer(e?.stdout) ? e.stdout.toString("utf8") : String(e?.stdout ?? "")).trim().split(/\r?\n/).slice(-5).join("\n");
    return {
      ok: false,
      error: (code !== undefined && exitCodeMessages[code]) || `export script failed (exit ${code ?? "?"})`,
      log: tail,
    };
  }

  // exit 0: locate the produced file (extension depends on the module's type).
  const bookDir = path.join(exportDir, bookName);
  const exported = [".bas", ".cls", ".frm", ".txt"]
    .map((ext) => path.join(bookDir, moduleName + ext))
    .filter((p) => fs.existsSync(p))
    .sort((a, b) => fs.statSync(b).mtimeMs - fs.statSync(a).mtimeMs)[0];
  if (!exported) {
    return {
      ok: false,
      error: `script reported success but no exported file found under ${bookDir}`,
      log: stdoutText.trim().split(/\r?\n/).slice(-5).join("\n"),
    };
  }
  // Record what we just wrote, so the next export can tell "unchanged since my last
  // export" (safe to replace) from "someone edited this in between" (refuse). Best-effort.
  try {
    const newHash = createHash("sha256").update(fs.readFileSync(exported)).digest("hex");
    fs.mkdirSync(path.dirname(sidecarPath), { recursive: true });
    fs.writeFileSync(sidecarPath, JSON.stringify({ file: path.basename(exported), sha256: newHash, at: new Date().toISOString() }));
  } catch { /* a sidecar failure must not fail the export itself */ }

  return {
    ok: true,
    workbook: bookName,
    module: moduleName,
    exportedPath: exported,
    previousFileBackups: previousFileBackups.length > 0 ? previousFileBackups : undefined,
    note: "File reflects Excel's current in-memory VBA project. The workbook itself is still not saved to disk.",
  };
}

server.tool(
  "excel_export_module",
  "Export ONE module from an open workbook to the extension's export-folder layout (<exportDir>\\<workbook name without extension>\\<module>.bas/.cls/.frm), refreshing the on-disk copy a human edits in VS Code. " +
  "Call this right after a confirmed excel_update_module_code write when the user keeps exported files -- otherwise their next manual edit starts from a stale file and a later Import reverts your change. " +
  "An existing exported file (incl. a .frm's paired .frx) is backed up to <exportDir>\\.excel-vba-sync-backups\\ before being overwritten. " +
  "If the file changed since this tool's last export (likely a human's not-yet-imported edit), the call is refused with ERR_EXPORTED_FILE_MODIFIED -- ask the user, then import their file first or re-call with force:true. " +
  "Requires Excel already running with the workbook open (no auto-launch). Returns the exported file path, never the code text.",
  {
    workbook: z.string().optional().describe("Workbook display name (e.g. Book1.xlsm). Give this or workbookPath."),
    workbookPath: z.string().optional().describe("Full path to the workbook; only its base name is used to pick the open workbook."),
    module: z.string().describe("Module name to export."),
    exportDir: z.string().describe("Export ROOT folder (the extension's configured export folder). Must already exist; a subfolder named after the workbook is created inside it. Must not be under OneDrive (refused, same rule as the manual export command)."),
    force: z.boolean().optional().describe("Overwrite even if the exported file was modified since this tool's last export. Only after the user confirmed discarding that change (a backup is still taken)."),
  },
  async ({ workbook, workbookPath, module: moduleName, exportDir, force }) => {
    const source = workbookPath ?? workbook;
    if (!source) {
      return { content: [{ type: "text", text: JSON.stringify({ error: "workbook or workbookPath required" }) }] };
    }
    const result = await performModuleExport(source, moduleName, exportDir, force === true);
    return {
      content: [{ type: "text", text: JSON.stringify(result) }],
      ...(result.ok ? {} : { isError: true as const }),
    };
  }
);

// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ vba_search_code ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
server.tool(
  "vba_search_code",
  "Search VBA source code for a literal string or regex across all currently open workbooks (or one specific workbook via workbookPath / workbookFilter). Returns matching LINES with context (workbook/module/proc/line number), not full module code -- use excel_get_module_code to read a whole module. Results are capped at maxResults (default 50); if there were more matches, the response sets truncated:true and totalMatchCount so you know to narrow the query rather than assuming there were no more hits. Values that look like a hardcoded password/API key/Authorization header within a matched snippet are automatically masked as [REDACTED] -- always on, best-effort only.",
  {
    query: z.string().describe("Search text: plain substring, or a .NET regex if useRegex is true. Case-insensitive."),
    moduleFilter: z.string().optional().describe("Limit to one module name."),
    workbookFilter: z.string().optional().describe("Limit to one open workbook's display name."),
    useRegex: z.boolean().optional().describe("Treat 'query' as a .NET regex instead of a literal substring."),
    workbookPath: z.string().optional().describe("Full path to a workbook to include. Auto-launches Excel and opens the file if needed."),
    maxResults: z.number().optional().describe("Hit cap for one call. Default 50; excess is dropped with truncated:true and totalMatchCount."),
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
      return classifyResultWithRedaction(outText, { arrayField: { field: "hits", subField: "snippet" } });
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
  "Analyze a VBA procedure's control flow (If/ElseIf/Else, Do/Loop, For/Next, Select Case, With, GoTo/labels, Exit/Return, Err.Raise, calls) as structured JSON -- answers 'where does this GoTo jump to' or 'is this branch reachable' from data instead of re-reading raw code. " +
  "Runs against a live snapshot of the module, not the exported files on disk. " +
  "Omit 'procedure' to get a cheap list of {name, kind, startLine, endLine} for every procedure in the module -- do that first if you don't know the exact name. " +
  "Cross-module calls are resolved (all modules are snapshotted), so resolved:false means the callee is genuinely absent from this workbook, not merely unchecked. " +
  "Read-only; writes nothing to disk. Pass workbookPath (full path) unless the workbook is already open. Fails with ERR_VBOM_TRUST_DISABLED if Excel's 'Trust access to the VBA project object model' is off.",
  {
    workbook: z.string().optional().describe("Workbook display name. Give this or workbookPath; workbookPath is preferred (it can auto-launch Excel and open the file)."),
    workbookPath: z.string().optional().describe("Full path to the workbook. Auto-launches Excel and opens the file if needed."),
    module: z.string().describe("VBA module name to analyze."),
    procedure: z.string().optional().describe("Procedure to get full flow detail for. Omit to get a cheap list of every procedure in the module instead."),
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
  "Render a VBA procedure's control flow as Mermaid text (flowchart TD), or the module's call graph if 'procedure' is omitted -- paste it into a Markdown preview or a Mermaid-rendering client to see a diagram instead of reading vba_analyze_flow's raw JSON. " +
  "Runs against a live snapshot of the module, not the exported files on disk. Cross-module calls in the call graph are resolved. " +
  "Read-only; writes nothing to disk (no .mmd files). Pass workbookPath (full path) unless the workbook is already open. Fails with ERR_VBOM_TRUST_DISABLED if Excel's 'Trust access to the VBA project object model' is off.",
  {
    workbook: z.string().optional().describe("Workbook display name. Give this or workbookPath; workbookPath is preferred (it can auto-launch Excel and open the file)."),
    workbookPath: z.string().optional().describe("Full path to the workbook. Auto-launches Excel and opens the file if needed."),
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
  "Scan VBA source for external/platform dependencies: Windows API Declare statements, CreateObject/GetObject, Shell, Application.Run (dynamic dispatch, incl. cross-workbook), native file I/O (Open/Kill/FileCopy/MkDir/RmDir), Scripting.FileSystemObject file/folder methods, and Workbooks.Open. " +
  "Useful for scoping migration work, auditing what a workbook touches outside VBA, and finding files that must travel with it. " +
  "A procedure that vba_analyze_flow shows as uncalled is not necessarily dead -- check here for an Application.Run dispatching to it by name first. " +
  "fileIo entries with methodNameOnly:true were matched by method name only, so the call target's type was not verified. " +
  "Read-only, best-effort regex matching -- not a real VBA parser; it can miss dynamic cases and occasionally match lookalike text inside a string literal. " +
  "Omit 'module' to scan the whole workbook in one COM session. Modules with no findings are omitted. " +
  "Pass workbookPath (full path) unless the workbook is already open. Fails with ERR_VBOM_TRUST_DISABLED if Excel's 'Trust access to the VBA project object model' is off.",
  {
    workbook: z.string().optional().describe("Workbook display name. Give this or workbookPath; workbookPath is preferred (it can auto-launch Excel and open the file)."),
    workbookPath: z.string().optional().describe("Full path to the workbook. Auto-launches Excel and opens the file if needed."),
    module: z.string().optional().describe("Module to scan. Omit to scan the whole workbook."),
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
  "List event-procedure entry points and internal Excel object references, for impact analysis: what triggers this workbook has, and what breaks if a sheet is renamed or a named range removed. " +
  "Events: Workbook_*/Worksheet_*/UserForm_* (matched by VBA's naming convention) plus legacy Auto_Open/Auto_Close. These are triggers nobody calls directly, so a procedure with no incoming calls in vba_analyze_flow is not necessarily dead. Embedded ActiveX control events (e.g. CommandButton1_Click) are NOT detected -- separating those from an ordinary Sub needs control names this tool does not read. " +
  "References: Worksheets(...)/Sheets(...) by name (sheetName null with dynamic:true when computed), plus likely named ranges via Range(...)/Names(...). Range(\"...\") is reported only when the literal does not look like a plain cell address, and dynamic Range(variable) is skipped entirely -- otherwise ordinary cell access would drown out real named-range use. Names(...) is always reported. " +
  "Read-only, best-effort regex matching -- not a real VBA parser. " +
  "Omit 'module' to scan the whole workbook in one COM session. Modules with no findings are omitted. " +
  "Pass workbookPath (full path) unless the workbook is already open. Fails with ERR_VBOM_TRUST_DISABLED if Excel's 'Trust access to the VBA project object model' is off.",
  {
    workbook: z.string().optional().describe("Workbook display name. Give this or workbookPath; workbookPath is preferred (it can auto-launch Excel and open the file)."),
    workbookPath: z.string().optional().describe("Full path to the workbook. Auto-launches Excel and opens the file if needed."),
    module: z.string().optional().describe("Module to scan. Omit to scan the whole workbook."),
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
  "List variable/constant declarations (Dim/Private/Public/Static/Const) classified by scope: 'procedure' (local to one Sub/Function/Property, named by declaredIn), 'module' (module-wide, not visible to other modules), or 'public' (project-wide). " +
  "Use it before a rename: VBE's own Find & Replace has no concept of scope and will happily hit an unrelated same-named local in another procedure. " +
  "Omit variableName to list all declarations in 'module' (or the whole workbook if 'module' is also omitted). " +
  "Give variableName ('module' then required) to instead find that one declaration's usages within its correct boundary -- module/public lookups skip any procedure that shadows the name locally, so unrelated same-named locals are never mixed in. " +
  "If the name matches several declarations, the response is ambiguous_declaration listing each candidate (scope, declaredIn, line); pass 'procedure' matching one candidate's declaredIn to choose. " +
  "Usages are classified 'write' (assignment-shaped) or 'reference' (everything else); the declaration's own line is never counted. " +
  "Read-only, best-effort text matching -- not a real VBA parser. " +
  "Pass workbookPath (full path) unless the workbook is already open, in which case 'workbook' (display name) is enough. Fails with ERR_VBOM_TRUST_DISABLED if Excel's 'Trust access to the VBA project object model' is off.",
  {
    workbook: z.string().optional().describe("Workbook display name. Give this or workbookPath; workbookPath is preferred (it can auto-launch Excel and open the file)."),
    workbookPath: z.string().optional().describe("Full path to the workbook. Auto-launches Excel and opens the file if needed."),
    module: z.string().optional().describe("Module to scan. Omit to scan the whole workbook (declaration-listing mode only; required with variableName)."),
    variableName: z.string().optional().describe("Switches to usage-finding mode for this one name. Requires 'module'."),
    procedure: z.string().optional().describe("Picks one candidate when variableName is ambiguous -- match a declaredIn from the ambiguous_declaration response."),
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

// Advisory only, dry-run only, best-effort: reads every module in the project (one extra
// COM round-trip beyond the dry-run's own read) to flag Public procedure names in newCode
// that already exist in some OTHER module -- catches the case where an agent "moves" a
// procedure to a new/different module but forgets to also remove the original copy. Never
// blocks the dry-run itself: any failure here (e.g. a transient COM hiccup) just yields no
// warnings rather than failing the whole preview.
async function computeDuplicateProcedureWarnings(
  wbEscaped: string,
  wbPathEscaped: string,
  targetModule: string,
  newCode: string
): Promise<CrossModuleDuplicate[]> {
  const allModules = await readAllModulesCode(wbEscaped, wbPathEscaped);
  if (!allModules.ok) { return []; }
  return findCrossModuleDuplicates(targetModule, newCode, allModules.modules);
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
  "Overwrite an EXISTING VBA module's code, or create a new one when moduleType is set. " +
  "REQUIRED two-step flow: call with dryRun:true to preview and get a confirmToken, then call again with identical arguments plus that token to actually write; writing without a valid token is rejected. " +
  "The token is bound to the target's state at dry-run time -- if that changed in between, the write is rejected (ERR_MODULE_CHANGED_SINCE_DRYRUN when overwriting, ERR_MODULE_ALREADY_EXISTS_SINCE_DRYRUN when creating) instead of clobbering it; re-run dryRun for a fresh token. " +
  "UserForms: only the code-behind is replaced. Controls, their layout and the .frx binary are untouched, so adding/moving/renaming a control must be done by a human in the VBE (check excel_list_form_controls before referencing one). Creating a new UserForm is not supported. " +
  "Overwrites write a timestamped backup to '<workbook folder>/.excel-vba-sync-backups' first; creates do not (nothing to back up). " +
  "The dry-run response carries advisory findings that never block the write: lintWarnings (regex-based static checks, not a real VBA parser), duplicateProcedureWarnings (procedure names in newCode that already exist as Public in another module; risk is public_duplicate or private_name_reused), and willLoseShortcutAttributes -- true only for Sheet/ThisWorkbook modules (componentType 100), where per-procedure Attribute lines such as macro shortcut keys cannot survive this write path. " +
  "IMPORTANT: this never saves the workbook to disk. The write is live in the VBA project (visible in the VBE, runnable via excel_run_macro) but is not persisted until the workbook is saved -- tell the user their change is not yet saved.",
  {
    workbook: z.string().optional().describe("Workbook display name. Give this or workbookPath; workbookPath is preferred (it can auto-launch Excel and open the file)."),
    workbookPath: z.string().optional().describe("Full path to the workbook. Auto-launches Excel and opens the file if needed."),
    module: z.string().describe("VBA module name. Must already exist when moduleType is omitted; must NOT exist when moduleType is set."),
    moduleType: z.enum(["standard", "class"]).optional().describe("Set to create a new module instead of overwriting: 'standard' = .bas, 'class' = .cls. Omit to overwrite an existing module."),
    newCode: z.string().describe("Full replacement source code (procedure bodies only -- no Attribute lines)."),
    dryRun: z.boolean().optional().describe("Preview only and return a confirmToken; writes nothing."),
    confirmToken: z.string().optional().describe("Token from a prior dryRun call with identical arguments. Required to actually write."),
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
          newCode: redactCodeText(params.newCode),
          lintWarnings: lintVbaCode(params.newCode),
          confirmToken: createToken,
          note: "Call this tool again with the same workbook/module/moduleType/newCode and this confirmToken to create the module. If a module with this name gets created by someone else before that call, the write will be rejected with ERR_MODULE_ALREADY_EXISTS_SINCE_DRYRUN instead of colliding with it.",
        };
        if (existsResult.launchedExcelPid) { createPreview.launchedExcelPid = existsResult.launchedExcelPid; }
        createPreview.duplicateProcedureWarnings = await computeDuplicateProcedureWarnings(wb, wbPath, params.module, params.newCode);
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
        currentCode: redactCodeText(readResult.currentCode),
        newCode: redactCodeText(params.newCode),
        willLoseShortcutAttributes: readResult.componentType === 100,
        lintWarnings: lintVbaCode(params.newCode),
        confirmToken: expectedToken,
        note: "Call this tool again with the same workbook/module/newCode and this confirmToken to apply the write. If the module's code changes before that call (e.g. another client writes to it first), the token will no longer match and the write will be rejected rather than silently overwriting that change.",
      };
      if (readResult.launchedExcelPid) { preview.launchedExcelPid = readResult.launchedExcelPid; }
      preview.duplicateProcedureWarnings = await computeDuplicateProcedureWarnings(wb, wbPath, params.module, params.newCode);
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

      // Post-write export sync (excelVbaSync.mcpSyncMode, handed over by the extension as
      // MCP_SYNC_MODE / MCP_EXPORT_DIR). The server's instructions are advisory and some
      // clients never show them to the model at all, so the reminder -- or the auto-export
      // result -- rides IN the write response, the one channel every client delivers.
      // This matters most for newly created modules: no exported file exists yet and
      // nothing is open in an editor, so an agent has no other cue to export, nor any
      // way to learn the export root on its own.
      const syncMode = process.env.MCP_SYNC_MODE;
      const syncDir = process.env.MCP_EXPORT_DIR;
      if ((syncMode === "remind" || syncMode === "auto") && syncDir) {
        const jsonStart = outText.indexOf("{");
        if (jsonStart >= 0) {
          try {
            const payload = JSON.parse(outText.slice(jsonStart));
            const source = params.workbookPath || params.workbook || "";
            const bookName = source ? path.basename(source, path.extname(source)) : "";
            // Fire only when this workbook's export folder already exists -- the on-disk
            // evidence that the user actually maintains exported files for it. This is what
            // makes 'remind' safe as the DEFAULT: MCP_EXPORT_DIR always resolves to
            // something (the extension falls back to a default path), so without this
            // check every non-exporting user would start growing a vbaExport folder they
            // never chose. Trade-off accepted: a brand-new workbook that has never been
            // exported stays silent until its first manual/AI export creates the folder.
            if (payload && payload.ok === true && bookName && fs.existsSync(path.join(syncDir, bookName))) {
              if (syncMode === "auto") {
                // Guard applies (force:false): a human's not-yet-imported edit refuses the
                // auto-export, and the refusal is reported instead of being forced over.
                // An export failure must not fail the write -- the write already happened.
                payload.autoExport = await performModuleExport(source, params.module, syncDir, false);
              } else {
                payload.nextAction =
                  `Sync mode 'remind' is on: call excel_export_module for module '${params.module}' with exportDir '${syncDir}' now, ` +
                  `so the exported file matches what was just written.`;
              }
              return { content: [{ type: "text", text: JSON.stringify(payload) }] };
            }
          } catch { /* response was not parseable JSON -> fall through to raw passthrough */ }
        }
      }
      return classifyResult(outText);
    } catch (e: any) {
      return extractFailureResult(e);
    } finally {
      try { fs.unlinkSync(tmpFile); } catch { /* noop */ }
    }
  }
);
console.log("# vba-excel-mcp server: ready");
