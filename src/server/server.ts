import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import * as fs from "node:fs";
import * as path from "node:path";
import * as os from "node:os";
import { createHash } from "node:crypto";
const execFileAsync = promisify(execFile);

console.log("# vba-excel-mcp server: booting...");

const server = new McpServer({ name: "vba-excel-mcp", version: "0.1.0" });
server.tool("ping", {}, async () => ({ content: [{ type: "text", text: "pong" }] }));

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

// PowerShell からの JSON 出力を見て、ERR_ プレフィックスのエラーなら isError:true として返す
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
    }
  } catch { /* JSON以外の出力はそのまま返す */ }
  return { content: [{ type: "text", text: outText }] };
}

// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ excel_get_module_code ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
server.tool(
  "excel_get_module_code",
  {
    workbook: z.string(),
    module: z.string(),
    workbookPath: z.string().optional(),
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
  @{ ok=$true; workbook=$wb.Name; module=$vbc.Name; lines=$cm.CountOfLines; code=$code } | ConvertTo-Json -Depth 6
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

// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ excel_list_macros ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
server.tool(
  "excel_list_macros",
  {
    moduleName: z.string(),
    basPath: z.string().optional(),
    workbookPath: z.string().optional(),
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
      "-ModuleName", params.moduleName,
      "-ListOutput","JSON"
    ];
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
  {
    qualified: z.string().optional(),      // 例："'Book1.xlsm'!Module1.aaa"（最優先）
    moduleName: z.string().optional(),     // qualified が無い場合に使用
    procName: z.string().optional(),       // qualified が無い場合に使用
    workbookName: z.string().optional(),   // 同名対策で限定したい場合に使用（.ps1 側で対応していれば）
    basPath: z.string().optional(),        // 内容一致で限定する場合
    workbookPath: z.string().optional(),   // フルパス指定時、未オープンなら自動でExcelを起動・Open
    ActivateExcel: z.boolean().optional(),
    ShowStatus: z.boolean().optional(),
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

    try {
      const { stdout } = await execFileAsync("powershell.exe", args, {
        windowsHide: true ,
        encoding: "buffer",
        maxBuffer: 2 * 1024 * 1024,
        cwd: path.dirname(ps)
    });
      const outText  = Buffer.isBuffer(stdout) ? stdout.toString("utf8") : String(stdout);
      return classifyResult(outText);
    } catch (e: any) {
      return { content: [{ type: "text", text: JSON.stringify({ error: "ps failed", detail: String(e?.message ?? e) }) }] };
    }
  }
);

// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ vba_search_code ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
server.tool(
  "vba_search_code",
  {
    query: z.string(),
    moduleFilter: z.string().optional(),
    workbookFilter: z.string().optional(),
    useRegex: z.boolean().optional(),
    workbookPath: z.string().optional(),
  },
  async (params) => {
    // PowerShellワンライナーで開いている全ブックの全モジュールを走査
    // ・TrustOM 必須（VBAプロジェクトOMへのアクセスを信頼）
    // ・全コンポーネント種別を対象 vbext_ct_StdModule(1), Class(2), Document(100)
    const wbPath = psq(params.workbookPath ?? "");
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
$reRaw=${JSON.stringify(params.query)}
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
@{ ok=$true; query=$reRaw; hits=$hits; count=$hits.Count } | ConvertTo-Json -Depth 6
`;

    try {
      const { stdout } = await execFileAsync(
        "powershell.exe",
        ["-NoLogo","-NoProfile","-NonInteractive","-STA","-ExecutionPolicy","Bypass","-Command", psScript],
        { windowsHide: true, encoding: "buffer", timeout: 20000, maxBuffer: 2*1024*1024 }
      );
      const outText  = Buffer.isBuffer(stdout) ? stdout.toString("utf8") : String(stdout);
      //return { content: [{ type: "text", text: stdout }] };
      return { content: [{ type: "text", text: outText }] };
    } catch (e:any) {
      return { content: [{ type: "text", text: JSON.stringify({ ok:false, error:"ps_failed", detail:String(e?.message ?? e) }) }] };
    }
  }
);

// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ excel_update_module_code ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
// dry-run（プレビュー）→ confirmToken付き呼び出し（実書き込み）の2段階フロー。
// 実際の書き込みは import_single_module.ps1 経由（= import_opened_vba.ps1 の
// VBComponents.Import()ベースのロジックを再利用）で行い、CodeModule.AddFromString()を
// 直接呼ぶことは絶対に行わない（Attribute行を含むコードでコンパイルエラーになるため）。
function computeConfirmToken(workbook: string, module: string, newCode: string): string {
  return createHash("sha256").update(`${workbook}\u0000${module}\u0000${newCode}`).digest("hex").slice(0, 16);
}

server.tool(
  "excel_update_module_code",
  {
    workbook: z.string().optional(),
    workbookPath: z.string().optional(),
    module: z.string(),
    newCode: z.string(),
    dryRun: z.boolean().optional(),
    confirmToken: z.string().optional(),
  },
  async (params) => {
    const wb = psq(params.workbook ?? "");
    const wbPath = psq(params.workbookPath ?? "");
    const mod = psq(params.module);

    if (!params.workbook && !params.workbookPath) {
      return { content: [{ type: "text", text: JSON.stringify({ ok: false, error: "workbook or workbookPath is required" }) }], isError: true };
    }

    const expectedToken = computeConfirmToken(params.workbook ?? params.workbookPath ?? "", params.module, params.newCode);

    // --- dry-run: 現在のコードを読み取り、差分プレビューと確認トークンを返すだけ。書き込みは行わない ---
    if (params.dryRun) {
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

try { $vbc=$wb.VBProject.VBComponents.Item('${mod}') }
catch { @{ ok=$false; error='module_not_found'; module='${mod}' } | ConvertTo-Json ; exit }

try {
  $cm=$vbc.CodeModule
  $code = if ($cm.CountOfLines -gt 0) { $cm.Lines(1, $cm.CountOfLines) } else { "" }
  @{ ok=$true; workbook=$wb.Name; module=$vbc.Name; componentType=$vbc.Type; currentCode=$code } | ConvertTo-Json -Depth 6
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
        if (classified.isError) { return classified; }

        let payload: any = null;
        try {
          const start = Math.min(...['{', '['].map(ch => { const i = outText.indexOf(ch); return i === -1 ? Number.POSITIVE_INFINITY : i; }));
          payload = Number.isFinite(start) ? JSON.parse(outText.slice(start)) : null;
        } catch { /* noop */ }

        if (!payload?.ok) {
          return { content: [{ type: "text", text: outText }], isError: true };
        }

        const preview = {
          ok: true,
          workbook: payload.workbook,
          module: payload.module,
          componentType: payload.componentType,
          currentCode: payload.currentCode,
          newCode: params.newCode,
          willLoseShortcutAttributes: payload.componentType === 100,
          confirmToken: expectedToken,
          note: "Call this tool again with the same workbook/module/newCode and this confirmToken to apply the write.",
        };
        return { content: [{ type: "text", text: JSON.stringify(preview, null, 2) }] };
      } catch (e: any) {
        return { content: [{ type: "text", text: JSON.stringify({ ok: false, error: "ps_failed", detail: String(e?.message ?? e) }) }], isError: true };
      }
    }

    // --- confirmToken必須（dry-runを経ずにいきなり書き込むことを防ぐ） ---
    if (!params.confirmToken) {
      return {
        content: [{ type: "text", text: JSON.stringify({ ok: false, error: "confirmToken is required. Call this tool with dryRun:true first to preview the change and obtain a confirmToken." }) }],
        isError: true,
      };
    }
    if (params.confirmToken !== expectedToken) {
      return {
        content: [{ type: "text", text: JSON.stringify({ ok: false, error: "confirmToken does not match the current (workbook, module, newCode). The code may have changed since the dry-run; re-run with dryRun:true to get a fresh token." }) }],
        isError: true,
      };
    }

    // --- 実書き込み: import_single_module.ps1 経由（Import-ModuleToVBProjectを再利用） ---
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
      const { stdout } = await execFileAsync("powershell.exe", args, {
        windowsHide: true,
        encoding: "buffer",
        timeout: 30000,
        maxBuffer: 2 * 1024 * 1024,
      });
      const outText = Buffer.isBuffer(stdout) ? stdout.toString("utf8") : String(stdout);
      return classifyResult(outText);
    } catch (e: any) {
      return { content: [{ type: "text", text: JSON.stringify({ ok: false, error: "ps_failed", detail: String(e?.message ?? e) }) }], isError: true };
    } finally {
      try { fs.unlinkSync(tmpFile); } catch { /* noop */ }
    }
  }
);

console.log("# vba-excel-mcp server: ready");
