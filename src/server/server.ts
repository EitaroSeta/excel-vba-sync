import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import * as fs from "node:fs";
import * as path from "node:path";
const execFileAsync = promisify(execFile);

console.log("# vba-excel-mcp server: booting...");

const server = new McpServer({ name: "vba-excel-mcp", version: "0.1.0" });
server.tool("ping", {}, async () => ({ content: [{ type: "text", text: "pong" }] }));

const transport = new StdioServerTransport();
server.connect(transport);

// 文字列の ' をエスケープ
function psq(s: string) { return s.replace(/'/g, "''"); }

// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ excel_get_module_code ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
server.tool(
  "excel_get_module_code",
  {
    workbook: z.string(),
    module: z.string(),
  },
  async (params) => {
    const wb = psq(params.workbook);
    const mod = psq(params.module);

    // PowerShell ワンライナーで COM 経由取得
    const psScript = `
$ErrorActionPreference='Stop'
# --- Force UTF-8 output (no BOM) ---
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$OutputEncoding           = [Console]::OutputEncoding

try { $excel=[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') }
catch { @{ ok=$false; error='excel_not_found' } | ConvertTo-Json ; exit }

$wb=$excel.Workbooks | Where-Object { $_.Name -eq '${wb}' }
if(-not $wb){ @{ ok=$false; error='workbook_not_found'; workbook='${wb}' } | ConvertTo-Json ; exit }

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
      //const errText  = Buffer.isBuffer(stderr) ? stderr.toString("utf8") : String(stderr);
      //return { content: [{ type: "text", text: stdout }] };
      return { content: [{ type: "text", text: outText }] };
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

    try {
      const { stdout } = await execFileAsync("powershell.exe", args, { 
        windowsHide: true,
        encoding: "buffer",      // Buffer で受け取ってから UTF-8 に変換
        cwd: path.dirname(ps),   // ps1 のあるフォルダをカレントに
        timeout: 20000,          // ★ 20 秒で強制終了
        maxBuffer: 2 * 1024 * 1024
      });
      const outText  = Buffer.isBuffer(stdout) ? stdout.toString("utf8") : String(stdout);
      //return { content: [{ type: "text", text: stdout }] };
      return { content: [{ type: "text", text: outText }] };
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
      //return { content: [{ type: "text", text: stdout }] };
      return { content: [{ type: "text", text: outText }] };
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
  },
  async (params) => {
    // PowerShellワンライナーで開いている全ブックの全モジュールを走査
    // ・TrustOM 必須（VBAプロジェクトOMへのアクセスを信頼）
    // ・全コンポーネント種別を対象 vbext_ct_StdModule(1), Class(2), Document(100)
    const psScript = `
# --- Force UTF-8 (no BOM) for stdout/stderr ---
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$OutputEncoding           = [Console]::OutputEncoding

$ErrorActionPreference='Stop'
try{
  $excel=[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
}catch{ 
  Write-Output (@{ ok=$false; error='excel_not_found' } | ConvertTo-Json); exit 
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

console.log("# vba-excel-mcp server: ready");
