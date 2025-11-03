/*
    File: extension.ts
    Description: VBAモジュールをVSCodeと同期させる拡張機能
    Author: Eitaro SETA
    License: MIT License
    Copyright (c) 2025 Eitaro Seta

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
*/

// extension.ts
import * as vscode from 'vscode';
import * as path from 'path';
import * as cp from 'child_process';
import * as iconv from 'iconv-lite';
import { readdir } from 'fs/promises';
import * as fs from 'fs';
import * as os from "os";
import { spawn, ChildProcessWithoutNullStreams } from "child_process";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
const execFileAsync = promisify(execFile);

let messages: Record<string, string> = {};
const outputChannel = vscode.window.createOutputChannel("Excel VBA Sync Messages");
let extCtx: vscode.ExtensionContext | undefined;

// MCP server settings
let mcpProc: ChildProcessWithoutNullStreams | null = null;
let channel: vscode.OutputChannel;
let reqId = 0;
const pending = new Map<number, { resolve: (v:any)=>void; reject:(e:any)=>void }>();

// Load localized messages
function loadMessages(context: vscode.ExtensionContext) {
  const locale = vscode.env.language || 'en';
  const filePath = path.join(context.extensionPath, 'locales', `${locale}.json`);

  try {
    messages = JSON.parse(fs.readFileSync(filePath, 'utf-8'));
  } catch {
    messages = JSON.parse(fs.readFileSync(path.join(context.extensionPath, 'locales', 'en.json'), 'utf-8'));
  }

}

// Simple i18n function
function t(key: string, values?: Record<string, string>) {
  let msg = messages[key] || key;
  if (values) {
    for (const [k, v] of Object.entries(values)) {
      msg = msg.replace(`{${k}}`, v);
    }
  }
  return msg;
}

// Get current timestamp
function getTimestamp(): string {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  const hours = String(now.getHours()).padStart(2, '0');
  const minutes = String(now.getMinutes()).padStart(2, '0');
  const seconds = String(now.getSeconds()).padStart(2, '0');
  return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
}

// Watch a folder for changes
function watchFolder(folderPath: string, treeProvider: SimpleTreeProvider) {
  fs.watch(folderPath, { recursive: true }, (eventType, filename) => {
    const timestamp = getTimestamp();
    if (filename) {
      //outputChannel.appendLine(`[${timestamp}] Change detected: ${eventType} - ${filename}`);
      //outputChannel.show();
      treeProvider.refresh(); // Tree view refresh
    } else {
      //outputChannel.appendLine(`[${timestamp}] Change detected, but filename is undefined.`);
      //outputChannel.show();
    }
  });
}

// Search result hit type
type VbaHit = {
  workbook?: string;
  module?: string | { name?: string };
  proc?: string | null;
  line?: number;
  startLine?: number;
  matchLine?: number;
  snippet?: string;
  qualified?: string | null;
  compType?: number;        // 1: bas / 3: frm / その他: cls
  exportExt?: string;       // "bas" | "cls" | "frm"
};

// module field to string
function toModuleName(m: VbaHit["module"]): string {
  if (typeof m === "string") {return m;}
  if (m && typeof m === "object" && typeof (m as any).name === "string") {return (m as any).name;}
  return String(m ?? "");
}

// extension from compType or exportExt
function inferExt(hit: VbaHit): string {
  if (hit.exportExt) {return String(hit.exportExt);}
  switch (hit.compType) { case 1: return "bas"; case 3: return "frm"; default: return "cls"; }
}

// sanitize for directory name
function safeName(s: string): string {
  return s.replace(/[\\/:*?"<>|]/g, "_").trim();
}

// generate candidate directory names for a workbook
// ex: "Book1.xlsm" -> ["Book1.xlsm", "Book1"]
function workbookDirCandidates(workbook: string): string[] {
  const withExt = safeName(workbook);
  const noExt = safeName(path.basename(workbook, path.extname(workbook)));
  // 重複除去
  return [...new Set([withExt, noExt])];
}

// infer export root folder
function resolveExportRoot(): string {
  // 1) globalState（activateで保持した extCtx を使う）
  const globalVal = extCtx?.globalState.get<string>("vbaExportFolder");
  let base = globalVal && globalVal.trim().length ? globalVal : "";
  console.log(`resolveExportRoot: from globalState: ${base}`);
  // 2) 設定
  if (!base) {
    const cfg = vscode.workspace.getConfiguration("excelVbaSync");
    base = cfg.get<string>("vbaExportFolder") ?? "";
  }

  // 3) デフォルト
  if (!base) {
    const ws = vscode.workspace.workspaceFolders?.[0]?.uri.fsPath;
    base = ws ? path.join(ws, "vbaExport") : path.join(os.homedir(), "Excel-VBA-Sync", "vba");
  }

  // ~ と ${workspaceFolder} 展開
  if (base.startsWith("~")) { base = path.join(os.homedir(), base.slice(1)); }
  const ws = vscode.workspace.workspaceFolders?.[0]?.uri.fsPath;
  if (ws) { base = base.replace(/\$\{workspaceFolder\}/g, ws); }

  return base;
}

// exported file search
/// 既存ファイルの探索：<root>/<Workbook候補>/<Module>.<ext> を優先
function findExportedFile(root: string, workbook: string, moduleName: string, ext: string): string | null {
  const wbDirs = workbookDirCandidates(workbook);

  const candidates: string[] = [];
  for (const dir of wbDirs) {
    candidates.push(
      path.join(root, dir, `${moduleName}.${ext}`),            // vba/Book1/Module1.bas  または vba/Book1.xlsm/Module1.bas
      path.join(root, dir, ext, `${moduleName}.${ext}`),       // vba/Book1/bas/Module1.bas など（サブフォルダ運用している場合）
      path.join(root, dir, moduleName, `${moduleName}.${ext}`) // vba/Book1/Module1/Module1.bas（保険）
    );
  }
  // 平置きフォールバック
  candidates.push(path.join(root, `${moduleName}.${ext}`));

  for (const p of candidates) {
    try { if (fs.existsSync(p)) {return p;} } catch {}
  }
  return null;
}

// calculate header offset for exported files
function calcExportHeaderOffset(text: string, ext: string): number {
  const lines = text.split(/\r\n|\n|\r/);
  let i = 0;

  // 1) 先頭の VERSION 行
  while (i < lines.length && /^\s*VERSION\b/i.test(lines[i])) {i++;}

  // 2) .frm のデザイナ領域: Begin ... End ブロックを丸ごと飛ばす（複数あり得る）
  if (ext.toLowerCase() === "frm") {
    let idx = i;
    while (idx < lines.length) {
      if (/^\s*Begin\b/i.test(lines[idx])) {
        // 対応する End までスキップ
        idx++; let depth = 1;
        while (idx < lines.length && depth > 0) {
          if (/^\s*Begin\b/i.test(lines[idx])) {depth++;}
          else if (/^\s*End\b/i.test(lines[idx])) {depth--;}
          idx++;
        }
        i = idx; // 最後の End の次の行
      } else {break;}

    }
  }

  // 3) Attribute VB_* 行（連続していることが多い）
  while (i < lines.length && /^\s*Attribute\b/i.test(lines[i])) {i++;}

  // 4) ここまでが“非表示ヘッダ”。この直後からが CodeModule の先頭とみなす
  return i;
}

// get code from Excel and open in editor, then jump to line 
async function openFromExcelAndJump(hit: VbaHit, moduleName: string) {
  const res2 = await callTool("excel_get_module_code", {
    workbook: hit.workbook ?? "",
    module: moduleName,
  });
  const txt = (res2?.content?.[0]?.text ?? "").toString().trim();

  // JSONだけを安全に抽出
  const start = Math.min(
    ...['{','['].map(ch => {
      const i = txt.indexOf(ch);
      return i === -1 ? Number.POSITIVE_INFINITY : i;
    })
  );
  if (!Number.isFinite(start)) {
    //vscode.window.showWarningMessage(`コード取得に失敗（レスポンス不正）`);
    vscode.window.showWarningMessage(t('extension.error.invalidResponse'));
    return;
  }
  const payload = JSON.parse(txt.slice(start));
  if (!payload?.ok || typeof payload?.code !== "string") {
    //vscode.window.showWarningMessage(`コード取得に失敗: ${payload?.error ?? "unknown"}`);
    vscode.window.showWarningMessage(t('extension.error.codeRetrievalFailed', { error: payload?.error ?? "unknown" }));
    return;
  }

  const doc = await vscode.workspace.openTextDocument({ language: "vb", content: payload.code });
  const ed = await vscode.window.showTextDocument(doc, { preview: false });
  const lineNum = (hit.matchLine ?? hit.startLine ?? hit.line ?? 1);
  const pos = new vscode.Position(Math.max(0, lineNum - 1), 0);
  ed.selection = new vscode.Selection(pos, pos);
  ed.revealRange(new vscode.Range(pos, pos), vscode.TextEditorRevealType.InCenter);
  //vscode.window.setStatusBarMessage(`テンポラリ表示（保存するにはエクスポート先に保存してください）`, 5000);
  vscode.window.setStatusBarMessage(t('extension.statusBarMessage.temporaryDisplay'), 5000);
}

// jump to existing exported file
async function openHit(hit: VbaHit) {
  const moduleName = toModuleName(hit.module);
  const ext = inferExt(hit);

  const root = resolveExportRoot();   // ← context を渡す
  //console.log(`openHit: root=${root}, workbook=${hit.workbook}, module=${moduleName}, ext=${ext}`);
  const existing = (hit.workbook)
    ? findExportedFile(root, hit.workbook, moduleName, ext)
    : null;

  if (existing) {
    const doc = await vscode.workspace.openTextDocument(existing);
    const ed = await vscode.window.showTextDocument(doc, { preview: false });

    // calculate header offset
    const offset = calcExportHeaderOffset(doc.getText(), ext);
    const vbeLine = (hit.matchLine ?? hit.startLine ?? hit.line ?? 1); // ← VBE基準（1始まり）
    const fileLine = Math.max(1, vbeLine + offset);                    // ← ファイル基準に補正

    const pos = new vscode.Position(fileLine - 1, 0);
    ed.selection = new vscode.Selection(pos, pos);
    ed.revealRange(new vscode.Range(pos, pos), vscode.TextEditorRevealType.InCenter);
    return;
  }

  // If not found, show warning
  //vscode.window.showWarningMessage(`未エクスポートのモジュールです。`);
  vscode.window.showWarningMessage(t('extension.warning.unexportedModule'));  
  // If not found, get code from Excel and open in editor
  //await openFromExcelAndJump(hit, moduleName);
}

// ensure MCP server is running
async function ensureServer(context: vscode.ExtensionContext) {
  if (!mcpProc) {
    await startServer(context);
  }
}

// start MCP server
async function startServer(context: vscode.ExtensionContext) {
  if (mcpProc) { vscode.window.showInformationMessage("VBA Tools server already running."); return; }
  if (os.platform() !== "win32") {
    vscode.window.showWarningMessage("VBA Tools features require Windows (Excel + PowerShell).");
    return;
  }

  const cfg = vscode.workspace.getConfiguration("vbaMcp");
  const serverJs = path.join(context.extensionPath, "dist-server", "server.js");
  const exists = await vscode.workspace.fs.stat(vscode.Uri.file(serverJs)).then(()=>true, ()=>false);
  if (!exists) {
    //vscode.window.showErrorMessage(`server.js が見つかりません: ${serverJs}`);
    vscode.window.showErrorMessage(t('extension.error.serverJsNotFound', { path: serverJs }));
    channel.appendLine(`[VBA Tools] NOT FOUND: ${serverJs}`);
    return;
  }
  channel.show(true); // 起動時に Output を前面に
  channel.appendLine(`[VBA Tools] launching: ${serverJs}`);
  const scriptsDir = path.join(context.extensionPath, "scripts");
  const listPs = path.join(scriptsDir, "FindAndRun-ExcelMacroByModule.ps1");
  const runPs  = path.join(scriptsDir, "FindAndRun-ExcelMacroByModule.ps1");

  // ここで存在確認（OutputChannel に出力＆早期 return）
  const okList = await vscode.workspace.fs.stat(vscode.Uri.file(listPs)).then(()=>true, ()=>false);
  if (!okList) {
    channel.show(true);
    channel.appendLine(`[VBA Tools] NOT FOUND: ${listPs}`);
    //vscode.window.showErrorMessage(`.ps1 が見つかりません: ${listPs}（.vscodeignoreで除外していないか確認）`);
    vscode.window.showErrorMessage(t('extension.error.ps1NotFound', { path: listPs }));
    return;
  }

  const env = {
    ...process.env,
    //MCP_VBA_ROOT: cfg.get<string>("vbaRoot") || (vscode.workspace.workspaceFolders?.[0]?.uri.fsPath ?? process.cwd()),
    //MCP_PS_LIST: cfg.get<string>("psListPath") || path.join(scriptsDir, "FindAndRun-ExcelMacroByModule.ps1"),
    //MCP_PS_RUN:  cfg.get<string>("psRunPath")  || path.join(scriptsDir, "FindAndRun-ExcelMacroByModule.ps1")
    MCP_PS_LIST: listPs,   // ★ プロジェクトルートではなく拡張配下を注入
    MCP_PS_RUN:  runPs,
    MCP_VBA_ROOT: vscode.workspace.workspaceFolders?.[0]?.uri.fsPath ?? process.cwd(),
  };

  channel.appendLine(`[MCP] starting: ${serverJs}`);
  mcpProc = spawn(process.execPath, [serverJs], { env });

  mcpProc.stdout.on("data", (d) => {
    const text = d.toString();
    channel.append(text);
    // JSON-RPC のレスポンス取り出し
    for (const line of text.split(/\r?\n/)) {
      try {
        const msg = JSON.parse(line);
        if (msg.id && (msg.result || msg.error)) {
          const h = pending.get(msg.id);
          if (h) {
            pending.delete(msg.id);
            msg.error ? h.reject(msg.error) : h.resolve(msg.result);
          }
        }
      } catch { /* ログ行などは無視 */ }
    }
  });
  mcpProc.stderr.on("data", (d) => channel.append(`[MCP:ERR] ${d}`));
  mcpProc.on("exit", (code) => {
    channel.appendLine(`[MCP] exited: ${code}`);
    mcpProc = null;
  });

  // JSON-RPC handshake（initialize）
  await rpcSend("initialize", {
    protocolVersion: "2024-11-05",   //  SDK version your SDK README.md refers to
    capabilities: {
    // クライアントが使える機能。最小なら空オブジェクトでOK
    tools: {},        // tools/call を使う
    resources: {}     // 使わないなら空でも可
    // （notificationsなど他機能を使う場合はここに宣言を足す）
    },
    clientInfo: {
      name: "vscode-ext",
      version: "0.1.0"
    }
  });
  await rpcSend("tools/list", {}); // ツール一覧取得（存在確認）
  vscode.window.showInformationMessage("MCP server started.");
}

// stop MCP server
function stopServer() {
  if (mcpProc) {
    mcpProc.kill();
    mcpProc = null;
    vscode.window.showInformationMessage("MCP server stopped.");
  }
}

// Lightweight JSON-RPC client (line-delimited)
function rpcSend(method: string, params: any): Promise<any> {
  return new Promise((resolve, reject) => {
    if (!mcpProc) { return reject(new Error("MCP server not running")); }
    const id = ++reqId;
    pending.set(id, { resolve, reject });
    const msg = JSON.stringify({ jsonrpc: "2.0", id, method, params }) + "\n";
    mcpProc.stdin.write(msg, "utf8");
  });
}

// Model Context Protocol tools/call
async function callTool(name: string, args: any): Promise<any> {
  const res = await rpcSend("tools/call", { name, arguments: args });
  // SDK標準のレスポンスは { content: [{type:"text", text:"..."}] } 等
  return res;
}

// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ Command: Search and jump to line ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
async function cmdSearchAndJump() {
  const query = await vscode.window.showInputBox({
    prompt: t('extension.prompt.searchQuery'),
    placeHolder: t('extension.placeholder.searchQuery')
  });
  if (!query) {return;}

  const isRegex = /^\/.*\/$/.test(query);
  const searchQuery  = isRegex ? query.slice(1, -1) : query;

  // parameters are useRegex
  const res = await callTool("vba_search_code", {
    query: searchQuery ,
    useRegex: isRegex,
    // moduleFilter: "...", workbookFilter: "..." // 必要に応じて
  });

  // ---- Only JSON safe parse ----
  const raw = (res?.content?.[0]?.text ?? "").toString().trim();
  const start = Math.min(
    ...['{', '['].map(ch => {
      const i = raw.indexOf(ch);
      return i === -1 ? Number.POSITIVE_INFINITY : i;
    })
  );
  if (!Number.isFinite(start)) {
    vscode.window.showErrorMessage("検索結果の解析に失敗しました。");
    channel.appendLine("[ParseError] " + raw);
    return;
  }
  let payload: any;
  try { payload = JSON.parse(raw.slice(start)); }
  catch (e) {
    vscode.window.showErrorMessage("検索結果の解析に失敗しました。");
    channel.appendLine("[ParseError] " + raw);
    return;
  }

  const hits: any[] = Array.isArray(payload.hits) ? payload.hits : [];
  const count = Number.isFinite(payload.count) ? payload.count : hits.length;
  if (!hits.length) {
    vscode.window.showInformationMessage("ヒットなし");
    return;
  }

  // 1) ヒットの型（受信JSONの形に合わせて必要なら調整）
  type VbaHit = {
    workbook?: string;
    module?: string;
    proc?: string | null;
    line?: number;
    snippet?: string;
    qualified?: string | null;
    compType?: number;   // 1,2,3,100...
    exportExt?: string;  // "bas" | "cls" | "frm"    
  };

  // 2) QuickPick のアイテム型を拡張
  type HitItem = vscode.QuickPickItem & { hit: VbaHit; runnable: boolean; qualified: string | null };

  const items: HitItem[] = hits.map((h: VbaHit) => {
    const wb  = h.workbook ?? "(unknown)";
    const mod = h.module ?? "(unknown)";
    const pr  = h.proc ?? "";
    const line = h.line ?? "?";
    const desc = pr ? `${mod}.${pr}` : `${mod}（実行不可）`;
    const qualified =
      h.qualified
        ? String(h.qualified).replace(/\\u0027/gi, "'")
        : (h.workbook && h.module && h.proc ? `'${h.workbook}'!${h.module}.${h.proc}` : null);

      return {
        label: `${wb}:${line}`,
        description: desc,
        detail: h.snippet ?? "",
        hit: h,                                   // ★ ここで保持
        qualified,
        runnable: !!(qualified && qualified.includes("!") && qualified.includes(".")),
      };
  });

  const picked = await vscode.window.showQuickPick(items, { /* ... */ });
  if (!picked) {return;}
  await openHit(picked.hit);

}

// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ Command: List and run macros ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
async function cmdListAndRunMacro() {
  //const moduleName = await vscode.window.showInputBox({ prompt: "VBA モジュール名（VB_Name）", placeHolder: "Module1" });
  const moduleName = await vscode.window.showInputBox({
    prompt: t('extension.prompt.moduleName'),
    placeHolder: t('extension.placeholder.moduleName')
  });
  if (!moduleName) { return; }

  const active = vscode.window.activeTextEditor?.document;
  const basPath = active && active.fileName.toLowerCase().endsWith(".bas") ? active.fileName : undefined;

  await vscode.window.withProgress(
    { //location: vscode.ProgressLocation.Notification, title: "Excel マクロ一覧を取得中…", cancellable: false },
      location: vscode.ProgressLocation.Notification,
      title: t('extension.progress.fetchingMacroList'),
      cancellable: false
    },
    async (progress) => {
      const listRes = await callTool("excel_list_macros", { moduleName, basPath });
      const listText = (listRes?.content?.[0]?.text ?? "").toString();
      let ary: Array<{ Qualified: string; Proc: string; WorkbookName: string; Module: string }> = [];
      try { ary = JSON.parse(listText); } catch {
        //vscode.window.showErrorMessage("マクロ一覧の解析に失敗しました。Output: VBA Tools を確認してください。");
        vscode.window.showErrorMessage(t('extension.error.macroListParsingFailed'));
        return;
      }
      if (!ary.length) {
        //vscode.window.showWarningMessage("Public Sub が見つかりませんでした。");
        vscode.window.showWarningMessage(t('extension.warning.publicSubNotFound'));
        return;
      }

      type MacroItem = { WorkbookName: string; Module: string; Proc: string; Qualified: string };

      const items = ary.map((x: MacroItem) => ({
        label: x.Proc,
        description: `${x.WorkbookName} / ${x.Module}`,
        detail: x.Qualified,
        macro: x,                        // ← 元データを保持
      }));

      const picked = await vscode.window.showQuickPick(items, {
        //placeHolder: "実行するマクロを選択",
        placeHolder: t('extension.placeholder.selectMacroToRun'),
        matchOnDetail: true,
        ignoreFocusOut: true,
      });
      if (!picked) {return;}

      const m = picked.macro as MacroItem;
      // ★ Qualified を優先して渡す（完全修飾で一意）
      const runArgs: any = {
        qualified: m.Qualified,          // 例：'Book1.xlsm'!Module1.aaa
        // 後方互換のために補助情報も添える（サーバ側で fallback に使える）
        moduleName: m.Module,
        procName: m.Proc,
        workbookName: m.WorkbookName,
        basPath,
        ActivateExcel: true,            // 実行時は前面化
        ShowStatus: true                // 実行時はステータスをON
      };
      const runRes = await callTool("excel_run_macro", runArgs);

      progress.report({ message: `実行中: ${picked.detail}` });
      const runText = (runRes?.content?.[0]?.text ?? "").toString();
      try {
        const payload = JSON.parse(runText);
        if (payload.ok) {
          vscode.window.showInformationMessage(`実行完了: ${payload.ran}`);
        } else {
          vscode.window.showErrorMessage(`実行失敗: ${payload.ran} (${payload.lastError?.error ?? "unknown"})`);
        }
      } catch {
        // JSONでなければ生文字列を通知
        vscode.window.showInformationMessage(runText || `Executed: ${picked.detail}`);
      }
    }
  );
}

/** Tree Item */
class FileTreeItem extends vscode.TreeItem {
  constructor(
    public readonly uri: vscode.Uri,
    public readonly collapsibleState: vscode.TreeItemCollapsibleState
  ) {
    super(uri, collapsibleState);

    const isFile = collapsibleState === vscode.TreeItemCollapsibleState.None;
    const ext = path.extname(uri.fsPath).toLowerCase();

    this.label = path.basename(uri.fsPath);
    this.tooltip = this.label;
    this.resourceUri = uri;

    // アイコン：フォルダ / 通常ファイル / FRX(バイナリ)
    if (!isFile) {
      this.iconPath = new vscode.ThemeIcon('folder');
    } else if (ext === '.frx') {
      this.iconPath = new vscode.ThemeIcon('file-binary'); // or 'lock'
    } else if (ext === '.bas') {
      this.iconPath = new vscode.ThemeIcon('file-code');
    } else if (ext === '.cls') {
      this.iconPath = new vscode.ThemeIcon('file-code');
    } else if (ext === '.frm') {
      this.iconPath = new vscode.ThemeIcon('file-code');
    } else {
      this.iconPath = new vscode.ThemeIcon('file');
    }

    // クリック動作：.frx は開かないようにする
    this.command = (isFile && ext !== '.frx')
      ? { 
          command: 'vscode.open', 
          title: 'Open File', 
          arguments: [this.resourceUri, { preview: false, preserveFocus: false }] 
        }
      : undefined;

    // 右クリック用の判定
    if (ext === '.frx') {
      this.contextValue = 'binaryFrx';
    } else if (['.bas', '.cls', '.frm'].includes(ext)) {
      this.contextValue = 'importableFile'; // インポート可能なファイル
    } else {
      this.contextValue = 'unknownFile'; // その他のファイル
    }
  }
}

/** Tree Provider */
class SimpleTreeProvider implements vscode.TreeDataProvider<FileTreeItem> {
  private _onDidChangeTreeData = new vscode.EventEmitter<FileTreeItem | undefined | void>();
  readonly onDidChangeTreeData = this._onDidChangeTreeData.event;

  constructor(private folderPath: string | undefined) {}

  refresh(): void {
    this._onDidChangeTreeData.fire();
  }

  getTreeItem(element: FileTreeItem): vscode.TreeItem {
    return element;
  }

  async getChildren(element?: FileTreeItem): Promise<FileTreeItem[]> {
    const dirPath = element ? element.uri.fsPath : this.folderPath;
    const timestamp = getTimestamp();

    if (!dirPath) {
      return [];
    }

    try {
      const entries = await readdir(dirPath, { withFileTypes: true });
      return entries.map(entry => {
        const fullPath = path.join(dirPath, entry.name);
        const uri = vscode.Uri.file(fullPath);
        const collapsibleState = entry.isDirectory()
          ? vscode.TreeItemCollapsibleState.Collapsed
          : vscode.TreeItemCollapsibleState.None;
        return new FileTreeItem(uri, collapsibleState);
      });
    } catch (err) {
      //vscode.window.showErrorMessage(t('extension.error.loadFolderFailed', { 0: `${dirPath}` }));
      outputChannel.appendLine(`[${timestamp}] ${t('extension.error.loadFolderFailed', { 0: `${dirPath}` })}`);
      outputChannel.show();
      console.error(err);
      return [];
    }
  }
}

/** Webview Provider */
class VbaSyncViewProvider implements vscode.WebviewViewProvider {
  resolveWebviewView(
    webviewView: vscode.WebviewView,
    _context: vscode.WebviewViewResolveContext,
    _token: vscode.CancellationToken
  ) {
    webviewView.webview.options = { enableScripts: true };
    webviewView.webview.html = `
      <!DOCTYPE html>
      <html lang="ja">
      <head><meta charset="UTF-8"><style>
        body { font-family: sans-serif; padding: 10px; }
        button { width: 100%; margin: 5px 0; font-size: 14px; }
      </style></head>
      <body>
        <button onclick="vscode.postMessage({ command: 'export' })"> Export VBA Modules</button>
        <button onclick="vscode.postMessage({ command: 'import' })"> Import VBA Modules</button>
        <script>const vscode = acquireVsCodeApi();</script>
      </body>
      </html>
    `;

    webviewView.webview.onDidReceiveMessage(msg => {
      if (msg.command === 'export') {
        vscode.commands.executeCommand('excel-vba-sync.exportVBA');
      } else if (msg.command === 'import') {
        vscode.commands.executeCommand('excel-vba-sync.importVBA');
      }
    });
  }
}

/** activate */
export function activate(context: vscode.ExtensionContext) {
  loadMessages(context);
  const folderPath = context.globalState.get<string>('vbaExportFolder');
  const treeProvider = new SimpleTreeProvider(folderPath);
  const treeView = vscode.window.createTreeView('exportPanel', { treeDataProvider: treeProvider });
  //for MCP
  channel = vscode.window.createOutputChannel("VBA Tools");
  context.subscriptions.push(channel);

  context.subscriptions.push(
    vscode.commands.registerCommand("excelVbaSync.searchAndJump", async () => {
      await cmdSearchAndJump();         // ← 引数で渡さない
    })
  );

  // VBAフローチャート生成コマンド
  context.subscriptions.push(
    vscode.commands.registerCommand('excel-vba-sync.generateFlowChart', async (uri?: vscode.Uri) => {
      let vbaFilePath: string;
      const timestamp = getTimestamp();
      
      if (uri && uri.fsPath && uri.fsPath.length > 0) {
        // 右クリックから呼び出された場合
        vbaFilePath = uri.fsPath;
      } else {
        // コマンドパレットから呼び出された場合、または現在開いているファイルを使用
        const activeEditor = vscode.window.activeTextEditor;
        if (activeEditor && activeEditor.document.uri.scheme === 'file') {
          const activeFile = activeEditor.document.uri.fsPath;
          const ext = path.extname(activeFile).toLowerCase();
          if (['.bas', '.cls', '.frm'].includes(ext)) {
            vbaFilePath = activeFile;
          } else {
            const fileUri = await vscode.window.showOpenDialog({
              filters: {
                'VBA Files': ['bas', 'cls', 'frm']
              },
              canSelectFiles: true,
              canSelectMany: false,
              title: 'VBAファイルを選択してフローチャートを生成'
            });

            if (!fileUri || fileUri.length === 0) {
              return;
            }
            
            vbaFilePath = fileUri[0].fsPath;
          }
        } else {
          const fileUri = await vscode.window.showOpenDialog({
            filters: {
              'VBA Files': ['bas', 'cls', 'frm']
            },
            canSelectFiles: true,
            canSelectMany: false,
            title: 'VBAファイルを選択してフローチャートを生成'
          });

          if (!fileUri || fileUri.length === 0) {
            return;
          }
          
          vbaFilePath = fileUri[0].fsPath;
        }
      }

      const folderPath = path.dirname(vbaFilePath);
      const baseName = path.basename(vbaFilePath, path.extname(vbaFilePath));
      
      // mmdフォルダを作成
      const mmdFolderPath = path.join(folderPath, 'mmd');
      if (!fs.existsSync(mmdFolderPath)) {
        fs.mkdirSync(mmdFolderPath, { recursive: true });
      }

      // ステップ1: VBA → JSON
      const vbaScript = path.join(context.extensionPath, 'scripts', 'VBA-FlowJson.ps1');
      const vbaCmd = `powershell -NoLogo -NoProfile -ExecutionPolicy Bypass `
        + `-Command "& { `
        + `$OutputEncoding=[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false); `
        + `& '${vbaScript}' -FolderPath '${folderPath}' -FilePath '${vbaFilePath}' -OutputFolder '${mmdFolderPath}' ;exit $LASTEXITCODE; `
        + `}"`;

      await vscode.window.withProgress({
        location: vscode.ProgressLocation.Notification,
        title: t('extension.flowchart.generating'),
        cancellable: false
      }, () => new Promise<void>(resolve => {
        outputChannel.appendLine(" > > > > > > > > > > > > > > > > > > > >");
        outputChannel.appendLine(`[${timestamp}] ${t('extension.flowchart.generating')}`);
        outputChannel.show();
        
        cp.exec(vbaCmd, { encoding: 'buffer' }, (err, stdout, stderr) => {
          const out = iconv.decode(stdout as Buffer, 'utf-8').trim();
          const errStr = iconv.decode(stderr as Buffer, 'utf-8').trim();
          outputChannel.append(out.endsWith('\n') ? out : out + '\n');
          
          if (errStr && errStr.trim().length > 0) {
            outputChannel.appendLine(`[${getTimestamp()}] STDERR: ${errStr.trim()}`);
          }
          
          const exitCode = err?.code;
          const timestamp = getTimestamp();
          
          if (exitCode === 0 || exitCode === undefined) {
            // JSON生成成功、次にMermaid生成
            const jsonPath = path.join(mmdFolderPath, `${baseName}.flow.json`);
            const mermaidScript = path.join(context.extensionPath, 'scripts', 'Convert-FlowJsonToMermaid.ps1');
            const mermaidCmd = `powershell -NoLogo -NoProfile -ExecutionPolicy Bypass `
              + `-Command "& { `
              + `$OutputEncoding=[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false); `
              + `& '${mermaidScript}' '${jsonPath}' -OutDir '${mmdFolderPath}' ;exit $LASTEXITCODE; `
              + `}"`;
              
            cp.exec(mermaidCmd, { encoding: 'buffer' }, (err2, stdout2, stderr2) => {
              const out2 = iconv.decode(stdout2 as Buffer, 'utf-8').trim();
              const errStr2 = iconv.decode(stderr2 as Buffer, 'utf-8').trim();
              outputChannel.append(out2.endsWith('\n') ? out2 : out2 + '\n');
              
              if (errStr2 && errStr2.trim().length > 0) {
                outputChannel.appendLine(`[${getTimestamp()}] STDERR: ${errStr2.trim()}`);
              }
              
              const exitCode2 = err2?.code;
              const timestamp2 = getTimestamp();
              
              if (exitCode2 === 0 || exitCode2 === undefined) {
                outputChannel.appendLine(`[${timestamp2}] ${t('extension.flowchart.completed', { 0: baseName })}`);
                const filesCreatedMsg = t('extension.flowchart.filesCreated', { 0: baseName }).replace('{0}', baseName);
                outputChannel.appendLine(`[${timestamp2}] ${filesCreatedMsg}`);
                outputChannel.show();
              } else {
                outputChannel.appendLine(`[${timestamp2}] ${t('extension.flowchart.mermaidError', { 0: errStr2 })}`);
                outputChannel.show();
              }
              resolve();
            });
          } else {
            outputChannel.appendLine(`[${timestamp}] ${t('extension.flowchart.jsonError', { 0: errStr })}`);
            outputChannel.show();
            resolve();
          }
        });
      }));
    })
  );
  extCtx = context;

  // Watch the folder for changes
  if (folderPath) {
    watchFolder(folderPath, treeProvider);
  }

  const timestamp = getTimestamp();

  // ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ Export VBA ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
  context.subscriptions.push(
    vscode.commands.registerCommand('excel-vba-sync.exportVBA', async (fp?: string) => {
      // Confirm export folder
      const folder = (typeof fp === 'string' ? fp : context.globalState.get<string>('vbaExportFolder'));
      if (!folder || typeof folder !== 'string') {
        //vscode.window.showErrorMessage(t('extension.error.exportFolderNotConfigured'));
        outputChannel.appendLine(`[${timestamp}] ${t('extension.error.exportFolderNotConfigured')}`);
        outputChannel.show();
        return;
      }
      /*if (!folder) {
        return vscode.window.showErrorMessage(t('extension.error.exportFolderNotConfigured'));
      }*/

      // Get script path
      const script = path.join(context.extensionPath, 'scripts', 'export_opened_vba.ps1');
      //const cmd = `powershell -NoProfile -ExecutionPolicy Bypass -File "${script}" "${folder}"`;
      const cmd = `powershell -NoLogo -NoProfile -ExecutionPolicy Bypass `
        + `-Command "& { `
        + `$OutputEncoding=[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false); `
        + `& '${script}' '${folder}' ;exit $LASTEXITCODE; `
        + `}"`;

      await vscode.window.withProgress({
        location: vscode.ProgressLocation.Notification,
        title: t('extension.info.exporting'), cancellable: false
      }, () => new Promise<void>(resolve => {
        outputChannel.appendLine(" > > > > > > > > > > > > > > > > > > > >");
        outputChannel.appendLine(`[${timestamp}] ${t('extension.info.exporting')}`);
        outputChannel.show();
        cp.exec(cmd, { encoding: 'buffer' }, (err, stdout, stderr) => {
          const out = iconv.decode(stdout as Buffer, 'utf-8').trim();
          const errStr = iconv.decode(stderr as Buffer, 'utf-8').trim();
          //vscode.window.createOutputChannel("Excel VBA Sync Messages").appendLine(out);
          outputChannel.append(out.endsWith('\n') ? out : out + '\n');
          // stderr も表示
          if (errStr && errStr.trim().length > 0) {
            outputChannel.appendLine(`[${getTimestamp()}] STDERR: ${errStr.trim()}`);
          }
          const exitCode = err?.code;

          //console.log(exitCode);
          const timestamp = getTimestamp();
          // Powershell exit code handling
          switch (exitCode) {
            case 1:
              //vscode.window.showErrorMessage(t('common.error.noPath'));
              outputChannel.appendLine(`[${timestamp}] ${t('common.error.noPath')}`);
              outputChannel.show();
              treeProvider.refresh();
              break;
            case 2:
              //vscode.window.showErrorMessage(t('common.error.oneDriveFolder'));
              outputChannel.appendLine(`[${timestamp}] ${t('common.error.oneDriveFolder')}`);
              outputChannel.show();
              break;
            case 3:
              //vscode.window.showErrorMessage(t('common.error.noExcel'));
              outputChannel.appendLine(`[${timestamp}] ${t('common.error.noExcel')}`);
              outputChannel.show();
              break;
            case 4:
              //vscode.window.showErrorMessage(t('common.error.noSavedWorkbook'));
              outputChannel.appendLine(`[${timestamp}] ${t('common.error.noSavedWorkbook')}`);
              outputChannel.show();
              break;
            case 5:
              //vscode.window.showErrorMessage(t('common.error.invalidFolder'));
              outputChannel.appendLine(`[${timestamp}] ${t('common.error.invalidFolder')}`);
              outputChannel.show();
              treeProvider.refresh();
              break;
            case 6:
              //vscode.window.showErrorMessage(t('common.error.invalidFolder'));
              outputChannel.appendLine(`[${timestamp}] ${t('common.error.exportFailedFinal')}`);
              outputChannel.show();
              treeProvider.refresh();
              break;
            case 0:
              //vscode.window.showInformationMessage(t('common.info.exportCompleted'));
              outputChannel.appendLine(`[${timestamp}] ${t('common.info.exportCompleted')}`);
              outputChannel.show();
              treeProvider.refresh();
              break;
            case undefined:
              //vscode.window.showInformationMessage(t('common.info.exportCompleted'));
              outputChannel.appendLine(`[${timestamp}] ${t('common.info.exportCompleted')}`);
              outputChannel.show(); 
              treeProvider.refresh();
              break;
            default:
              //vscode.window.showErrorMessage(t('common.error.exportFailed', { 0: exitCode.toString(), 1: errStr }));
              outputChannel.appendLine(`[${timestamp}] ${t('common.error.exportFailed', { 0: exitCode.toString(), 1: errStr })}`);
              outputChannel.show();
              treeProvider.refresh();
              break;
          }
          resolve();
       });
      }));
    })
  );

  // ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ Import VBA ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
  context.subscriptions.push(
    vscode.commands.registerCommand('excel-vba-sync.importVBA', async (item: FileTreeItem) => {
      let filePath: string;
      const timestamp = getTimestamp();
      //if (!item || !(item.uri instanceof vscode.Uri)) {
      //  outputChannel.appendLine(`[${timestamp}] ${t('extension.error.noFileSelected')}`);
      //  outputChannel.show();
      //  return;
      //}
      //filePath = item.uri.fsPath;
      if (!item) {
        const result = await vscode.window.showOpenDialog({
          canSelectFolders: true,
          canSelectFiles: false,
          canSelectMany: false,
          openLabel: 'Select Folder'
        });

        if (!result || result.length === 0) {
          outputChannel.appendLine(`[${timestamp}] ${t('extension.error.noFileSelected')}`);
          outputChannel.show();
          return;
        }
        filePath = result[0].fsPath;
      } else {
          filePath = item.uri.fsPath;
      }

      if (!filePath || filePath.length === 0) {
         //vscode.window.showWarningMessage(t('extension.error.noFileSelected'));
         outputChannel.appendLine(`[${timestamp}] ${t('extension.error.noFileSelected')}`);
         outputChannel.show();
         return;
      }

      //const folder = context.globalState.get<string>('vbaExportFolder');
      //if (!folder) {
      //if (!folder) {
        //vscode.window.showErrorMessage(t('extension.error.importFolderNotConfigured'));
      //  outputChannel.appendLine(`[${timestamp}] ${t('extension.error.importFolderNotConfigured')}`);
      //  outputChannel.show();
      //  return;
      //}

      // directory check and file type check
      const isDirectory = fs.statSync(filePath).isDirectory();
      const ext = path.extname(filePath).toLowerCase();
      if (!isDirectory && !['.bas', '.cls', '.frm'].includes(ext)) {
        outputChannel.appendLine(`[${timestamp}] ${t('extension.error.invalidFileType', { 0: ext })}`);
        outputChannel.show();
        return;
      }

      const script = path.join(context.extensionPath, 'scripts', 'import_opened_vba.ps1');
      //const cmd = `powershell -NoProfile -ExecutionPolicy Bypass -File "${script}" "${filePath}"`;
      const cmd = `powershell -NoLogo -NoProfile -ExecutionPolicy Bypass `
        + `-Command "& { `
        + `$OutputEncoding=[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false); `
        + `& '${script}' '${filePath}' ;exit $LASTEXITCODE;`
        + `}"`;
      //vscode.window.showInformationMessage(t('extension.info.targetFolderFiles', { 0: filePath }));
      outputChannel.appendLine(" > > > > > > > > > > > > > > > > > > > >");
      outputChannel.appendLine(`[${timestamp}] ${t('extension.info.targetFolderFiles', { 0: filePath })}`);
      outputChannel.show();

      //vscode.window.showInformationMessage(t('extension.info.scriptFile', { 0: script }));
      outputChannel.appendLine(`[${timestamp}] ${t('extension.info.scriptFile', { 0: script })}`);
      outputChannel.show();

      await vscode.window.withProgress({
        location: vscode.ProgressLocation.Notification,
        title: t('extension.info.importing'),
        cancellable: false
      }, () => new Promise<void>((resolve) => {
        outputChannel.appendLine(`[${timestamp}] ${t('extension.info.importing')}`);
        outputChannel.show();
        cp.exec(cmd, { encoding: 'buffer' }, (err, stdout, stderr) => {
          const out = iconv.decode(stdout as Buffer, 'utf-8').trim();
          const errStr = iconv.decode(stderr as Buffer, 'utf-8').trim();
          //vscode.window.createOutputChannel("PS Msg(VBA Import)").appendLine(out);
          outputChannel.append(out.endsWith('\n') ? out : out + '\n');
          // stderr も表示
          if (errStr && errStr.trim().length > 0) {
            outputChannel.appendLine(`[${getTimestamp()}] STDERR: ${errStr.trim()}`);
          }
          const exitCode = err?.code;

          //console.log(exitCode);
          const timestamp = getTimestamp();
          // PowerShellからのexit codeに応じてエラー処理
          switch (exitCode) {
            case 1:
              //vscode.window.showErrorMessage(t('common.error.noPath'));
              outputChannel.appendLine(`[${timestamp}] ${t('common.error.noPath')}`);
              outputChannel.show();
              break;
            case 2:
              //vscode.window.showErrorMessage(t('common.error.invalidImportFolder'));
              outputChannel.appendLine(`[${timestamp}] ${t('common.error.invalidImportFolder')}`);
              outputChannel.show();
              treeProvider.refresh();
              break;
            case 3:
              //vscode.window.showErrorMessage(t('common.error.noExcel'));
              outputChannel.appendLine(`[${timestamp}] ${t('common.error.noExcel')}`);
              outputChannel.show();
              break;
            case 4:
              //vscode.window.showErrorMessage(t('common.error.noSavedWorkbook'));
              outputChannel.appendLine(`[${timestamp}] ${t('common.error.noSavedWorkbook')}`);
              outputChannel.show();
              break;
            case 5:
              //vscode.window.showErrorMessage(t('common.error.importFailed', { 0: exitCode.toString(), 1: errStr }));
              outputChannel.appendLine(`[${timestamp}] ${t('common.error.importFailed')}`);
              outputChannel.show();
              treeProvider.refresh();
              break;
            case 0:
              //vscode.window.showInformationMessage(t('common.info.importCompleted'));
              outputChannel.appendLine(`[${timestamp}] ${t('common.info.importCompleted')}`);
              outputChannel.show();
              break;
            case undefined:
              //vscode.window.showInformationMessage(t('common.info.importCompleted'));
              outputChannel.appendLine(`[${timestamp}] ${t('common.info.importCompleted')}`);
              outputChannel.show();
              break;
            default:
              //vscode.window.showErrorMessage(t('common.error.importFailed', { 0: exitCode.toString(), 1: errStr }));
              outputChannel.appendLine(`[${timestamp}] ${t('common.error.importFailed', { 0: exitCode.toString(), 1: errStr })}`);
              outputChannel.show();
              treeProvider.refresh();
              break;
          }
          resolve();
        });
      }));
    })
  );

  // ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ Set Export Folder ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
  context.subscriptions.push(outputChannel); // 拡張機能終了時にチャネルを破棄
  context.subscriptions.push(
    vscode.commands.registerCommand('excel-vba-sync.setExportFolder', async () => {
      const result = await vscode.window.showOpenDialog({ canSelectFolders: true });
      if (!result || result.length === 0) {
        return;
      }
      const selected = result[0].fsPath;
      await context.globalState.update('vbaExportFolder', selected);
      //vscode.window.showInformationMessage(t('extension.info.exportFolderName', { 0: selected }));
      outputChannel.appendLine(`[${timestamp}] ${t('extension.info.exportFolderName', { 0: selected })}`);
      outputChannel.show();
      treeProvider['folderPath'] = selected;
      treeProvider.refresh();
    })
  );

  // ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ Export Module by Name ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
  context.subscriptions.push(
    vscode.commands.registerCommand('excel-vba-sync.exportModuleByName', async (item: FileTreeItem) => {
      const timestamp = getTimestamp();

      if (!item || !(item.uri instanceof vscode.Uri)) {
        outputChannel.appendLine(`[${timestamp}] ${t('extension.error.noFileSelected')}`);
        outputChannel.show();
        return;
      }

      const fileName = path.basename(item.uri.fsPath, path.extname(item.uri.fsPath)); // ファイル名（拡張子なし）
      const exportFolder = context.globalState.get<string>('vbaExportFolder');

      if (!exportFolder) {
        outputChannel.appendLine(`[${timestamp}] ${t('extension.error.exportFolderNotConfigured')}`);
        outputChannel.show();
        return;
      }

      const bookName = path.basename(path.dirname(item.uri.fsPath)); // フォルダ名（ブック名）
      let moduleName = path.basename(item.uri.fsPath, path.extname(item.uri.fsPath)); // モジュール名（拡張子なし）
      if(!moduleName){
        let moduleName = await vscode.window.showInputBox({
          prompt: "VBA モジュール名（VB_Name）", placeHolder: "Module1",
          validateInput: (v)=> v.trim().length === 0 ? "モジュール名を入力してください" :
                      /\s/.test(v) ? "空白は使用できません" : null
          });
      }
      if (!moduleName) {
        vscode.window.showInformationMessage("キャンセルしました。");
        return;
      }
      const script = path.join(context.extensionPath, 'scripts', 'export_opened_vba.ps1');
      const cmd = `powershell -NoLogo -NoProfile -ExecutionPolicy Bypass `
        + `-Command "& { `
        + `$OutputEncoding=[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false); `
        + `& '${script}' '${exportFolder}' '${bookName}' '${moduleName}' ;exit $LASTEXITCODE; `
        + `}"`;

      outputChannel.appendLine(" > > > > > > > > > > > > > > > > > > > >");
      outputChannel.appendLine(`[${timestamp}] ${t('extension.info.exportingModule', { 0: fileName })}`);
      outputChannel.show();

      await vscode.window.withProgress({
        location: vscode.ProgressLocation.Notification,
        title: t('extension.info.exporting'),
        cancellable: false
      }, () => new Promise<void>((resolve) => {
        cp.exec(cmd, { encoding: 'buffer' }, (err, stdout, stderr) => {
          const out = iconv.decode(stdout as Buffer, 'utf-8').trim();
          const errStr = iconv.decode(stderr as Buffer, 'utf-8').trim();
          outputChannel.append(out.endsWith('\n') ? out : out + '\n');
          if (errStr && errStr.trim().length > 0) {
            outputChannel.appendLine(`[${getTimestamp()}] STDERR: ${errStr.trim()}`);
          }
          const exitCode = err?.code;
          //outputChannel.appendLine(`[${getTimestamp()}] Exit Code: ${exitCode}`);

          // PowerShellからのexit codeに応じてエラー処理
          switch (exitCode) {
            case 1:
              //vscode.window.showErrorMessage(t('common.error.noPath'));
              outputChannel.appendLine(`[${timestamp}] ${t('common.error.noPath')}`);
              outputChannel.show();
              treeProvider.refresh();
              break;
            case 2:
              //vscode.window.showErrorMessage(t('common.error.oneDriveFolder'));
              outputChannel.appendLine(`[${timestamp}] ${t('common.error.oneDriveFolder')}`);
              outputChannel.show();
              break;
            case 3:
              //vscode.window.showErrorMessage(t('common.error.noExcel'));
              outputChannel.appendLine(`[${timestamp}] ${t('common.error.noExcel')}`);
              outputChannel.show();
              break;
            case 4:
              //vscode.window.showErrorMessage(t('common.error.noSavedWorkbook'));
              outputChannel.appendLine(`[${timestamp}] ${t('common.error.noSavedWorkbook')}`);
              outputChannel.show();
              break;
            case 5:
              //vscode.window.showErrorMessage(t('common.error.invalidFolder'));
              outputChannel.appendLine(`[${timestamp}] ${t('common.error.invalidFolder')}`);
              outputChannel.show();
              treeProvider.refresh();
              break;
            case 6:
              //vscode.window.showErrorMessage(t('common.error.invalidFolder'));
              outputChannel.appendLine(`[${timestamp}] ${t('common.error.exportFailedFinal')}`);
              outputChannel.show();
              treeProvider.refresh();
              break;
            case 0:
              //vscode.window.showInformationMessage(t('common.info.exportCompleted'));
              outputChannel.appendLine(`[${timestamp}] ${t('common.info.exportCompleted')}`);
              outputChannel.show();
              treeProvider.refresh();
              break;
            case undefined:
              //vscode.window.showInformationMessage(t('common.info.exportCompleted'));
              outputChannel.appendLine(`[${timestamp}] ${t('common.info.exportCompleted')}`);
              outputChannel.show(); 
              treeProvider.refresh();
              break;
            default:
              //vscode.window.showErrorMessage(t('common.error.exportFailed', { 0: exitCode.toString(), 1: errStr }));
              outputChannel.appendLine(`[${timestamp}] ${t('common.error.exportFailed', { 0: exitCode.toString(), 1: errStr })}`);
              outputChannel.show();
              treeProvider.refresh();
              break;
          }
          resolve();
        });
      }));
    })
  );

  context.subscriptions.push(
    vscode.commands.registerCommand("vbaMcp.start", () => startServer(context)),
    vscode.commands.registerCommand("vbaMcp.stop",  () => stopServer()),
    vscode.commands.registerCommand("vbaMcp.searchCode", async () => {
      await ensureServer(context);
      await cmdSearchAndJump();
    }),
    vscode.commands.registerCommand("vbaMcp.listAndRunMacro", async () => {
      await ensureServer(context);
      await cmdListAndRunMacro();
    }),
    { dispose: () => stopServer() }
  );

  // スタートアップ時に表示されるステータスバーアイコンの登録
  const statusExport = vscode.window.createStatusBarItem(vscode.StatusBarAlignment.Left, 100);
  statusExport.text = '$(cloud-download) Export';
  statusExport.command = 'excel-vba-sync.exportVBA';
  statusExport.tooltip = t('extension.info.tooltip_exportVBA');
  statusExport.show();

  const statusImport = vscode.window.createStatusBarItem(vscode.StatusBarAlignment.Left, 99);
  statusImport.text = '$(cloud-upload) Import';
  statusImport.command = 'excel-vba-sync.importVBA';
  statusImport.tooltip = t('extension.info.tooltip_importVBA');
  statusImport.show();

  context.subscriptions.push(statusExport, statusImport, treeView);

  // Webview Panel 登録
  context.subscriptions.push(
    vscode.window.registerWebviewViewProvider('vbaSyncPanel', new VbaSyncViewProvider())
  );
}

//export function deactivate() {}
export function deactivate() {
   stopServer();
   extCtx = undefined;
}
