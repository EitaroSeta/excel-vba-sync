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
import { spawn } from 'child_process';

let messages: Record<string, string> = {};
const outputChannel = vscode.window.createOutputChannel("Excel VBA Sync Messages");

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
    //this.contextValue = (ext === '.frx') ? 'binaryFrx' : 'vbaModuleFile';
    //this.contextValue = (ext === '.frx') ? 'binaryFrx' : 'vbaModuleFile';
    //console.log(`Context value: ${this.contextValue}`); // デバッグ用ログ
    //console.log(`File extension: ${ext}`);
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

  // Watch the folder for changes
  if (folderPath) {
    watchFolder(folderPath, treeProvider);
  }

  const timestamp = getTimestamp();

  // Export VBA
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
        outputChannel.appendLine(">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>");
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

  // Import VBA
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
      outputChannel.appendLine(">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>");
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

  // exportModuleByName
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
      const moduleName = path.basename(item.uri.fsPath, path.extname(item.uri.fsPath)); // モジュール名（拡張子なし）
      const script = path.join(context.extensionPath, 'scripts', 'export_opened_vba.ps1');
      const cmd = `powershell -NoLogo -NoProfile -ExecutionPolicy Bypass `
        + `-Command "& { `
        + `$OutputEncoding=[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false); `
        + `& '${script}' '${exportFolder}' '${bookName}' '${moduleName}' ;exit $LASTEXITCODE; `
        + `}"`;

      outputChannel.appendLine(">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>");
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

export function deactivate() {}
