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

let messages: Record<string, string> = {};

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
function t(key: string, values?: Record<string, string>) {
  let msg = messages[key] || key;
  if (values) {
    for (const [k, v] of Object.entries(values)) {
      msg = msg.replace(`{${k}}`, v);
    }
  }
  return msg;
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
    } else if (ext === '.frm') {
      this.iconPath = new vscode.ThemeIcon('file-code');
    } else {
      this.iconPath = new vscode.ThemeIcon('file');
    }

    // クリック動作：.frx は開かせない
    this.command = (isFile && ext !== '.frx')
      ? { 
          command: 'vscode.open', 
          title: 'Open File', 
          arguments: [this.resourceUri, { preview: false, preserveFocus: false }] 
        }
      : undefined;

    // 右クリッく用の判定
    this.contextValue = (ext === '.frx') ? 'binaryFrx' : 'vbaModuleFile';

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
      vscode.window.showErrorMessage(t('extension.error.loadFolderFailed', { 0: `${dirPath}` }));
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
        <button onclick="vscode.postMessage({ command: 'export' })">�? Export VBA Modules</button>
        <button onclick="vscode.postMessage({ command: 'import' })">�? Import VBA Modules</button>
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

  context.subscriptions.push(
    vscode.commands.registerCommand('excel-vba-sync.exportVBA', async (fp?: string) => {
      // フォルダの取得
      const folder = (typeof fp === 'string' ? fp : context.globalState.get<string>('vbaExportFolder'));
      if (!folder || typeof folder !== 'string') {
        vscode.window.showErrorMessage(t('extension.error.exportFolderNotConfigured'));
        return;
      }
      /*if (!folder) {
        return vscode.window.showErrorMessage(t('extension.error.exportFolderNotConfigured'));
      }*/

      // スクリプトのパス
      const script = path.join(context.extensionPath, 'scripts', 'export_opened_vba.ps1');
      const cmd = `powershell -NoProfile -ExecutionPolicy Bypass -File "${script}" "${folder}"`;
      await vscode.window.withProgress({
        location: vscode.ProgressLocation.Notification,
        title: t('extension.info.exporting'), cancellable: false
      }, () => new Promise<void>(resolve => {
        cp.exec(cmd, { encoding: 'buffer' }, (err, stdout, stderr) => {
          const out = iconv.decode(stdout as Buffer, 'utf-8').trim();
          const errStr = iconv.decode(stderr as Buffer, 'utf-8').trim();
          vscode.window.createOutputChannel("VBA Export").appendLine(out);
          const exitCode = err?.code;

          // PowerShellからのexit codeに応じてエラー処理
          switch (exitCode) {
            case 1:
              vscode.window.showErrorMessage(t('common.error.noPath'));
              break;
            case 2:
              vscode.window.showErrorMessage(t('common.error.oneDriveFolder'));
              break;
            case 3:
              vscode.window.showErrorMessage(t('common.error.noExcel'));
              break;
            case 4:
              vscode.window.showErrorMessage(t('common.error.noSavedWorkbook'));
              break;
            case 5:
              vscode.window.showErrorMessage(t('common.error.invalidFolder'));
              break;
            case 0:
              vscode.window.showInformationMessage(t('common.info.exportCompleted'));
              treeProvider.refresh();
              break;
            case undefined:
              vscode.window.showInformationMessage(t('common.info.exportCompleted'));
              treeProvider.refresh();
              break;
            default:
              vscode.window.showErrorMessage(t('common.error.exportFailed', { 0: exitCode.toString(), 1: errStr }));
              break;
          }
          resolve();
       });
      }));
    })
  );

  context.subscriptions.push(
    vscode.commands.registerCommand('excel-vba-sync.importVBA', async (item: FileTreeItem) => {
      let filePath: string;
      if (!item || !item.uri) {
        const files = await vscode.window.showOpenDialog({
          canSelectMany: false,
          openLabel: t('extension.info.selectVBA'),
          filters: {
            "VBA Modules": ["bas", "cls", "frm"],
            "All Files": ["*"]
          }
        });
        if (!files || files.length === 0) {
          vscode.window.showErrorMessage(t('extension.error.noFileSelected'));
          return;
        }
        filePath = files[0].fsPath;
      } else {
        filePath = item.uri.fsPath;
      }

      //vscode.window.showWarningMessage("対象ファイルが見つかりません�?");
      //return;
      

      //const filePath = item.uri.fsPath;
      const folder = context.globalState.get<string>('vbaExportFolder');
      if (!folder) {
        vscode.window.showErrorMessage(t('extension.error.importFolderNotConfigured'));
        return;
      }

      const script = path.join(context.extensionPath, 'scripts', 'import_opened_vba.ps1');
      //const cmd = `powershell -ExecutionPolicy Bypass -File "${script}" "${folder}" "${filePath}"`;
      const cmd = `powershell -NoProfile -ExecutionPolicy Bypass -File "${script}" "${filePath}"`;
      vscode.window.showInformationMessage(t('extension.info.targetFolderFiles', { 0: filePath }));
      vscode.window.showInformationMessage(t('extension.info.scriptFile', { 0: script }));

      await vscode.window.withProgress({
        location: vscode.ProgressLocation.Notification,
        title: t('extension.info.importing'),
        cancellable: false
      }, () => new Promise<void>((resolve) => {
        cp.exec(cmd, { encoding: 'buffer' }, (err, stdout, stderr) => {
          const out = iconv.decode(stdout as Buffer, 'utf-8').trim();
          const errStr = iconv.decode(stderr as Buffer, 'utf-8').trim();
          vscode.window.createOutputChannel("VBA Import").appendLine(out);
          const exitCode = err?.code;

          // PowerShellからのexit codeに応じてエラー処理
          switch (exitCode) {
            case 1:
              vscode.window.showErrorMessage(t('common.error.noPath'));
              break;
            case 2:
              vscode.window.showErrorMessage(t('common.error.invalidImportFolder'));
              break;
            case 3:
              vscode.window.showErrorMessage(t('common.error.noExcel'));
              break;
            case 4:
              vscode.window.showErrorMessage(t('common.error.noSavedWorkbook'));
              break;
            case 5:
              vscode.window.showErrorMessage(t('common.error.importFailed'));
              break;
            case 0:
              vscode.window.showInformationMessage(t('common.info.importCompleted'));
              break;
            case undefined:
              vscode.window.showInformationMessage(t('common.info.importCompleted'));
              break;
            default:
              vscode.window.showErrorMessage(t('common.error.importFailed', { 0: exitCode.toString(), 1: errStr }));
              break;
          }
          resolve();
        });
      }));
    })
  );

  context.subscriptions.push(
    vscode.commands.registerCommand('excel-vba-sync.setExportFolder', async () => {
      const result = await vscode.window.showOpenDialog({ canSelectFolders: true });
      if (!result || result.length === 0) {
        return;
      }
      const selected = result[0].fsPath;
      await context.globalState.update('vbaExportFolder', selected);
      vscode.window.showInformationMessage(t('extension.info.exportFolderName', { 0: selected }));
      treeProvider['folderPath'] = selected;
      treeProvider.refresh();
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
