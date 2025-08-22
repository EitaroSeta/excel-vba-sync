> VS CodeでVBAを編集・レビューしやすくする双方向同期拡張（Excelとモジュールを行き来）

### Quickstart (.vsix)
Windows:
```powershell
iwr -useb https://github.com/EitaroSeta/excel-vba-sync/releases/latest/download/vba-module-sync-latest.vsix -OutFile $env:TEMP\vba-module-sync-latest.vsix; code --install-extension $env:TEMP\vba-module-sync-latest.vsix
```

> アップデート: 上と同じコマンドで再実行（`--force` 推奨）。UI手順は下の「インストール（VSIX）」参照

﻿# 📤 VBA Module Sync - VSCode ⇄ Excel

## 概要（Japanese）

**VBA Module Sync** は、Excel の VBA モジュールを VSCode 上で編集するための拡張機能です。  
VBA モジュールのエクスポート・インポートを双方向で行えます。

- ✅ Excelから `.bas` / `.cls` / `.frm` 内のコードをエクスポート（保存）
- ✅ VSCode 上で編集
- ✅ 編集したモジュールを Excel にインポート（反映）
- ✅ インポートはモジュール差し替えにて行います

### 🔧 主な機能

| 機能                           | 説明                                      |
|--------------------------------|-------------------------------------------|
| Export All Modulus From VBA    | Excel から全モジュールを抽出・保存します。|
| Import Module To VBA           | VSCode 上で編集したコードを Excel に反映（単一モジュール単位/ファイル単位） |
| Set Export Folder              | ダイアログ選択可能                        |
| コマンドパレット／ボタン対応   | GUI 操作または `Ctrl+Shift+P` から実行可  |

---

## Overview (English)

**VBA Module Sync** is a VSCode extension for editing Excel VBA modules.  
It enables **bidirectional sync** between Excel and VSCode.

- ✅ Export inner code of  `.bas` / `.cls` / `.frm` from Excel
- ✅ Edit VBA modules in VSCode
- ✅ Import modules back into Excel
- ✅ Import is performed by replacing the module.

### 🔧 Features

| Feature                        | Description                                      |
|--------------------------------|------------------------------------------------------------------|
| Export All Modules From VBA    | Extract and save all VBA modules from Excel                      |
| Import Module To VBA           | Reflect modified code back to Excel(Module-based/File-based)     |
| Set Export Folder              | Change export folder via Dialog                                  |
| Command Palette / GUI support  | Use commands or side panel buttons                               |

---

## ⚙️ ローカライズ設定例 / Localization Example

拡張機能の表示テキストは locales フォルダの言語別 JSON ファイルで管理しています。
現在は以下の2言語に対応していますので、*.jsonを使用したい言語に合わせて作ってください。

The extension’s display text is managed in language-specific JSON files located in the locales folder.
Currently, the following two languages are supported, so please create a *.json file for the language you want to use.

 locales/
  ├─ ja.json
  └─ en.json

---

## 🛠 開発 (Development)

### 前提 / Requirements
- Windows + Microsoft Excel（VBA を実行するため）
- Node.js LTS（18 以上推奨）と npm
- Visual Studio Code（拡張の起動・デバッグに使用）

### セットアップ / Setup
npm install
```

### ビルド & 実行 / Build & Run
npm run compile
```
- VS Code で `F5` を押して **Extension Development Host** を起動

### 主要コマンド / Key Commands
- **Export All Modules From VBA** — Excel から VBA モジュールを一括エクスポート
- **Import Module To VBA** — 編集したモジュールを Excel に取り込み
- **Set Export Folder** — エクスポート先フォルダの指定
> いずれもコマンドパレット（`Ctrl+Shift+P`）から実行できます。

### パッケージ化（任意） / Package (optional)
`vsce` で配布用 `.vsix` を作成できます（CLI）。
npm i -g @vscode/vsce
vsce package
```
`.vscodeignore` により TypeScript やテスト等はパッケージから除外されます。

### リポジトリ構成（抜粋） / Repo Layout
- `src/` — 拡張のソースコード（TypeScript）
- `scripts/` — Excel 連携用 PowerShell スクリプト
- `locales/` — 多言語リソース（`ja.json`, `en.json`）

## 🧩 インストール（VSIX） / Install from VSIX

### 方法A: VS Code の UI から
1. VS Code を開く
2. 拡張機能ビュー（Ctrl+Shift+X / Cmd+Shift+X）を開く
3. 右上の「…」メニュー → **Install from VSIX...** を選択
4. 作成した `.vsix` ファイル（例: `your-extension-0.1.0.vsix`）を選択
5. 再読み込み（Reload）が求められたら実行