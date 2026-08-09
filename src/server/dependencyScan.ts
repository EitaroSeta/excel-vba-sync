import { redactSecrets } from "./secretRedaction.js";

// Pure, COM-free regex scanning for vba_list_dependencies.
// No Excel/COM access, no PowerShell invocation -- takes module source text already
// read by the caller and returns structured findings. Advisory/best-effort text
// matching only (not a real VBA parser), matching the rigor level already established
// by lintVbaCode's regex-based checks elsewhere in this project: it can miss dynamic
// or commented-out cases, and can rarely false-positive on text that merely resembles
// the pattern inside a string literal.

export interface ApiDeclareEntry {
  module: string;
  line: number;
  scope: "Public" | "Private" | "Friend" | null;
  kind: "Sub" | "Function";
  name: string;
  ptrSafe: boolean;
  lib: string;
  alias: string | null;
  raw: string;
}

export interface ComObjectEntry {
  module: string;
  line: number;
  api: "CreateObject" | "GetObject";
  progId: string | null;
  dynamic: boolean;
  raw: string;
}

export interface ShellCallEntry {
  module: string;
  line: number;
  raw: string;
}

export interface ApplicationRunEntry {
  module: string;
  line: number;
  target: string | null;
  dynamic: boolean;
  raw: string;
}

export interface FileIoEntry {
  module: string;
  line: number;
  operation:
    | "open" | "kill" | "filecopy" | "mkdir" | "rmdir"
    | "fso_opentextfile" | "fso_createtextfile" | "fso_copyfile" | "fso_deletefile"
    | "fso_movefile" | "fso_createfolder" | "fso_deletefolder" | "fso_movefolder";
  // true for the fso_* operations: matched by method name only (e.g. ".CopyFile(")
  // with no way to confirm the object it's called on is actually a
  // Scripting.FileSystemObject instance -- some other object with a
  // same-named method would also match. false for the native VBA statements.
  methodNameOnly: boolean;
  raw: string;
}

export interface ExternalWorkbookEntry {
  module: string;
  line: number;
  target: string | null;
  dynamic: boolean;
  raw: string;
}

export interface ModuleDependencyScan {
  module: string;
  apiDeclares: ApiDeclareEntry[];
  comObjects: ComObjectEntry[];
  shellCalls: ShellCallEntry[];
  applicationRunCalls: ApplicationRunEntry[];
  fileIo: FileIoEntry[];
  externalWorkbooks: ExternalWorkbookEntry[];
}

const RX_DECLARE =
  /^\s*(Public\s+|Private\s+|Friend\s+)?Declare\s+(PtrSafe\s+)?(Sub|Function)\s+([A-Za-z_]\w*)\s+Lib\s+"([^"]+)"(?:\s+Alias\s+"([^"]+)")?/i;

const RX_COM_OBJECT = /\b(CreateObject|GetObject)\s*\(\s*(?:"([^"]*)")?/gi;

// Negative lookbehind excludes '.' (e.g. the literal "WScript.Shell" ProgID string)
// and '"' (Shell mentioned inside an unrelated string) immediately before the word.
const RX_SHELL = /(?<![\w."])\bShell\b\s*[("]/gi;

// Optional captured string is the literal macro/target name; absence means it was
// passed as a variable (dynamic dispatch, can't be resolved by text matching alone).
const RX_APPLICATION_RUN = /\bApplication\.Run\s*\(?\s*(?:"([^"]*)")?/gi;
const RX_WORKBOOKS_OPEN = /\bWorkbooks\.Open\s*\(?\s*(?:"([^"]*)")?/gi;

const FILE_IO_PATTERNS: { operation: FileIoEntry["operation"]; methodNameOnly: boolean; regex: RegExp }[] = [
  { operation: "open", methodNameOnly: false, regex: /\bOpen\s+.+\bFor\s+(?:Input|Output|Append|Random|Binary)\b/i },
  { operation: "kill", methodNameOnly: false, regex: /\bKill\s+\S/i },
  { operation: "filecopy", methodNameOnly: false, regex: /\bFileCopy\s+\S/i },
  { operation: "mkdir", methodNameOnly: false, regex: /\bMkDir\s+\S/i },
  { operation: "rmdir", methodNameOnly: false, regex: /\bRmDir\s+\S/i },
  // FileSystemObject methods -- matched by name only, no type-tracking of the
  // variable they're called on (see FileIoEntry.methodNameOnly). No trailing
  // "(" requirement: Sub-like methods (CopyFile/DeleteFile/MoveFile/*Folder)
  // are commonly called without parens as a statement (`fso.CopyFile "a", "b"`),
  // only Function-like ones (OpenTextFile/CreateTextFile, used with Set x = ...)
  // reliably have them.
  { operation: "fso_opentextfile", methodNameOnly: true, regex: /\.OpenTextFile\b/i },
  { operation: "fso_createtextfile", methodNameOnly: true, regex: /\.CreateTextFile\b/i },
  { operation: "fso_copyfile", methodNameOnly: true, regex: /\.CopyFile\b/i },
  { operation: "fso_deletefile", methodNameOnly: true, regex: /\.DeleteFile\b/i },
  { operation: "fso_movefile", methodNameOnly: true, regex: /\.MoveFile\b/i },
  { operation: "fso_createfolder", methodNameOnly: true, regex: /\.CreateFolder\b/i },
  { operation: "fso_deletefolder", methodNameOnly: true, regex: /\.DeleteFolder\b/i },
  { operation: "fso_movefolder", methodNameOnly: true, regex: /\.MoveFolder\b/i },
];

export function scanModuleForDependencies(moduleName: string, code: string): ModuleDependencyScan {
  const lines = code.split(/\r\n|\r|\n/);
  const apiDeclares: ApiDeclareEntry[] = [];
  const comObjects: ComObjectEntry[] = [];
  const shellCalls: ShellCallEntry[] = [];
  const applicationRunCalls: ApplicationRunEntry[] = [];
  const fileIo: FileIoEntry[] = [];
  const externalWorkbooks: ExternalWorkbookEntry[] = [];

  lines.forEach((line, idx) => {
    const lineNo = idx + 1;

    // Skip whole-line comments (leading ' or Rem) -- otherwise a commented-out
    // example line matches identically to live code. Trailing inline comments
    // (e.g. `Kill "x" ' note`) are not stripped: doing that safely requires
    // knowing where string literals end, which is out of scope for this
    // best-effort regex pass (same limitation as lintVbaCode elsewhere).
    if (/^\s*(?:'|Rem\b)/i.test(line)) { return; }

    const declareMatch = RX_DECLARE.exec(line);
    if (declareMatch) {
      const [, scopeRaw, ptrSafeRaw, kindRaw, name, lib, alias] = declareMatch;
      apiDeclares.push({
        module: moduleName,
        line: lineNo,
        scope: scopeRaw ? (scopeRaw.trim() as "Public" | "Private" | "Friend") : null,
        kind: kindRaw as "Sub" | "Function",
        name,
        ptrSafe: !!ptrSafeRaw,
        lib,
        alias: alias ?? null,
        raw: redactSecrets(line.trim()),
      });
    }

    RX_COM_OBJECT.lastIndex = 0;
    let comMatch: RegExpExecArray | null;
    while ((comMatch = RX_COM_OBJECT.exec(line)) !== null) {
      const [, api, progId] = comMatch;
      comObjects.push({
        module: moduleName,
        line: lineNo,
        api: api as "CreateObject" | "GetObject",
        progId: progId ?? null,
        dynamic: progId === undefined,
        raw: redactSecrets(line.trim()),
      });
    }

    RX_SHELL.lastIndex = 0;
    if (RX_SHELL.test(line)) {
      shellCalls.push({ module: moduleName, line: lineNo, raw: redactSecrets(line.trim()) });
    }

    RX_APPLICATION_RUN.lastIndex = 0;
    let runMatch: RegExpExecArray | null;
    while ((runMatch = RX_APPLICATION_RUN.exec(line)) !== null) {
      const [, target] = runMatch;
      applicationRunCalls.push({
        module: moduleName,
        line: lineNo,
        target: target ?? null,
        dynamic: target === undefined,
        raw: redactSecrets(line.trim()),
      });
    }

    RX_WORKBOOKS_OPEN.lastIndex = 0;
    let wbMatch: RegExpExecArray | null;
    while ((wbMatch = RX_WORKBOOKS_OPEN.exec(line)) !== null) {
      const [, target] = wbMatch;
      externalWorkbooks.push({
        module: moduleName,
        line: lineNo,
        target: target ?? null,
        dynamic: target === undefined,
        raw: redactSecrets(line.trim()),
      });
    }

    for (const p of FILE_IO_PATTERNS) {
      if (p.regex.test(line)) {
        fileIo.push({ module: moduleName, line: lineNo, operation: p.operation, methodNameOnly: p.methodNameOnly, raw: redactSecrets(line.trim()) });
      }
    }
  });

  return { module: moduleName, apiDeclares, comObjects, shellCalls, applicationRunCalls, fileIo, externalWorkbooks };
}
