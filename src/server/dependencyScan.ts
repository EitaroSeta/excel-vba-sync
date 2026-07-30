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

export interface ModuleDependencyScan {
  module: string;
  apiDeclares: ApiDeclareEntry[];
  comObjects: ComObjectEntry[];
  shellCalls: ShellCallEntry[];
}

const RX_DECLARE =
  /^\s*(Public\s+|Private\s+|Friend\s+)?Declare\s+(PtrSafe\s+)?(Sub|Function)\s+([A-Za-z_]\w*)\s+Lib\s+"([^"]+)"(?:\s+Alias\s+"([^"]+)")?/i;

const RX_COM_OBJECT = /\b(CreateObject|GetObject)\s*\(\s*(?:"([^"]*)")?/gi;

// Negative lookbehind excludes '.' (e.g. the literal "WScript.Shell" ProgID string)
// and '"' (Shell mentioned inside an unrelated string) immediately before the word.
const RX_SHELL = /(?<![\w."])\bShell\b\s*[("]/gi;

export function scanModuleForDependencies(moduleName: string, code: string): ModuleDependencyScan {
  const lines = code.split(/\r\n|\r|\n/);
  const apiDeclares: ApiDeclareEntry[] = [];
  const comObjects: ComObjectEntry[] = [];
  const shellCalls: ShellCallEntry[] = [];

  lines.forEach((line, idx) => {
    const lineNo = idx + 1;

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
        raw: line.trim(),
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
        raw: line.trim(),
      });
    }

    RX_SHELL.lastIndex = 0;
    if (RX_SHELL.test(line)) {
      shellCalls.push({ module: moduleName, line: lineNo, raw: line.trim() });
    }
  });

  return { module: moduleName, apiDeclares, comObjects, shellCalls };
}
