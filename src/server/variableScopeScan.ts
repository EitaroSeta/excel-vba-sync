// Pure, COM-free regex scanning for vba_list_variable_scopes.
// No Excel/COM access, no PowerShell invocation -- takes module source text already
// read by the caller and returns structured findings. Advisory/best-effort text
// matching only (not a real VBA parser), same rigor tier as dependencyScan.ts /
// referenceScan.ts.
//
// Purpose: VBE's own Find & Replace ("Search In: Current Project") is a blind text
// substitution with no concept of scope -- it will happily replace an unrelated local
// variable in a different procedure just because it shares the same name. This tool
// answers "is this declaration local to one procedure, module-wide, or project-wide"
// so a rename's true impact can be scoped correctly before using VBE's search, or
// vba_search_code (with moduleFilter) to actually find/verify occurrences within the
// reported boundary. It does not itself list usage sites -- only declarations and
// their scope.

export interface VariableDeclaration {
  module: string;
  line: number;
  name: string;
  kind: "variable" | "constant";
  scope: "procedure" | "module" | "public";
  // Procedure name this declaration is local to, when scope === "procedure".
  // Combine with vba_analyze_flow's startLine/endLine for that procedure (or this
  // module's own procedure boundaries, computed the same way here) to bound a
  // vba_search_code query to just that procedure's lines.
  declaredIn: string | null;
  raw: string;
}

export interface ModuleVariableScopeScan {
  module: string;
  declarations: VariableDeclaration[];
}

interface ProcedureBounds {
  name: string;
  startLine: number;
  endLine: number;
}

const RX_PROC_HEAD = /^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)\s*\(/i;
const RX_PROC_END = /^\s*End\s+(?:Sub|Function|Property)\b/i;

const RX_PROC_KEYWORD_AFTER_DECL = /^(Sub|Function|Property)\b/i;

interface ParsedDeclKeyword {
  modifier: "Public" | "Private" | null;
  isConst: boolean;
  rest: string;
}

// VBA's declaration grammar only allows one specific two-keyword combination
// (Public/Private followed by Const); Dim/Static/Const never take a modifier.
// Handled as an explicit small parser rather than one regex, since a single
// regex covering "optional modifier, then optionally Const, else require
// Dim/Static/Const" got unreadable fast -- and a single-keyword regex missed
// `Public Const NAME = ...` entirely (captured "Const" itself as the name).
function parseDeclKeyword(line: string): ParsedDeclKeyword | null {
  let s = line;
  let modifier: "Public" | "Private" | null = null;

  const modMatch = /^\s*(Public|Private)\s+/i.exec(s);
  if (modMatch) {
    modifier = /^public/i.test(modMatch[1]) ? "Public" : "Private";
    s = s.slice(modMatch[0].length);
  }

  const constMatch = /^\s*Const\s+/i.exec(s);
  if (constMatch) {
    return { modifier, isConst: true, rest: s.slice(constMatch[0].length) };
  }

  if (modifier) {
    // "Public x As Long" / "Private x As Long" -- direct variable declaration.
    return { modifier, isConst: false, rest: s };
  }

  const declMatch = /^\s*(Dim|Static)\s+/i.exec(s);
  if (declMatch) {
    return { modifier: null, isConst: false, rest: s.slice(declMatch[0].length) };
  }

  return null;
}

function findProcedureBounds(lines: string[]): ProcedureBounds[] {
  const bounds: ProcedureBounds[] = [];
  let open: { name: string; startLine: number } | null = null;
  lines.forEach((line, idx) => {
    const lineNo = idx + 1;
    if (!open) {
      const m = RX_PROC_HEAD.exec(line);
      if (m) { open = { name: m[2], startLine: lineNo }; }
    } else if (RX_PROC_END.test(line)) {
      bounds.push({ name: open.name, startLine: open.startLine, endLine: lineNo });
      open = null;
    }
  });
  return bounds;
}

function findEnclosingProcedure(lineNo: number, bounds: ProcedureBounds[]): ProcedureBounds | null {
  for (const b of bounds) {
    if (lineNo >= b.startLine && lineNo <= b.endLine) { return b; }
  }
  return null;
}

// Splits on commas that are not nested inside parentheses, so array bounds like
// `Dim arr(1 To 10, 1 To 5) As Variant` don't get split on their internal comma.
function splitTopLevelCommas(s: string): string[] {
  const parts: string[] = [];
  let depth = 0;
  let current = "";
  for (const ch of s) {
    if (ch === "(") { depth++; current += ch; }
    else if (ch === ")") { depth = Math.max(0, depth - 1); current += ch; }
    else if (ch === "," && depth === 0) { parts.push(current); current = ""; }
    else { current += ch; }
  }
  if (current.trim().length > 0) { parts.push(current); }
  return parts;
}

export function scanModuleForVariableScopes(moduleName: string, code: string): ModuleVariableScopeScan {
  const lines = code.split(/\r\n|\r|\n/);
  const bounds = findProcedureBounds(lines);
  const declarations: VariableDeclaration[] = [];

  lines.forEach((line, idx) => {
    const lineNo = idx + 1;

    // Skip whole-line comments -- see dependencyScan.ts for the same check and its
    // rationale (trailing inline comments are not stripped).
    if (/^\s*(?:'|Rem\b)/i.test(line)) { return; }

    const parsed = parseDeclKeyword(line);
    if (!parsed) { return; }

    // `Public Sub Foo()` / `Static Sub Foo()` -- a procedure header, not a
    // variable declaration.
    if (RX_PROC_KEYWORD_AFTER_DECL.test(parsed.rest.trimStart())) { return; }

    const enclosing = findEnclosingProcedure(lineNo, bounds);
    const scope: VariableDeclaration["scope"] =
      enclosing ? "procedure" : (parsed.modifier === "Public" ? "public" : "module");
    const kind: VariableDeclaration["kind"] = parsed.isConst ? "constant" : "variable";

    for (const segment of splitTopLevelCommas(parsed.rest)) {
      const nameMatch = /^\s*([A-Za-z_]\w*)/.exec(segment);
      if (!nameMatch) { continue; }
      declarations.push({
        module: moduleName,
        line: lineNo,
        name: nameMatch[1],
        kind,
        scope,
        declaredIn: enclosing ? enclosing.name : null,
        raw: line.trim(),
      });
    }
  });

  return { module: moduleName, declarations };
}
