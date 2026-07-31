// Pure, COM-free regex scanning for vba_list_references.
// No Excel/COM access, no PowerShell invocation -- takes module source text already
// read by the caller and returns structured findings. Advisory/best-effort text
// matching only (not a real VBA parser), same rigor tier as dependencyScan.ts.

export interface EventProcedureEntry {
  module: string;
  line: number;
  procedure: string;
  eventKind: "workbook_event" | "worksheet_event" | "userform_event" | "auto_macro";
  raw: string;
}

export interface SheetReferenceEntry {
  module: string;
  line: number;
  api: "Worksheets" | "Sheets";
  sheetName: string | null;
  dynamic: boolean;
  raw: string;
}

export interface NamedRangeReferenceEntry {
  module: string;
  line: number;
  source: "Range" | "Names";
  name: string | null;
  dynamic: boolean;
  raw: string;
}

export interface ModuleReferenceScan {
  module: string;
  eventProcedures: EventProcedureEntry[];
  sheetReferences: SheetReferenceEntry[];
  namedRangeReferences: NamedRangeReferenceEntry[];
}

// Workbook_*/Worksheet_*/UserForm_* is VBA's own naming convention for event
// procedures -- matching the prefix catches every event without needing to
// enumerate the (long, version-dependent) list of actual event names.
// Embedded ActiveX control events (e.g. CommandButton1_Click) are NOT covered:
// telling those apart from an ordinary Sub requires knowing the sheet/form's
// control names, which needs the Controls/OLEObjects COM collection (not read
// by this tool -- out of scope for a text-only pass).
const RX_EVENT_PROCEDURE =
  /^\s*(?:Public\s+|Private\s+|Friend\s+)?Sub\s+((Workbook|Worksheet|UserForm)_\w+|Auto_Open|Auto_Close)\s*\(/i;

const RX_SHEET_REF = /\b(Worksheets|Sheets)\s*\(\s*(?:"([^"]*)")?/gi;
// The lookahead requires ')' or ',' right after the closing quote (i.e. the
// literal is the *whole* argument) -- without it, a concatenated computed
// address like Range("A1:A" & row) or Range("U" & row & ":W" & row) would
// capture just the literal fragment ("A1:A", "U") and that fragment often
// doesn't look like a cell address on its own, misclassifying an ordinary
// computed range as a named-range reference (caught via the real test
// workbook, not a synthetic case -- this pattern is common in real code).
const RX_RANGE_REF = /\bRange\s*\(\s*"([^"]*)"\s*(?=[),])/gi;
const RX_NAMES_REF = /\bNames\s*\(\s*(?:"([^"]*)")?/gi;

// Recognizes classic A1-style cell/range addresses (A1, $A$1, A1:B10, full
// column A:A, full row 1:1) so Range("...") calls that merely address cells
// are excluded -- Excel doesn't allow a defined name shaped like a cell
// address, so anything that doesn't match this is almost certainly a named
// range, not a coincidence.
const RX_LOOKS_LIKE_CELL_ADDRESS =
  /^\$?[A-Za-z]{1,3}\$?\d+(:\$?[A-Za-z]{1,3}\$?\d+)?$|^\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}$|^\$?\d+:\$?\d+$/;

export function scanModuleForReferences(moduleName: string, code: string): ModuleReferenceScan {
  const lines = code.split(/\r\n|\r|\n/);
  const eventProcedures: EventProcedureEntry[] = [];
  const sheetReferences: SheetReferenceEntry[] = [];
  const namedRangeReferences: NamedRangeReferenceEntry[] = [];

  lines.forEach((line, idx) => {
    const lineNo = idx + 1;

    // Skip whole-line comments -- see dependencyScan.ts for the same check
    // and its rationale (trailing inline comments are not stripped).
    if (/^\s*(?:'|Rem\b)/i.test(line)) { return; }

    const eventMatch = RX_EVENT_PROCEDURE.exec(line);
    if (eventMatch) {
      const name = eventMatch[1];
      const prefix = eventMatch[2];
      let eventKind: EventProcedureEntry["eventKind"];
      if (/^Auto_(Open|Close)$/i.test(name)) { eventKind = "auto_macro"; }
      else if (/^Workbook$/i.test(prefix ?? "")) { eventKind = "workbook_event"; }
      else if (/^Worksheet$/i.test(prefix ?? "")) { eventKind = "worksheet_event"; }
      else { eventKind = "userform_event"; }
      eventProcedures.push({ module: moduleName, line: lineNo, procedure: name, eventKind, raw: line.trim() });
    }

    RX_SHEET_REF.lastIndex = 0;
    let sheetMatch: RegExpExecArray | null;
    while ((sheetMatch = RX_SHEET_REF.exec(line)) !== null) {
      const [, api, sheetName] = sheetMatch;
      sheetReferences.push({
        module: moduleName,
        line: lineNo,
        api: api as "Worksheets" | "Sheets",
        sheetName: sheetName ?? null,
        dynamic: sheetName === undefined,
        raw: line.trim(),
      });
    }

    // Range("..."): only kept when the literal does NOT look like a plain
    // cell address (see RX_LOOKS_LIKE_CELL_ADDRESS) -- otherwise ordinary
    // Range("A1")/Range("B2:C10") calls (the overwhelming majority of
    // Range(...) usage in real code) would drown out genuine named-range
    // references. Dynamic Range(variable) calls are not reported at all:
    // most of those are ordinary computed cell addresses (e.g. built from a
    // row/column loop), not named-range access, and including them produced
    // far more noise than signal.
    RX_RANGE_REF.lastIndex = 0;
    let rangeMatch: RegExpExecArray | null;
    while ((rangeMatch = RX_RANGE_REF.exec(line)) !== null) {
      const literal = rangeMatch[1];
      if (!RX_LOOKS_LIKE_CELL_ADDRESS.test(literal)) {
        namedRangeReferences.push({ module: moduleName, line: lineNo, source: "Range", name: literal, dynamic: false, raw: line.trim() });
      }
    }

    RX_NAMES_REF.lastIndex = 0;
    let namesMatch: RegExpExecArray | null;
    while ((namesMatch = RX_NAMES_REF.exec(line)) !== null) {
      const [, name] = namesMatch;
      namedRangeReferences.push({
        module: moduleName,
        line: lineNo,
        source: "Names",
        name: name ?? null,
        dynamic: name === undefined,
        raw: line.trim(),
      });
    }
  });

  return { module: moduleName, eventProcedures, sheetReferences, namedRangeReferences };
}
