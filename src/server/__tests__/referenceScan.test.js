// Persisted regression tests for referenceScan.ts (vba_list_references,
// v0.0.59), including the real false positive found against the test
// workbook (concatenated computed range addresses misread as named ranges).
"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const { scanModuleForReferences } = require("../../../dist-server/referenceScan.js");

test("event procedures: Workbook_/Worksheet_/UserForm_/Auto_Open, and non-events excluded", () => {
  const code = [
    "Private Sub Workbook_Open()",
    "End Sub",
    "",
    "Private Sub Worksheet_Change(ByVal Target As Range)",
    "End Sub",
    "",
    "Private Sub UserForm_Initialize()",
    "End Sub",
    "",
    "Public Sub Auto_Open()",
    "End Sub",
    "",
    "Sub NotAnEvent_ButLooksLikeIt_Change()",
    "End Sub",
  ].join("\n");
  const r = scanModuleForReferences("M", code);
  assert.equal(r.eventProcedures.length, 4);
  assert.deepEqual(
    r.eventProcedures.map((e) => e.eventKind),
    ["workbook_event", "worksheet_event", "userform_event", "auto_macro"]
  );
});

test("sheet references: static vs dynamic, Worksheets and Sheets", () => {
  const code = ['Set ws = Worksheets("Sheet1")', "Set ws = Sheets(dynamicSheetName)"].join("\n");
  const r = scanModuleForReferences("M", code);
  assert.equal(r.sheetReferences.length, 2);
  assert.deepEqual(
    { api: r.sheetReferences[0].api, sheetName: r.sheetReferences[0].sheetName, dynamic: r.sheetReferences[0].dynamic },
    { api: "Worksheets", sheetName: "Sheet1", dynamic: false }
  );
  assert.equal(r.sheetReferences[1].dynamic, true);
});

test("named-range references: Range excludes cell addresses, Names includes dynamic", () => {
  const code = [
    'v = Range("A1").Value',
    'v = Range("B2:C10").Value',
    'v = Range("MyNamedRange").Value',
    "v = Range(dynamicCellRef).Value",
    'Set n = Names("AnotherNamedRange")',
    "Set n = Names(dynamicNameVar)",
  ].join("\n");
  const r = scanModuleForReferences("M", code);
  // Only the genuine named-range cases should survive: A1/B2:C10 (address-shaped)
  // and the dynamic Range() call are all excluded by design.
  assert.deepEqual(
    r.namedRangeReferences.map((n) => `${n.source}:${n.name}:${n.dynamic}`),
    ["Range:MyNamedRange:false", "Names:AnotherNamedRange:false", "Names:null:true"]
  );
});

test("real false positive (caught against the test workbook): concatenated computed range addresses", () => {
  const code = [
    'ws.Range("A1:A" & x - 1).Select',
    'outList1.Add ws.Range("U" & T & ":W" & T).Value',
    'targetws.Range("B5:B" & Rows.count).ClearContents',
    'v = Range("MyNamedRange").Value',
  ].join("\n");
  const r = scanModuleForReferences("M", code);
  assert.equal(r.namedRangeReferences.length, 1);
  assert.equal(r.namedRangeReferences[0].name, "MyNamedRange");
});

test("whole-line comments are excluded", () => {
  const code = ["' Private Sub Workbook_BeforeClose(Cancel As Boolean)", "' v = Range(\"CommentedOutRange\").Value"].join("\n");
  const r = scanModuleForReferences("M", code);
  assert.equal(r.eventProcedures.length, 0);
  assert.equal(r.namedRangeReferences.length, 0);
});

test("a hardcoded credential on a matched line is redacted in the raw field (v0.0.72)", () => {
  const code = 'Set ws = Worksheets("Config"): Password = "hunter2"';
  const r = scanModuleForReferences("M", code);
  assert.equal(r.sheetReferences.length, 1);
  assert.equal(r.sheetReferences[0].raw, 'Set ws = Worksheets("Config"): Password = "[REDACTED]"');
});
