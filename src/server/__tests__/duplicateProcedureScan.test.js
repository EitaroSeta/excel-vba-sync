// Regression tests for duplicateProcedureScan.ts (excel_update_module_code's
// duplicateProcedureWarnings, v0.0.65). Motivated by a real case: moving a Function
// to a new class module left the original copy behind in the source standard module,
// unnoticed until reported live -- this scan exists to catch that before the write.
"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const { listProcedureSignatures, findCrossModuleDuplicates } = require("../../../dist-server/duplicateProcedureScan.js");

test("listProcedureSignatures: visibility, kind, and comment-line exclusion", () => {
  const code = [
    "Public Sub DoWork()",
    "End Sub",
    "",
    "Private Function Helper() As Long",
    "End Function",
    "",
    "Function ImplicitlyPublic() As String",
    "End Function",
    "",
    "Public Property Get Value() As Long",
    "End Property",
    "",
    "' Public Sub CommentedOut()",
  ].join("\n");
  const sigs = listProcedureSignatures(code);
  assert.equal(sigs.length, 4);
  assert.deepEqual(
    sigs.map((s) => ({ name: s.name, kind: s.kind, isPublic: s.isPublic })),
    [
      { name: "DoWork", kind: "Sub", isPublic: true },
      { name: "Helper", kind: "Function", isPublic: false },
      { name: "ImplicitlyPublic", kind: "Function", isPublic: true },
      { name: "Value", kind: "Property Get", isPublic: true },
    ]
  );
});

test("findCrossModuleDuplicates: detects a same-name Public procedure in another module, regardless of Sub/Function/Property kind", () => {
  const newCode = "Public Function CalcTotal(x As Long) As Long\nEnd Function\n";
  const otherModules = [
    { name: "ModA", code: "Public Sub CalcTotal()\nEnd Sub\n" }, // same name, different kind -- still a collision
    { name: "ModB", code: "Public Sub Unrelated()\nEnd Sub\n" },
  ];
  const dups = findCrossModuleDuplicates("ModNew", newCode, otherModules);
  assert.equal(dups.length, 1);
  assert.deepEqual(
    { name: dups[0].name, existsInModule: dups[0].existsInModule },
    { name: "CalcTotal", existsInModule: "ModA" }
  );
});

test("findCrossModuleDuplicates: a Private procedure with the same name does not count as a collision", () => {
  const newCode = "Public Sub CalcTotal()\nEnd Sub\n";
  const otherModules = [{ name: "ModA", code: "Private Sub CalcTotal()\nEnd Sub\n" }];
  const dups = findCrossModuleDuplicates("ModNew", newCode, otherModules);
  assert.equal(dups.length, 0);
});

test("findCrossModuleDuplicates: the target module itself is excluded from comparison", () => {
  const newCode = "Public Sub CalcTotal()\nEnd Sub\n";
  const otherModules = [{ name: "ModSelf", code: "Public Sub CalcTotal()\nEnd Sub\n" }];
  const dups = findCrossModuleDuplicates("ModSelf", newCode, otherModules);
  assert.equal(dups.length, 0);
});

test("findCrossModuleDuplicates: name matching is case-insensitive (VBA identifiers are case-insensitive)", () => {
  const newCode = "Public Sub calctotal()\nEnd Sub\n";
  const otherModules = [{ name: "ModA", code: "Public Sub CalcTotal()\nEnd Sub\n" }];
  const dups = findCrossModuleDuplicates("ModNew", newCode, otherModules);
  assert.equal(dups.length, 1);
});

test("findCrossModuleDuplicates: no duplicates returns an empty array", () => {
  const newCode = "Public Sub UniqueName()\nEnd Sub\n";
  const otherModules = [{ name: "ModA", code: "Public Sub SomethingElse()\nEnd Sub\n" }];
  const dups = findCrossModuleDuplicates("ModNew", newCode, otherModules);
  assert.deepEqual(dups, []);
});
