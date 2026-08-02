// Regression tests for duplicateProcedureScan.ts (excel_update_module_code's
// duplicateProcedureWarnings, v0.0.65-v0.0.66). Motivated by two real cases found live:
// (1) moving a Function to a new class module left the original copy behind in the
// source standard module, unnoticed -- the original public_duplicate check;
// (2) a newly-added module's own Private helper happened to share a name (IsPrime) with
// two other modules' Public functions -- no functional collision (VBA resolves the local
// Private first), but confusing enough that it's worth its own lower-severity risk tier
// (private_name_reused), added in v0.0.66 after this was found via manual full-text
// search rather than being caught by the original public-only check.
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
    { name: dups[0].name, existsInModule: dups[0].existsInModule, risk: dups[0].risk },
    { name: "CalcTotal", existsInModule: "ModA", risk: "public_duplicate" }
  );
});

test("findCrossModuleDuplicates: the OTHER module's Private procedure with the same name does not count as a collision", () => {
  const newCode = "Public Sub CalcTotal()\nEnd Sub\n";
  const otherModules = [{ name: "ModA", code: "Private Sub CalcTotal()\nEnd Sub\n" }];
  const dups = findCrossModuleDuplicates("ModNew", newCode, otherModules);
  assert.equal(dups.length, 0);
});

test("findCrossModuleDuplicates: newCode's own Private procedure sharing a name with a Public one elsewhere is flagged as private_name_reused (real case: a newly-added module's Private helper shared a name with two existing modules' Public functions)", () => {
  const newCode = "Private Function CalcSomething(n As Long) As Boolean\nEnd Function\n";
  const otherModules = [
    { name: "ModExisting1", code: "Public Function CalcSomething(n As Long) As Boolean\nEnd Function\n" },
    { name: "ModExisting2", code: "Public Function CalcSomething(n As Long) As Boolean\nEnd Function\n" },
  ];
  const dups = findCrossModuleDuplicates("ModNew", newCode, otherModules);
  assert.equal(dups.length, 2);
  assert.ok(dups.every((d) => d.risk === "private_name_reused"));
  assert.deepEqual(dups.map((d) => d.existsInModule).sort(), ["ModExisting1", "ModExisting2"]);
});

test("findCrossModuleDuplicates: newCode's Private procedure vs. another module's Private procedure of the same name is NOT flagged (no functional or readability risk)", () => {
  const newCode = "Private Sub Helper()\nEnd Sub\n";
  const otherModules = [{ name: "ModA", code: "Private Sub Helper()\nEnd Sub\n" }];
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
