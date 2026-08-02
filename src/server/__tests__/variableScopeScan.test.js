// Persisted regression tests for variableScopeScan.ts (vba_list_variable_scopes,
// v0.0.60-0.0.62). Includes the two real bugs found in production: the
// Public Const / indentation parsing bugs (caught by unit testing before
// release) and the Japanese/Unicode identifier bug (found live via Copilot
// Chat against the real workbook, fixed same-day in v0.0.62) -- the latter
// is exactly the kind of regression this suite exists to prevent from
// recurring silently.
"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const { scanModuleForVariableScopes, resolveVariableUsages } = require("../../../dist-server/variableScopeScan.js");

test("declaration scope classification: module, public, procedure, and Public Const", () => {
  const code = [
    "Private moduleLevelVar As Long",
    "Public sharedCounter As Long",
    "Dim implicitPrivateVar As String",
    "Const MODULE_CONST = 42",
    'Public Const SHARED_CONST = "x"',
    'Private Const PRIVATE_CONST = "y"',
    "",
    "Sub ProcA()",
    "    Dim x As Long",
    "    Dim y As Long, z As String",
    "    Static persistentCounter As Long",
    "    Const LOCAL_CONST = 1",
    "    Dim arr(1 To 10, 1 To 5) As Variant",
    "    ' Dim commentedOut As Long",
    "End Sub",
  ].join("\n");
  const r = scanModuleForVariableScopes("M", code);
  const byName = (n) => r.declarations.filter((d) => d.name === n);

  assert.equal(byName("moduleLevelVar")[0].scope, "module");
  assert.equal(byName("sharedCounter")[0].scope, "public");
  // Regression: the keyword parser originally only recognized a single
  // leading declaration keyword, so "Public Const NAME = ..." had "Const"
  // itself misread as the variable name.
  assert.deepEqual(
    { name: byName("SHARED_CONST")[0]?.name, scope: byName("SHARED_CONST")[0]?.scope, kind: byName("SHARED_CONST")[0]?.kind },
    { name: "SHARED_CONST", scope: "public", kind: "constant" }
  );
  assert.equal(byName("PRIVATE_CONST")[0].scope, "module");
  // Regression: the Dim/Static/Const sub-matchers originally required zero
  // leading whitespace, so every indented declaration inside a procedure
  // body (i.e. nearly all of them) silently failed to match at all.
  assert.equal(byName("x").length, 1);
  assert.equal(byName("x")[0].declaredIn, "ProcA");
  assert.equal(byName("y").length, 1);
  assert.equal(byName("z").length, 1);
  assert.equal(byName("persistentCounter")[0].declaredIn, "ProcA");
  assert.deepEqual({ scope: byName("LOCAL_CONST")[0].scope, kind: byName("LOCAL_CONST")[0].kind }, { scope: "procedure", kind: "constant" });
  // Array bounds comma (1 To 10, 1 To 5) must not cause a false split.
  assert.equal(byName("arr").length, 1);
  assert.equal(byName("commentedOut").length, 0);
});

test("Static Sub / Public Sub headers are not misread as variable declarations", () => {
  const code = ["Static Sub LegacyStaticSub()", "End Sub", "", "Public Sub PublicSubHeader()", "End Sub"].join("\n");
  const r = scanModuleForVariableScopes("M", code);
  assert.equal(r.declarations.length, 0);
});

test("Japanese/Unicode procedure and variable names (real bug, found live in v0.0.61, fixed in v0.0.62)", () => {
  const code = [
    'Public Sub 出力処理(cb As String)',
    "    Dim targetws As Worksheet",
    '    Set targetws = ThisWorkbook.Worksheets("Data")',
    "End Sub",
  ].join("\n");
  const r = scanModuleForVariableScopes("M", code);
  const targetws = r.declarations.find((d) => d.name === "targetws");
  assert.equal(targetws.scope, "procedure");
  assert.equal(targetws.declaredIn, "出力処理");

  const usage = resolveVariableUsages("M", "targetws", null, [{ name: "M", code }]);
  assert.equal(usage.ok, true);
  assert.equal(usage.usages.length, 1);
  assert.equal(usage.usages[0].kind, "write");
});

test("resolveVariableUsages: procedure scope confines usages to the declaring procedure", () => {
  const code = [
    "Sub ProcA()",
    "    Dim localVar As Long",
    "    localVar = 1",
    "    Debug.Print localVar",
    "End Sub",
    "",
    "Sub ProcB()",
    "    Dim localVar As Long",
    "    localVar = 2",
    "End Sub",
  ].join("\n");
  const modules = [{ name: "M", code }];

  const ambiguous = resolveVariableUsages("M", "localVar", null, modules);
  assert.equal(ambiguous.ok, false);
  assert.equal(ambiguous.error, "ambiguous_declaration");
  assert.deepEqual(ambiguous.candidates.map((c) => c.declaredIn).sort(), ["ProcA", "ProcB"]);

  const resolved = resolveVariableUsages("M", "localVar", "ProcA", modules);
  assert.equal(resolved.ok, true);
  assert.equal(resolved.usages.length, 2);
  assert.equal(resolved.usages[0].kind, "write");
  assert.equal(resolved.usages[1].kind, "reference");
});

test("resolveVariableUsages: module/public scope excludes procedures that shadow the name locally", () => {
  const modA = [
    "Private sharedState As Long",
    "Public globalCounter As Long",
    "Private moduleOnlyVar As Long",
    "",
    "Sub ProcC()",
    "    Dim sharedState As Long",
    "    sharedState = 99",
    "End Sub",
    "",
    "Sub ProcD()",
    "    sharedState = sharedState + 5",
    "    moduleOnlyVar = moduleOnlyVar + 1",
    "End Sub",
  ].join("\n");
  const modB = ["Sub UseGlobal()", "    globalCounter = globalCounter + 1", "End Sub", "", "Sub ShadowGlobal()", "    Dim globalCounter As Long", "    globalCounter = 0", "End Sub"].join(
    "\n"
  );
  const modules = [
    { name: "ModA", code: modA },
    { name: "ModB", code: modB },
  ];

  // A module-level declaration that's ALSO shadowed by a same-named local in
  // the same module is ambiguous by design -- never auto-resolved.
  const ambiguous = resolveVariableUsages("ModA", "sharedState", null, modules);
  assert.equal(ambiguous.ok, false);
  assert.equal(ambiguous.error, "ambiguous_declaration");

  // Disambiguated to the local one specifically.
  const local = resolveVariableUsages("ModA", "sharedState", "ProcC", modules);
  assert.equal(local.ok, true);
  assert.equal(local.declaration.scope, "procedure");
  assert.equal(local.usages.length, 1); // "sharedState = 99"; the Dim line itself is excluded

  // A pure module-scoped variable with no same-module shadow resolves cleanly.
  const clean = resolveVariableUsages("ModA", "moduleOnlyVar", null, modules);
  assert.equal(clean.ok, true);
  assert.equal(clean.declaration.scope, "module");

  // Public scope: cross-module usage is found, but a DIFFERENT module's own
  // local shadow is correctly excluded.
  const pub = resolveVariableUsages("ModA", "globalCounter", null, modules);
  assert.equal(pub.ok, true);
  assert.equal(pub.declaration.scope, "public");
  const modBUsages = pub.usages.filter((u) => u.module === "ModB");
  assert.equal(modBUsages.length, 1);
  assert.match(modBUsages[0].raw, /\+ 1/);
});

test("resolveVariableUsages: not-found errors", () => {
  const modules = [{ name: "M", code: "Dim x As Long" }];
  const noVar = resolveVariableUsages("M", "noSuchVariable", null, modules);
  assert.equal(noVar.ok, false);
  assert.equal(noVar.error, "declaration_not_found");

  const noModule = resolveVariableUsages("NoSuchModule", "x", null, modules);
  assert.equal(noModule.ok, false);
  assert.equal(noModule.error, "declaration_not_found");
});
