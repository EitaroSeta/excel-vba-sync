// Persisted regression tests for dependencyScan.ts, covering cases found
// while building vba_list_dependencies (v0.0.56-0.0.58) and the real
// false positive caught against the test workbook (commented-out lines).
//
// Runs against the already-compiled dist-server/dependencyScan.js -- no
// TypeScript compilation step for the tests themselves. Run with:
//   npm run test:unit
"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const { scanModuleForDependencies } = require("../../../dist-server/dependencyScan.js");

test("Declare Sub/Function: scope, PtrSafe, Alias", () => {
  const code = [
    "Private Declare PtrSafe Function GetTickCount Lib \"kernel32\" () As Long",
    "Private Declare Function OldStyleApi Lib \"user32.dll\" Alias \"MessageBoxA\" (ByVal h As Long) As Long",
  ].join("\n");
  const r = scanModuleForDependencies("M", code);
  assert.equal(r.apiDeclares.length, 2);
  assert.deepEqual(
    { scope: r.apiDeclares[0].scope, ptrSafe: r.apiDeclares[0].ptrSafe, lib: r.apiDeclares[0].lib },
    { scope: "Private", ptrSafe: true, lib: "kernel32" }
  );
  assert.deepEqual(
    { alias: r.apiDeclares[1].alias },
    { alias: "MessageBoxA" }
  );
});

test("CreateObject: static ProgID vs dynamic", () => {
  const code = [
    'Set fso = CreateObject("Scripting.FileSystemObject")',
    "Set dynObj = CreateObject(progId)",
  ].join("\n");
  const r = scanModuleForDependencies("M", code);
  assert.equal(r.comObjects.length, 2);
  assert.deepEqual(
    { progId: r.comObjects[0].progId, dynamic: r.comObjects[0].dynamic },
    { progId: "Scripting.FileSystemObject", dynamic: false }
  );
  assert.deepEqual(
    { progId: r.comObjects[1].progId, dynamic: r.comObjects[1].dynamic },
    { progId: null, dynamic: true }
  );
});

test("Shell: both call syntaxes, no false positive on string mention", () => {
  const code = [
    'Shell "notepad.exe"',
    'Shell("calc.exe", vbNormalFocus)',
    'msg = "Please run the Shell command manually"',
  ].join("\n");
  const r = scanModuleForDependencies("M", code);
  assert.equal(r.shellCalls.length, 2);
});

test("Application.Run: literal, cross-workbook literal, dynamic", () => {
  const code = [
    'Application.Run "Module1.DoWork"',
    'Application.Run "\'OtherBook.xlsm\'!Module1.DoWork"',
    "Application.Run dynamicMacroName",
  ].join("\n");
  const r = scanModuleForDependencies("M", code);
  assert.equal(r.applicationRunCalls.length, 3);
  assert.equal(r.applicationRunCalls[0].target, "Module1.DoWork");
  assert.equal(r.applicationRunCalls[1].target, "'OtherBook.xlsm'!Module1.DoWork");
  assert.equal(r.applicationRunCalls[2].dynamic, true);
});

test("File I/O: native statements", () => {
  const code = [
    'Open "C:\\out\\log.txt" For Append As #fnum',
    'Kill "C:\\temp\\old.txt"',
    'FileCopy "C:\\a.txt", "C:\\b.txt"',
    'MkDir "C:\\temp\\newfolder"',
    'RmDir "C:\\temp\\oldfolder"',
  ].join("\n");
  const r = scanModuleForDependencies("M", code);
  assert.deepEqual(r.fileIo.map((f) => f.operation), ["open", "kill", "filecopy", "mkdir", "rmdir"]);
  assert.ok(r.fileIo.every((f) => f.methodNameOnly === false));
});

test("FileSystemObject methods: with and without parens", () => {
  const code = [
    'Set f = fso.OpenTextFile("C:\\a.txt", 1)',
    'fso.CopyFile "C:\\a.txt", "C:\\c.txt"',
    'fso.DeleteFile "C:\\a.txt"',
    'fso.CreateFolder "C:\\newdir"',
  ].join("\n");
  const r = scanModuleForDependencies("M", code);
  assert.equal(r.fileIo.length, 4);
  assert.ok(r.fileIo.every((f) => f.methodNameOnly === true));
  assert.deepEqual(r.fileIo.map((f) => f.operation), ["fso_opentextfile", "fso_copyfile", "fso_deletefile", "fso_createfolder"]);
});

test("Workbooks.Open: static path vs dynamic", () => {
  const code = [
    'Set wb = Workbooks.Open("C:\\data\\linked.xlsx")',
    "Set wb = Workbooks.Open(dynamicPath)",
  ].join("\n");
  const r = scanModuleForDependencies("M", code);
  assert.equal(r.externalWorkbooks.length, 2);
  assert.equal(r.externalWorkbooks[0].target, "C:\\data\\linked.xlsx");
  assert.equal(r.externalWorkbooks[1].dynamic, true);
});

test("whole-line comments are excluded from every detector", () => {
  const code = [
    "' Private Declare Function CommentedOutApi Lib \"user32.dll\" () As Long",
    "' Shell \"commented.exe\"",
    "Rem Kill \"C:\\temp\\rem-style-comment.txt\"",
    '\' fso.DeleteFile "C:\\commented.txt"',
  ].join("\n");
  const r = scanModuleForDependencies("M", code);
  assert.equal(r.apiDeclares.length, 0);
  assert.equal(r.shellCalls.length, 0);
  assert.equal(r.fileIo.length, 0);
});

test("real false positive (caught against the test workbook): commented-out Open statement", () => {
  const code = ["    'Open filePath For Output As #fnum", "    Open filePath For Append As #fnum"].join("\n");
  const r = scanModuleForDependencies("M", code);
  assert.equal(r.fileIo.length, 1);
  assert.equal(r.fileIo[0].raw, "Open filePath For Append As #fnum");
});

test("a hardcoded credential on a matched line is redacted in the raw field (v0.0.72)", () => {
  const code = 'Shell "net use \\\\server /user:admin Password=hunter2"';
  const r = scanModuleForDependencies("M", code);
  assert.equal(r.shellCalls.length, 1);
  assert.equal(r.shellCalls[0].raw, 'Shell "net use \\\\server /user:admin Password=[REDACTED]"');
});
