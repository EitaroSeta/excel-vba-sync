// Regression tests for responseClassification.ts, extracted from server.ts (v0.0.73) so
// this branching logic gets persistent regression coverage instead of relying only on
// disposable, one-off JSON-RPC harness checks against a real Excel instance. Motivated by
// a direct question from the user: "if you change branching logic that affects an existing
// route, shouldn't there be a regression test exercising that existing route too?" -- there
// was not, for this specific logic, until this file.
"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const { classifyResult, classifyResultWithRedaction } = require("../../../dist-server/responseClassification.js");

test("classifyResult: ok:true payload is not an error", () => {
  const r = classifyResult(JSON.stringify({ ok: true, module: "M", code: "Option Explicit" }));
  assert.equal(r.isError, undefined);
  assert.equal(r.content[0].text, JSON.stringify({ ok: true, module: "M", code: "Option Explicit" }));
});

test("classifyResult: ok:false payload is an error", () => {
  const r = classifyResult(JSON.stringify({ ok: false, error: "module_not_found", module: "ZZ" }));
  assert.equal(r.isError, true);
});

test("classifyResult: an ERR_-prefixed error string is an error, independent of ok", () => {
  const r = classifyResult(JSON.stringify({ error: "ERR_VBOM_TRUST_DISABLED" }));
  assert.equal(r.isError, true);
});

test("classifyResult: non-JSON output passes through unchanged, not flagged as an error", () => {
  const raw = "not json at all, e.g. a PowerShell crash dump";
  const r = classifyResult(raw);
  assert.equal(r.isError, undefined);
  assert.equal(r.content[0].text, raw);
});

test("classifyResult: a preamble before the JSON (e.g. -File's 'Input File:' banner) is preserved verbatim", () => {
  const raw = '----------------------------------------\nInput File: \n{"ok":true,"module":"M"}';
  const r = classifyResult(raw);
  assert.equal(r.isError, undefined);
  assert.equal(r.content[0].text, raw);
});

test("classifyResultWithRedaction: redacts a top-level string field and preserves ok:true", () => {
  const raw = JSON.stringify({ ok: true, module: "M", code: 'Public Const API_KEY = "sk-abc123"' });
  const r = classifyResultWithRedaction(raw, { stringFields: ["code"] });
  assert.equal(r.isError, undefined);
  const parsed = JSON.parse(r.content[0].text);
  assert.equal(parsed.code, 'Public Const API_KEY = "[REDACTED]"');
});

test("classifyResultWithRedaction: redacts a subfield within an array field (vba_search_code's hits[].snippet)", () => {
  const raw = JSON.stringify({
    ok: true,
    hits: [
      { module: "M", line: 3, snippet: 'conn.Open "Password=hunter2;"' },
      { module: "M", line: 9, snippet: "Dim x As Long" },
    ],
  });
  const r = classifyResultWithRedaction(raw, { arrayField: { field: "hits", subField: "snippet" } });
  const parsed = JSON.parse(r.content[0].text);
  assert.equal(parsed.hits[0].snippet, 'conn.Open "Password=[REDACTED];"');
  assert.equal(parsed.hits[1].snippet, "Dim x As Long");
});

test("classifyResultWithRedaction: still correctly flags an error payload as isError, same as classifyResult", () => {
  const raw = JSON.stringify({ ok: false, error: "module_not_found", module: "ZZ" });
  const r = classifyResultWithRedaction(raw, { stringFields: ["code"] });
  assert.equal(r.isError, true);
});

test("classifyResultWithRedaction: falls back to classifyResult's raw passthrough when JSON parsing fails", () => {
  const raw = "not json at all";
  const r = classifyResultWithRedaction(raw, { stringFields: ["code"] });
  assert.equal(r.isError, undefined);
  assert.equal(r.content[0].text, raw);
});

test("classifyResultWithRedaction: code without any secret pattern is returned unchanged in substance", () => {
  const raw = JSON.stringify({ ok: true, module: "M", code: "Dim x As Long\r\nx = 1" });
  const r = classifyResultWithRedaction(raw, { stringFields: ["code"] });
  const parsed = JSON.parse(r.content[0].text);
  assert.equal(parsed.code, "Dim x As Long\r\nx = 1");
});
