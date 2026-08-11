// Tests for the excel_export_module "modified since last export" guard (v0.0.80).
// The scenario that motivated it: a human edits the exported .bas in VS Code while
// an AI writes the same module via MCP; the AI's post-write export must not land
// silently on the human's saved-but-not-imported edits.
"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const { evaluateExportGuard, ERR_EXPORTED_FILE_MODIFIED } = require("../../../dist-server/exportGuard.js");

test("proceeds when no exported file exists yet", () => {
  assert.deepEqual(
    evaluateExportGuard({ currentFileHash: null, lastExportHash: null, force: false }),
    { ok: true }
  );
  // A stale sidecar without a file (file deleted by hand) must not block either.
  assert.deepEqual(
    evaluateExportGuard({ currentFileHash: null, lastExportHash: "abc", force: false }),
    { ok: true }
  );
});

test("proceeds when the file is exactly what this tool last wrote", () => {
  assert.deepEqual(
    evaluateExportGuard({ currentFileHash: "abc", lastExportHash: "abc", force: false }),
    { ok: true }
  );
});

test("proceeds when provenance is unknown (no sidecar: first export, manual export, pre-v0.0.80)", () => {
  assert.deepEqual(
    evaluateExportGuard({ currentFileHash: "abc", lastExportHash: null, force: false }),
    { ok: true }
  );
});

test("refuses when the file changed since the last export by this tool", () => {
  const d = evaluateExportGuard({ currentFileHash: "human-edit", lastExportHash: "abc", force: false });
  assert.equal(d.ok, false);
  assert.equal(d.error, ERR_EXPORTED_FILE_MODIFIED);
  // The detail must steer the agent to the two legitimate ways out.
  assert.match(d.detail, /import the file first/);
  assert.match(d.detail, /force:true/);
});

test("force:true overrides the refusal", () => {
  assert.deepEqual(
    evaluateExportGuard({ currentFileHash: "human-edit", lastExportHash: "abc", force: true }),
    { ok: true }
  );
});
