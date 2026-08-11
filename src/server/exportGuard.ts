// Decision logic for excel_export_module's "modified since last export" guard.
//
// The gap this closes: excel_update_module_code's optimistic lock
// (ERR_MODULE_CHANGED_SINCE_DRYRUN) watches only the VBE side, so a human editing
// the EXPORTED FILE in VS Code has no protection -- an AI's post-write export lands
// on top of their saved-but-not-yet-imported edits. The pre-overwrite backup
// (v0.0.79) makes that recoverable, but silently: the human keeps believing their
// edits exist. This guard turns the silent overwrite into an explicit refusal.
//
// Mechanism: after each successful export, the tool records a hash of the file it
// wrote (a sidecar next to the backups). On the next export, if the file on disk no
// longer matches that hash, someone else changed it -- refuse unless force:true.
//
// Kept COM/fs-free so it is unit-testable; the caller supplies the hashes.

export const ERR_EXPORTED_FILE_MODIFIED = "ERR_EXPORTED_FILE_MODIFIED";

export type ExportGuardInput = {
  // Hash of the exported file currently on disk; null if the file does not exist.
  currentFileHash: string | null;
  // Hash recorded at this tool's last successful export; null if no record exists
  // (first export, a manual "Export All Modules" run, or a pre-v0.0.80 export --
  // provenance unknown, so the guard stays out of the way).
  lastExportHash: string | null;
  force: boolean;
};

export type ExportGuardDecision =
  | { ok: true }
  | { ok: false; error: string; detail: string };

export function evaluateExportGuard(input: ExportGuardInput): ExportGuardDecision {
  const { currentFileHash, lastExportHash, force } = input;

  // Nothing on disk to protect.
  if (currentFileHash === null) { return { ok: true }; }
  // The caller explicitly chose to overwrite (a backup is still taken).
  if (force) { return { ok: true }; }
  // No record of what this tool last wrote -- cannot tell a human edit from a
  // manual export, so do not block. The pre-overwrite backup still applies.
  if (lastExportHash === null) { return { ok: true }; }
  // File is exactly what this tool last wrote -- safe to replace.
  if (currentFileHash === lastExportHash) { return { ok: true }; }

  return {
    ok: false,
    error: ERR_EXPORTED_FILE_MODIFIED,
    detail:
      "The exported file changed since this tool last wrote it -- likely a human's edit that has NOT been imported into Excel yet. " +
      "Overwriting now would discard their work. Ask the user; then either import the file first (so Excel has their edit), " +
      "or re-call with force:true to overwrite anyway (the previous file is still backed up).",
  };
}
