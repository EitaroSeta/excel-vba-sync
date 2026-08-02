// Cross-module duplicate procedure name detection, used by excel_update_module_code's
// dryRun response (see server.ts). Motivation: the tool only writes to the module it's
// told to -- if an agent "moves" a procedure to a new/different module, the original
// copy is left behind unless a separate write also removes it there. This scan flags
// that situation as an advisory warning before the write happens; it never blocks it.
//
// Also flags a softer case, found live: newCode declares a Private procedure whose name
// happens to already be Public in another module. VBA resolves the same-module Private
// first, so this is never a real runtime/compile collision -- but it's still confusing to
// read, and worth surfacing as a lower-severity "risk" distinct from the public-vs-public
// case (which usually means a stale copy was left behind and should probably be removed).
//
// Unicode identifier handling mirrors variableScopeScan.ts's RX_PROC_HEAD (VBA allows
// non-ASCII identifiers, e.g. Japanese procedure names -- see the v0.0.62 fix there for
// why \p{L}\p{N}_ instead of \w is required). This regex additionally captures the
// visibility modifier (RX_PROC_HEAD in variableScopeScan.ts leaves it non-capturing),
// since that's needed here to filter out Private procedures.
const RX_PROC_HEAD_VIS = /^\s*((?:Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+([\p{L}_][\p{L}\p{N}_]*)\s*\(/iu;

export interface ProcedureSignature {
  name: string;
  kind: string;
  line: number;
  isPublic: boolean;
}

export function listProcedureSignatures(code: string): ProcedureSignature[] {
  const lines = code.split(/\r\n|\r|\n/);
  const results: ProcedureSignature[] = [];
  lines.forEach((line, idx) => {
    if (/^\s*(?:'|Rem\b)/i.test(line)) { return; }
    const m = RX_PROC_HEAD_VIS.exec(line);
    if (!m) { return; }
    const vis = (m[1] || "").trim().toLowerCase();
    results.push({
      name: m[3],
      kind: m[2].replace(/\s+/g, " ").trim(),
      line: idx + 1,
      isPublic: vis !== "private",
    });
  });
  return results;
}

export interface CrossModuleDuplicate {
  name: string;
  kind: string;
  existsInModule: string;
  existsAtLine: number;
  risk: "public_duplicate" | "private_name_reused";
}

export function findCrossModuleDuplicates(
  targetModule: string,
  newCode: string,
  otherModules: { name: string; code: string }[]
): CrossModuleDuplicate[] {
  const mine = listProcedureSignatures(newCode);
  if (mine.length === 0) { return []; }

  const warnings: CrossModuleDuplicate[] = [];
  for (const mod of otherModules) {
    if (mod.name === targetModule) { continue; }
    // Only the OTHER module's Public-equivalent procedures are checked against: a Private
    // procedure there can never be part of a naming collision (own-module Private wins
    // locally, and it isn't visible from outside its module either way).
    const theirs = listProcedureSignatures(mod.code).filter((p) => p.isPublic);
    for (const candidate of mine) {
      const hit = theirs.find((p) => p.name.toLowerCase() === candidate.name.toLowerCase());
      if (hit) {
        warnings.push({
          name: candidate.name,
          kind: candidate.kind,
          existsInModule: mod.name,
          existsAtLine: hit.line,
          risk: candidate.isPublic ? "public_duplicate" : "private_name_reused",
        });
      }
    }
  }
  return warnings;
}
