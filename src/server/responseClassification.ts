// Pure text/JSON processing shared by most MCP tool handlers in server.ts, extracted out
// so it can be covered by the persistent node:test suite -- neither function touches COM
// or spawns PowerShell; they only interpret text that some other caller already obtained
// (execFileAsync's stdout). Before this extraction, this logic lived inside server.ts and
// was only ever exercised by disposable, one-off JSON-RPC harness scripts against a real
// Excel instance, with no automated regression coverage for its branching (success vs.
// ERR_-prefixed vs. ok:false) -- exactly the kind of gap the project's node:test suite
// exists to close for logic that doesn't actually require COM to test.
import { redactSecrets, redactCodeText } from "./secretRedaction.js";

export type ClassifiedResult = { content: { type: "text"; text: string }[]; isError?: boolean };

// Looks at PowerShell's JSON stdout and flags it isError:true if it's an ERR_-prefixed
// error code or ok:false. Note: isError NOT being set (ok:true) only means the script
// completed without throwing -- it does not confirm a macro did what was intended.
export function classifyResult(outText: string): ClassifiedResult {
  try {
    const start = Math.min(
      ...['{', '['].map(ch => { const i = outText.indexOf(ch); return i === -1 ? Number.POSITIVE_INFINITY : i; })
    );
    if (Number.isFinite(start)) {
      const payload = JSON.parse(outText.slice(start));
      if (payload && typeof payload.error === "string" && payload.error.startsWith("ERR_")) {
        return { content: [{ type: "text", text: outText }], isError: true };
      }
      if (payload && payload.ok === false) {
        return { content: [{ type: "text", text: outText }], isError: true };
      }
    }
  } catch { /* non-JSON output is returned as-is */ }
  return { content: [{ type: "text", text: outText }] };
}

// classifyResult() variant that also redacts hardcoded-credential-shaped values out of the
// parsed JSON payload's specified fields before turning it back into text -- always
// applied, no opt-out (see secretRedaction.ts for why). Falls back to classifyResult's own
// raw passthrough if JSON extraction/parsing fails, so non-JSON or malformed output is
// handled identically to before this existed.
export function classifyResultWithRedaction(
  outText: string,
  opts: { stringFields?: string[]; arrayField?: { field: string; subField: string } }
): ClassifiedResult {
  try {
    const start = Math.min(
      ...['{', '['].map(ch => { const i = outText.indexOf(ch); return i === -1 ? Number.POSITIVE_INFINITY : i; })
    );
    if (Number.isFinite(start)) {
      const payload = JSON.parse(outText.slice(start));
      if (payload && typeof payload === "object") {
        if (opts.stringFields) {
          for (const f of opts.stringFields) {
            if (typeof payload[f] === "string") { payload[f] = redactCodeText(payload[f]); }
          }
        }
        if (opts.arrayField && Array.isArray(payload[opts.arrayField.field])) {
          for (const item of payload[opts.arrayField.field]) {
            if (item && typeof item[opts.arrayField.subField] === "string") {
              item[opts.arrayField.subField] = redactSecrets(item[opts.arrayField.subField]);
            }
          }
        }
        const isErr = (typeof payload.error === "string" && payload.error.startsWith("ERR_")) || payload.ok === false;
        return { content: [{ type: "text", text: JSON.stringify(payload, null, 2) }], isError: isErr || undefined };
      }
    }
  } catch { /* fall through to classifyResult's own raw-passthrough behavior */ }
  return classifyResult(outText);
}
