// Best-effort masking of hardcoded credentials before VBA source text (or excerpts of it)
// leaves this process in any MCP tool response. Motivation: every tool that returns
// verbatim/near-verbatim source lines is a channel through which a credential hardcoded
// in a user's VBA project reaches a cloud-hosted AI model -- this exists to reduce that,
// not to guarantee it (regex over text, not a real parser; a secret with no recognizable
// keyword nearby will not be caught). Always applied, with no opt-out: an AI agent has no
// way to know it should ask for redaction before the fact, so making it optional would
// defeat the purpose.
//
// Only the captured VALUE is replaced with [REDACTED] -- the surrounding line is left
// intact, so a reader (human or AI) can still see "there is a hardcoded credential here"
// as a code-quality signal, without seeing the credential itself.

const REDACTED = "[REDACTED]";

const PATTERNS: RegExp[] = [
  // Direct assignment to a credential-shaped identifier: Password = "...", ApiKey = "...", etc.
  // The optional As-clause matters more than it looks: a declaration-with-initializer is the
  // most natural way to hardcode a credential in VBA (Const API_KEY As String = "..."), and
  // requiring the '=' to follow the identifier directly let exactly that form through while
  // catching the bare assignment on the next line. Fixed-length strings (As String * 10) are
  // allowed for completeness.
  /((?:Password|Pwd|ApiKey|Api_Key|SecretKey|Secret|Token|AccessKey|ClientSecret)(?:\s+As\s+[A-Za-z_][A-Za-z0-9_.]*(?:\s*\*\s*\d+)?)?\s*=\s*")([^"]*)(")/gi,
  // Credential embedded inside a larger string (e.g. an ADODB connection string), terminated
  // by ; or the closing quote rather than being its own VBA assignment statement. The
  // terminator is captured (not just look-ahead) so every pattern here has the same
  // (prefix, value, suffix) shape for the replace callback below. The value's first
  // character explicitly excludes whitespace AND '"': without that, this pattern can
  // re-match text already redacted by the pattern above it (Password = "[REDACTED]" --
  // \s*=\s* backtracks to hand the space, or even the opening quote, to the value group,
  // corrupting an already-correct redaction). Caught by this file's own unit tests.
  /((?:Password|Pwd)\s*=\s*)([^;"\s][^;"]*)([;"])/gi,
  // HTTP Authorization header, either via .setRequestHeader("Authorization", "...") or a
  // generic Authorization = "..." / Authorization: "..." style assignment.
  /(setRequestHeader\s*\(?\s*"Authorization"\s*,\s*")([^"]*)(")/gi,
  /(Authorization\s*[:=]\s*")([^"]*)(")/gi,
];

export function redactSecrets(line: string): string {
  let result = line;
  for (const pattern of PATTERNS) {
    result = result.replace(pattern, (_match, prefix: string, _value: string, suffix: string) => `${prefix}${REDACTED}${suffix}`);
  }
  return result;
}

// Applies redactSecrets() line-by-line to a full module's source (or any multi-line code
// text), for the tools that return an entire module's code rather than one matched line.
export function redactCodeText(code: string): string {
  return code.split(/\r\n|\r|\n/).map(redactSecrets).join("\r\n");
}
