// Regression tests for secretRedaction.ts, added after a user flagged a real concern:
// hardcoded credentials in VBA source can reach a cloud-hosted AI model verbatim through
// any tool that returns source text or line excerpts. This is a best-effort, regex-based
// mask (not a real parser) applied unconditionally -- see the file header for why it has
// no opt-out.
"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const { redactSecrets } = require("../../../dist-server/secretRedaction.js");

test("redacts a direct assignment to a credential-shaped identifier", () => {
  assert.equal(redactSecrets('Password = "hunter2"'), 'Password = "[REDACTED]"');
  assert.equal(redactSecrets('ApiKey = "sk-abc123"'), 'ApiKey = "[REDACTED]"');
  assert.equal(redactSecrets('Const SECRET = "topsecret"'), 'Const SECRET = "[REDACTED]"');
});

// A declaration-with-initializer is the most natural way to hardcode a credential in VBA,
// and it was slipping through: the pattern required '=' to follow the identifier directly,
// so the bare assignment was caught but "Const API_KEY As String = ..." was not.
test("redacts a declaration that carries an As-clause before the assignment", () => {
  assert.equal(
    redactSecrets('Const API_KEY As String = "sk-abc123"'),
    'Const API_KEY As String = "[REDACTED]"'
  );
  assert.equal(
    redactSecrets('Private Const SECRET As String = "topsecret"'),
    'Private Const SECRET As String = "[REDACTED]"'
  );
  // Fixed-length string form, allowed by the pattern for completeness.
  assert.equal(
    redactSecrets('Dim thePassword As String * 8 = "hunter2"'),
    'Dim thePassword As String * 8 = "[REDACTED]"'
  );
});

test("still redacts a later assignment when an earlier As-clause does not lead to one", () => {
  // The optional As-clause must not strand the match on the first identifier: here the
  // declaration has no initializer and the real assignment comes after the colon.
  assert.equal(
    redactSecrets('Dim apiKey As String: apiKey = "sk-abc123"'),
    'Dim apiKey As String: apiKey = "[REDACTED]"'
  );
  // Declaration alone carries no value, so there is nothing to mask.
  assert.equal(redactSecrets("Dim clientSecret As String"), "Dim clientSecret As String");
});

test("an As-clause declaration is not corrupted by the connection-string pattern", () => {
  // Pattern order regression guard: the Password/Pwd connection-string pattern runs after
  // this one and must not re-match text that was already replaced.
  assert.equal(
    redactSecrets('Const PWD As String = "hunter2"'),
    'Const PWD As String = "[REDACTED]"'
  );
});

test("is case-insensitive", () => {
  assert.equal(redactSecrets('password = "hunter2"'), 'password = "[REDACTED]"');
  assert.equal(redactSecrets('PASSWORD = "hunter2"'), 'PASSWORD = "[REDACTED]"');
  assert.equal(redactSecrets('PwD = "hunter2"'), 'PwD = "[REDACTED]"');
});

test("redacts a credential embedded inside a connection string", () => {
  const line = 'conn.Open "Provider=SQLOLEDB;Data Source=srv;Password=hunter2;User ID=admin;"';
  assert.equal(
    redactSecrets(line),
    'conn.Open "Provider=SQLOLEDB;Data Source=srv;Password=[REDACTED];User ID=admin;"'
  );
});

test("redacts an HTTP Authorization header set via setRequestHeader", () => {
  const line = 'http.setRequestHeader "Authorization", "Bearer sk-abc123xyz"';
  assert.equal(redactSecrets(line), 'http.setRequestHeader "Authorization", "[REDACTED]"');
});

test("redacts a generic Authorization assignment", () => {
  assert.equal(redactSecrets('Authorization = "Bearer sk-abc123xyz"'), 'Authorization = "[REDACTED]"');
});

test("leaves ordinary code lines with no credential pattern unchanged", () => {
  const lines = [
    "Dim x As Long",
    'Debug.Print "Hello, world"',
    "x = x + 1",
    'ws.Range("A1").Value = "some data"',
  ];
  for (const line of lines) {
    assert.equal(redactSecrets(line), line);
  }
});

test("does not falsely trigger on an identifier that merely contains a credential keyword without an assignment", () => {
  const line = "If IsPasswordValid(inputPwd) Then";
  assert.equal(redactSecrets(line), line);
});
