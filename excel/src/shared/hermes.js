/* global fetch */
// One place both the task pane and the custom functions call Hermes.
// Caddy (https://localhost:8643) terminates TLS and injects the Authorization
// header, so we never send the API key from here.

const ENDPOINT = "https://localhost:8643/v1/chat/completions";

export async function askHermes(messages, { idempotencyKey } = {}) {
  // Note: only Authorization, Content-Type, and Idempotency-Key are allowed by
  // the API server's CORS policy. Don't add other custom headers or the browser
  // preflight will fail ("Failed to fetch"). Conversation continuity comes from
  // sending the full message history each call.
  const headers = { "Content-Type": "application/json" };
  if (idempotencyKey) headers["Idempotency-Key"] = idempotencyKey;

  const res = await fetch(ENDPOINT, {
    method: "POST",
    headers,
    body: JSON.stringify({ model: "hermes-agent", messages }),
  });
  if (!res.ok) {
    throw new Error(`Hermes ${res.status}: ${await res.text()}`);
  }
  const data = await res.json();
  return data.choices[0].message.content;
}
