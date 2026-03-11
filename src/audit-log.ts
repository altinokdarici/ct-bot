import { appendFileSync, mkdirSync } from "node:fs";
import { join } from "node:path";

const AUDIT_DIR = join(import.meta.dirname ?? process.cwd(), "..", "audit-logs");

// Ensure directory exists on startup
mkdirSync(AUDIT_DIR, { recursive: true });

function sanitizeFilename(threadId: string): string {
  return threadId.replace(/[^a-zA-Z0-9_-]/g, "_");
}

function timestamp(): string {
  return new Date().toISOString();
}

export function auditLog(threadId: string, entry: {
  direction: "incoming" | "outgoing";
  from?: string;
  type: string;
  text: string;
  sessionId?: string | null;
  meta?: Record<string, unknown>;
}): void {
  const filename = `${sanitizeFilename(threadId)}.jsonl`;
  const filepath = join(AUDIT_DIR, filename);
  const line = JSON.stringify({
    ts: timestamp(),
    threadId,
    ...entry,
  });
  try {
    appendFileSync(filepath, line + "\n");
  } catch (err) {
    console.error(`[audit] Failed to write: ${(err as Error).message}`);
  }
}
