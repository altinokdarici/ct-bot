import { runClaudeSession, type ClaudeResponse, type ProgressCallback } from "./claude-session.ts";
import type { Session, HandoffRequest } from "./types.ts";
import { readFileSync, writeFileSync, existsSync } from "node:fs";
import { join } from "node:path";

const DEFAULT_CWD = process.env.DEFAULT_CWD || process.cwd();
const SESSIONS_PATH = join(import.meta.dirname ?? process.cwd(), "..", "sessions.json");

function loadSessions(): Map<string, Session> {
  if (existsSync(SESSIONS_PATH)) {
    try {
      const data = JSON.parse(readFileSync(SESSIONS_PATH, "utf8"));
      return new Map(Object.entries(data));
    } catch { /* ignore corrupt file */ }
  }
  return new Map();
}

function saveSessions(sessions: Map<string, Session>): void {
  writeFileSync(SESSIONS_PATH, JSON.stringify(Object.fromEntries(sessions), null, 2));
}

export class SessionManager {
  private sessions: Map<string, Session>;

  constructor() {
    this.sessions = loadSessions();
    console.log(`[session] Loaded ${this.sessions.size} persisted session(s)`);
  }

  private persist(): void {
    saveSessions(this.sessions);
  }

  getSession(threadId: string): Session | undefined {
    return this.sessions.get(threadId);
  }

  listSessions(): Session[] {
    return [...this.sessions.values()];
  }

  async sendMessage(threadId: string, text: string, onProgress?: ProgressCallback, signal?: AbortSignal): Promise<ClaudeResponse> {
    let session = this.sessions.get(threadId);

    if (!session) {
      session = {
        sessionId: null,
        threadId,
        cwd: DEFAULT_CWD,
        status: "active",
        createdAt: Date.now(),
        lastActiveAt: Date.now(),
        handoffContext: null,
      };
      this.sessions.set(threadId, session);
      this.persist();
      console.log(`[session] New session for thread ${threadId} in ${session.cwd}`);
    }

    session.status = "active";
    session.lastActiveAt = Date.now();

    // If there's handoff context from a previous environment, prepend it
    let prompt = text;
    if (session.handoffContext && !session.sessionId) {
      prompt = `[Context from previous environment]\n${session.handoffContext}\n\n[User message]\n${text}`;
      session.handoffContext = null;
      this.persist();
      console.log(`[session] Injected handoff context into prompt`);
    }

    console.log(`[session] ${session.sessionId ? `Resuming ${session.sessionId.slice(0, 8)}` : "New session"} | cwd=${session.cwd}`);

    let response: ClaudeResponse;
    try {
      response = await runClaudeSession({
        prompt,
        cwd: session.cwd,
        resumeSessionId: session.sessionId ?? undefined,
        onProgress,
        signal,
      });
    } catch (err) {
      // If resume failed, clear sessionId so next attempt starts fresh
      if (session.sessionId) {
        console.log(`[session] Clearing stale session ${session.sessionId.slice(0, 8)} for thread ${threadId}`);
        session.sessionId = null;
        this.persist();
      }
      throw err;
    }

    // Update session ID for future resumption
    if (response.sessionId) {
      session.sessionId = response.sessionId;
    }

    session.status = "idle";
    this.persist();

    // Handle handoff if requested
    if (response.handoff) {
      await this.performHandoff(threadId, response.handoff, response.text);
    }

    return response;
  }

  private async performHandoff(threadId: string, handoff: HandoffRequest, lastResponse: string): Promise<void> {
    const session = this.sessions.get(threadId);
    if (!session) return;

    console.log(`[session] Handoff for thread ${threadId}: ${handoff.type} → ${handoff.target}`);

    // Capture conversation context for the new environment
    session.handoffContext = [
      `The user was previously working in a different environment and has now switched to local directory "${handoff.target}".`,
      `Last assistant response before the switch:`,
      lastResponse.slice(0, 2000),
    ].join("\n");

    // Reset session for new environment
    session.sessionId = null;
    session.status = "handoff";
    session.cwd = handoff.target;

    console.log(`[session] Handoff complete: cwd=${session.cwd}`);
    this.persist();
  }

  destroySession(threadId: string): boolean {
    const result = this.sessions.delete(threadId);
    this.persist();
    return result;
  }
}
