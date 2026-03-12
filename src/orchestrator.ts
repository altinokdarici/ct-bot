import { startListener, sendThreadReply, editThreadReply, uploadImageToTeams, resolveImages } from "./teams-bridge.ts";
import { SessionManager } from "./session-manager.ts";
import { auditLog } from "./audit-log.ts";
import { escapeHtml } from "./escape-html.ts";
import type { TeamsMessage } from "./types.ts";
import type { ProgressEvent } from "./claude-session.ts";

function requireEnv(name: string): string {
  const value = process.env[name];
  if (!value) throw new Error(`${name} env var is required`);
  return value;
}

const CHANNEL_ID = requireEnv("CHANNEL_ID");

const MAX_MESSAGE_LENGTH = 20_000;

// Throttle Teams edits to avoid rate limits
const EDIT_THROTTLE_MS = 2000;

function formatResponse(text: string): string {
  if (/<[a-z][\s\S]*>/i.test(text)) return text;
  return escapeHtml(text).replace(/\n/g, "<br>");
}

const IMAGE_MARKER_REGEX = /<!--IMAGE:(.+?)-->/g;

async function processImageMarkers(html: string, channelId: string): Promise<string> {
  const matches = [...html.matchAll(IMAGE_MARKER_REGEX)];
  if (matches.length === 0) return html;

  let result = html;
  for (const match of matches) {
    const filePath = match[1]!;
    try {
      console.log(`[orchestrator] Uploading image: ${filePath}`);
      const imgTag = await uploadImageToTeams(channelId, filePath);
      result = result.replace(match[0], imgTag);
    } catch (err: any) {
      console.error(`[orchestrator] Image upload failed for ${filePath}:`, err.message);
      result = result.replace(match[0], `<em>(image upload failed: ${escapeHtml(filePath)})</em>`);
    }
  }
  return result;
}

function splitMessage(text: string): string[] {
  if (text.length <= MAX_MESSAGE_LENGTH) return [text];
  const chunks: string[] = [];
  let remaining = text;
  while (remaining.length > 0) {
    if (remaining.length <= MAX_MESSAGE_LENGTH) {
      chunks.push(remaining);
      break;
    }
    let splitAt = remaining.lastIndexOf("\n", MAX_MESSAGE_LENGTH);
    if (splitAt < MAX_MESSAGE_LENGTH / 2) splitAt = MAX_MESSAGE_LENGTH;
    chunks.push(remaining.slice(0, splitAt));
    remaining = remaining.slice(splitAt);
  }
  return chunks;
}

// ---------------------------------------------------------------------------
// Console logging helpers
// ---------------------------------------------------------------------------

const DIM = "\x1b[2m";
const RESET = "\x1b[0m";
const BOLD = "\x1b[1m";
const CYAN = "\x1b[36m";
const YELLOW = "\x1b[33m";
const GREEN = "\x1b[32m";
const MAGENTA = "\x1b[35m";
const BLUE = "\x1b[34m";

function logThinking(threadId: string, text: string) {
  const preview = text.replace(/\n/g, " ").slice(0, 120);
  console.log(`${DIM}[${threadId}] ${MAGENTA}thinking:${RESET}${DIM} ${preview}${text.length > 120 ? "..." : ""}${RESET}`);
}

function logToolUse(threadId: string, tool: string, input: string) {
  console.log(`${DIM}[${threadId}]${RESET} ${YELLOW}tool:${RESET} ${BOLD}${tool}${RESET} ${DIM}${input.slice(0, 100)}${input.length > 100 ? "..." : ""}${RESET}`);
}

function logToolResult(threadId: string, summary: string) {
  const preview = summary.replace(/\n/g, " ").slice(0, 120);
  console.log(`${DIM}[${threadId}]${RESET} ${CYAN}result:${RESET} ${DIM}${preview}${summary.length > 120 ? "..." : ""}${RESET}`);
}

function logText(threadId: string, text: string) {
  const preview = text.replace(/\n/g, " ").slice(0, 150);
  console.log(`${DIM}[${threadId}]${RESET} ${GREEN}text:${RESET} ${preview}${text.length > 150 ? "..." : ""}`);
}

function logResult(threadId: string, costUsd: number, durationMs: number, turns: number) {
  const secs = (durationMs / 1000).toFixed(1);
  const cost = costUsd.toFixed(4);
  console.log(`${BOLD}[${threadId}]${RESET} ${GREEN}done${RESET} ${DIM}${turns} turns, ${secs}s, $${cost}${RESET}`);
}

// ---------------------------------------------------------------------------
// Live Teams progress message
// ---------------------------------------------------------------------------

class TeamsProgressUpdater {
  private channelId: string;
  private threadId: string;
  private messageId: string | null = null;
  private status = "🧠 Thinking...";
  private toolSteps: string[] = [];
  private textPreview: string | null = null;
  private lastEditTime = 0;
  private editTimer: ReturnType<typeof setTimeout> | null = null;
  private finished = false;
  private inflightEdit: Promise<void> = Promise.resolve();

  constructor(channelId: string, threadId: string) {
    this.channelId = channelId;
    this.threadId = threadId;
  }

  async sendInitial(): Promise<void> {
    this.messageId = await sendThreadReply(
      this.channelId,
      this.threadId,
      `<em>🧠 Thinking...</em>`,
    );
  }

  setStatus(text: string): void {
    this.status = text;
    this.scheduleEdit();
  }

  addToolStep(tool: string, input: string): void {
    this.toolSteps.push(`🔧 <b>${tool}</b> <code>${input}</code>`);
    this.status = `Running ${tool}...`;
    this.textPreview = null; // clear text preview when a new tool starts
    this.scheduleEdit();
  }

  addTextPreview(text: string): void {
    this.textPreview = text;
    this.status = "⏳ Working...";
    this.scheduleEdit();
  }

  private scheduleEdit(): void {
    if (this.finished || !this.messageId) return;
    const now = Date.now();
    const elapsed = now - this.lastEditTime;
    if (elapsed >= EDIT_THROTTLE_MS) {
      this.doEdit();
    } else if (!this.editTimer) {
      this.editTimer = setTimeout(() => {
        this.editTimer = null;
        this.doEdit();
      }, EDIT_THROTTLE_MS - elapsed);
    }
  }

  private doEdit(): void {
    if (this.finished || !this.messageId) return;
    this.lastEditTime = Date.now();
    const html = this.buildProgressHtml();
    this.inflightEdit = editThreadReply(this.channelId, this.threadId, this.messageId, html).catch((e) => {
      console.error(`[progress] Edit error:`, e.message);
    });
  }

  private buildProgressHtml(): string {
    let html = "";
    if (this.toolSteps.length > 0) {
      const maxVisible = 6;
      const visible = this.toolSteps.slice(-maxVisible);
      const hidden = this.toolSteps.length - visible.length;
      if (hidden > 0) html += `<em>...${hidden} more step(s)</em><br>`;
      html += visible.join("<br>") + "<br><br>";
    }
    if (this.textPreview) {
      html += `<p>${this.textPreview}</p><br>`;
    }
    html += `<em>${this.status}</em>`;
    return html;
  }

  /** Replace the progress message entirely with the final response */
  async finalize(responseHtml: string): Promise<void> {
    this.finished = true;
    if (this.editTimer) { clearTimeout(this.editTimer); this.editTimer = null; }
    if (!this.messageId) return;
    // Wait for any in-flight progress edit to complete before sending final
    await this.inflightEdit;
    console.log(`[progress] Finalizing message ${this.messageId} (${responseHtml.length} chars)`);
    await editThreadReply(this.channelId, this.threadId, this.messageId, responseHtml);
  }

  async finalizeError(errorHtml: string): Promise<void> {
    this.finished = true;
    if (this.editTimer) { clearTimeout(this.editTimer); this.editTimer = null; }
    await this.inflightEdit;
    if (this.messageId) {
      await editThreadReply(this.channelId, this.threadId, this.messageId, errorHtml);
    } else {
      await sendThreadReply(this.channelId, this.threadId, errorHtml);
    }
  }
}

// ---------------------------------------------------------------------------
// Orchestrator
// ---------------------------------------------------------------------------

const CANCEL_KEYWORDS = /^(cancel|stop|abort|nevermind)$/i;

export async function startOrchestrator(): Promise<void> {
  const sessionManager = new SessionManager();
  const inFlight = new Map<string, { abort: AbortController; promise: Promise<void>; progress: TeamsProgressUpdater }>();
  const messageQueues = new Map<string, TeamsMessage[]>();

  async function processMessage(msg: TeamsMessage): Promise<void> {
    const { threadId, text } = msg;
    const abortController = new AbortController();
    let progress = new TeamsProgressUpdater(CHANNEL_ID, threadId);

    const promise = (async () => {
    try {
      let currentPrompt = await resolveImages(text);
      let handoffsRemaining = 3;

      // Loop: after a handoff with intent, re-send the extracted intent to the new environment
      while (true) {
      await progress.sendInitial();

      let finalized = false;

      const response = await sessionManager.sendMessage(threadId, currentPrompt, async (event: ProgressEvent) => {
        // If aborted, stop emitting progress
        if (abortController.signal.aborted) return;
        switch (event.type) {
          case "init":
            console.log(`${DIM}[${threadId}]${RESET} ${GREEN}init${RESET} model=${event.model} session=${event.sessionId.slice(0, 8)}`);
            auditLog(threadId, { direction: "outgoing", type: "init", text: "", sessionId: event.sessionId, meta: { model: event.model } });
            break;

          case "thinking":
            logThinking(threadId, event.text);
            auditLog(threadId, { direction: "outgoing", type: "thinking", text: event.text });
            progress.setStatus("🧠 Thinking...");
            break;

          case "tool_use":
            logToolUse(threadId, event.tool, event.input);
            auditLog(threadId, { direction: "outgoing", type: "tool_use", text: event.input, meta: { tool: event.tool } });
            progress.addToolStep(escapeHtml(event.tool), escapeHtml(event.input.slice(0, 80)));
            break;

          case "tool_result":
            logToolResult(threadId, event.summary);
            auditLog(threadId, { direction: "outgoing", type: "tool_result", text: event.summary });
            break;

          case "text":
            logText(threadId, event.text);
            auditLog(threadId, { direction: "outgoing", type: "text", text: event.text });
            progress.addTextPreview(formatResponse(event.text));
            break;

          case "result": {
            logResult(threadId, event.costUsd, event.durationMs, event.turns);
            auditLog(threadId, { direction: "outgoing", type: "result", text: event.text, meta: { costUsd: event.costUsd, durationMs: event.durationMs, turns: event.turns } });
            // Finalize immediately from the result event
            finalized = true;
            let resultHtml = event.text ? formatResponse(event.text) : "<p><em>(no response)</em></p>";
            resultHtml = await processImageMarkers(resultHtml, CHANNEL_ID);
            await progress.finalize(resultHtml);

            // Send overflow chunks if needed
            if (event.text) {
              const chunks = splitMessage(resultHtml);
              for (let i = 1; i < chunks.length; i++) {
                await sendThreadReply(CHANNEL_ID, threadId, chunks[i]!);
              }
            }
            break;
          }
        }
      }, abortController.signal);

      // Fallback: finalize after loop if result event didn't trigger it
      if (!finalized && !abortController.signal.aborted) {
        const formatted = response.text ? formatResponse(response.text) : "<p><em>(no response)</em></p>";
        await progress.finalize(formatted);
      }

      if (response.handoff) {
        const target = `directory <b>${escapeHtml(response.handoff.target)}</b>`;

        if (response.handoff.intent && --handoffsRemaining >= 0) {
          // Intent extracted — re-send it to the new environment automatically
          await sendThreadReply(
            CHANNEL_ID,
            threadId,
            `<p><em>Switched to ${target}. Running: ${escapeHtml(response.handoff.intent)}</em></p>`,
          );
          console.log(`${BOLD}[orchestrator]${RESET} Handoff with intent for ${threadId}: "${response.handoff.intent}"`);
          auditLog(threadId, { direction: "outgoing", type: "handoff", text: response.handoff.intent, meta: { target: response.handoff.target } });
          currentPrompt = response.handoff.intent;
          progress = new TeamsProgressUpdater(CHANNEL_ID, threadId);
          continue;
        }

        // No intent — just notify the user
        await sendThreadReply(
          CHANNEL_ID,
          threadId,
          `<p><em>Switched to ${target}. Next messages will run in the new environment.</em></p>`,
        );
      }

      break;
      } // end while
    } catch (err: any) {
      if (abortController.signal.aborted) {
        console.log(`${DIM}[${threadId}]${RESET} ${YELLOW}cancelled${RESET}`);
        auditLog(threadId, { direction: "outgoing", type: "cancelled", text: "" });
        await progress.finalizeError(`<p><em>Cancelled.</em></p>`);
      } else {
        console.error(`${BOLD}[${threadId}]${RESET} \x1b[31merror:${RESET}`, err.message);
        auditLog(threadId, { direction: "outgoing", type: "error", text: err.message ?? "Unknown error" });
        await progress.finalizeError(
          `<p><b>Error:</b> ${escapeHtml(err.message ?? "Unknown error")}</p>`,
        );
      }
    }
    })();

    inFlight.set(threadId, { abort: abortController, promise, progress });
    await promise;
    inFlight.delete(threadId);

    // Process queued messages for this thread
    const queue = messageQueues.get(threadId);
    if (queue && queue.length > 0) {
      const next = queue.shift()!;
      if (queue.length === 0) messageQueues.delete(threadId);
      console.log(`${DIM}[orchestrator] Processing queued message for ${threadId}${RESET}`);
      await processMessage(next);
    }
  }

  async function handleMessage(msg: TeamsMessage): Promise<void> {
    const { threadId, from, text } = msg;

    console.log(`\n${BOLD}[orchestrator]${RESET} ${msg.isNewThread ? "NEW" : "REPLY"} ${BLUE}[${threadId}]${RESET} ${from}: ${text.slice(0, 100)}`);
    auditLog(threadId, { direction: "incoming", from, type: "user_message", text, meta: { messageId: msg.messageId, isNewThread: msg.isNewThread } });

    const running = inFlight.get(threadId);

    if (running) {
      // Cancel command: abort the current session
      if (CANCEL_KEYWORDS.test(text.trim())) {
        console.log(`${YELLOW}[orchestrator] Cancelling thread ${threadId}${RESET}`);
        running.abort.abort();
        return;
      }

      // Queue the message
      console.log(`${DIM}[orchestrator] Queueing message for busy thread ${threadId}${RESET}`);
      let queue = messageQueues.get(threadId);
      if (!queue) {
        queue = [];
        messageQueues.set(threadId, queue);
      }
      queue.push(msg);
      await sendThreadReply(
        CHANNEL_ID,
        threadId,
        `<p><em>Queued — I'll process this after the current task finishes. Send <b>cancel</b> to abort the current task.</em></p>`,
      );
      return;
    }

    // Don't await — let it run in background so the listener stays free for other threads
    processMessage(msg).catch((err) => {
      console.error(`${BOLD}[${threadId}]${RESET} \x1b[31munhandled:${RESET}`, err.message);
    });
  }

  console.log(`${BOLD}[orchestrator]${RESET} Channel: ${CHANNEL_ID}`);
  console.log(`${BOLD}[orchestrator]${RESET} Starting Teams listener...`);

  await startListener(CHANNEL_ID, handleMessage);
}
