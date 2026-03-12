import { query, type Options } from "@anthropic-ai/claude-agent-sdk";
import { readFileSync, writeFileSync, existsSync } from "node:fs";
import { join } from "node:path";
import type { HandoffRequest } from "./types.ts";

const HANDOFF_REGEX = /<!--HANDOFF:local:(.+?)(?::(.+?))?-->/;

const MEMORY_PATH = join(import.meta.dirname ?? process.cwd(), "..", "memory.md");

function loadMemory(): string {
  if (existsSync(MEMORY_PATH)) {
    return readFileSync(MEMORY_PATH, "utf8").trim();
  }
  writeFileSync(MEMORY_PATH, "# Memory\n");
  return "# Memory";
}

const SYSTEM_PROMPT = `<CRITICAL-RULES>
<RULE name="HTML-ONLY" priority="HIGHEST">
ALL your responses are rendered in Microsoft Teams which ONLY supports HTML.
If you use markdown (**, *, \`, \`\`\`, ###, - item, [text](url), etc.) it will show as UGLY RAW TEXT to the user.

YOU MUST USE HTML TAGS FOR EVERYTHING. ZERO MARKDOWN. NO EXCEPTIONS.

  <b>bold</b>           NEVER **bold**
  <em>italic</em>       NEVER *italic*
  <code>x</code>        NEVER \`x\`
  <pre>code</pre>       NEVER \`\`\`code\`\`\`
  <h3>Title</h3>        NEVER ### Title
  <a href="U">t</a>     NEVER [t](U)
  <ul><li>item</li></ul> NEVER - item
  <ol><li>step</li></ol> NEVER 1. step
  <br>                   NEVER blank lines for spacing
  <p>text</p>            wrap every paragraph

Example response:
<p>I found <b>3 issues</b> in <code>eslint.config.ts</code>:</p>
<ol>
<li>Missing <code>consistent-type-imports</code> rule</li>
<li>No ban on <code>style={{}}</code> JSX prop</li>
</ol>
<p>Here is the fix:</p>
<pre>"@typescript-eslint/consistent-type-imports": ["error"]</pre>

Before finishing every response, re-check: does my output contain ANY markdown syntax? If yes, convert it to HTML.
</RULE>

<RULE name="HANDOFF" priority="HIGH">
Your cwd is: {{CWD}}. This session is scoped to THIS project only.

If the user references a project/repo/directory NOT inside {{CWD}}:
1. Look up the path in MEMORY below.
2. If not in memory, ask the user.
3. Emit ONLY: &lt;!--HANDOFF:local:/absolute/path:INTENT--&gt;
   (INTENT = the task, omit if just switching)
4. Do NO work in the other project. The marker restarts your session there.
Do NOT read files outside cwd. Do NOT tell the user to open a terminal.
</RULE>

<RULE name="WORKIQ" priority="HIGH">
You have WorkIQ MCP tools for querying the user's Microsoft 365 data (emails, calendar, documents, Teams messages, people).

For ANY work-related question (meetings, emails, files, colleagues, schedules), use these tools:
<ul>
<li><code>mcp__workiq__accept_eula</code> — Call ONCE if <code>ask_work_iq</code> fails with a EULA error, then retry.</li>
<li><code>mcp__workiq__ask_work_iq</code> — Main tool. Pass the question in natural language.</li>
<li><code>mcp__workiq__get_debug_link</code> — Get a diagnostic link if a query fails.</li>
</ul>
Prefer WorkIQ over web search for anything about the user's work context.
</RULE>

<RULE name="IMAGES" priority="HIGH">
To include an image or screenshot in your response, save it to a file and emit:
&lt;!--IMAGE:/absolute/path/to/file.png--&gt;

The system will upload it to Teams and display it inline. You can include multiple image markers.
Use this for screenshots, generated diagrams, diff images, or any visual content.
</RULE>
</CRITICAL-RULES>

<MEMORY>
Persistent memory file at {{MEMORY_PATH}}. Update it when the user shares repo locations, preferences, or anything worth remembering.
Current contents:
{{MEMORY}}
</MEMORY>`;

function buildSystemPrompt(cwd: string): string {
  const memory = loadMemory();
  return SYSTEM_PROMPT
    .replaceAll("{{CWD}}", cwd)
    .replace("{{MEMORY_PATH}}", MEMORY_PATH)
    .replace("{{MEMORY}}", memory);
}

// ---------------------------------------------------------------------------
// Progress events emitted during a session
// ---------------------------------------------------------------------------

export type ProgressEvent =
  | { type: "init"; sessionId: string; model: string }
  | { type: "thinking"; text: string }
  | { type: "tool_use"; tool: string; input: string }
  | { type: "tool_result"; summary: string }
  | { type: "text"; text: string }
  | { type: "result"; text: string; costUsd: number; durationMs: number; turns: number };

export type ProgressCallback = (event: ProgressEvent) => void | Promise<void>;

export interface ClaudeResponse {
  text: string;
  sessionId: string;
  handoff: HandoffRequest | null;
}

function summarizeToolInput(tool: string, input: any): string {
  if (!input || typeof input !== "object") return "";
  switch (tool) {
    case "Bash": return input.command?.slice(0, 120) ?? "";
    case "Read": return input.file_path ?? "";
    case "Write": return input.file_path ?? "";
    case "Edit": return input.file_path ?? "";
    case "Glob": return input.pattern ?? "";
    case "Grep": return `${input.pattern ?? ""} ${input.path ?? ""}`.trim();
    case "WebSearch": return input.query ?? "";
    case "WebFetch": return input.url ?? "";
    case "mcp__workiq__ask_work_iq": return input.question ?? input.query ?? JSON.stringify(input).slice(0, 120);
    case "mcp__workiq__accept_eula": return "accepting EULA";
    case "mcp__workiq__get_debug_link": return "fetching debug link";
    default: return JSON.stringify(input).slice(0, 100);
  }
}

export async function runClaudeSession(opts: {
  prompt: string;
  cwd: string;
  resumeSessionId?: string;
  onProgress?: ProgressCallback;
  signal?: AbortSignal;
}): Promise<ClaudeResponse> {
  const { prompt, cwd, resumeSessionId, onProgress, signal } = opts;
  const emit = onProgress ?? (() => {});

  // Strip CLAUDECODE env var so the child process doesn't inherit it
  const cleanEnv = { ...process.env };
  delete cleanEnv.CLAUDECODE;

  // Build abort controller that merges our external signal
  const abortController = new AbortController();
  if (signal) {
    signal.addEventListener("abort", () => abortController.abort(), { once: true });
  }

  const queryOptions: Options = {
    cwd,
    env: cleanEnv,
    abortController,
    permissionMode: "bypassPermissions",
    allowDangerouslySkipPermissions: true,
    stderr: (data: string) => { process.stderr.write(`[claude-stderr] ${data}`); },
    settingSources: ["user", "project"],
    maxTurns: 30,
    systemPrompt: {
      type: "preset",
      preset: "claude_code",
      append: buildSystemPrompt(cwd),
    },
    mcpServers: {
      workiq: {
        type: "stdio",
        command: "npx",
        args: ["-y", "@microsoft/workiq", "mcp"],
      },
    },
    ...(resumeSessionId ? { resume: resumeSessionId } : {}),
  };

  let sessionId: string = resumeSessionId ?? "";
  let resultText: string = "";

  for await (const message of query({ prompt, options: queryOptions })) {
    const msg = message as any;
    const msgJson = JSON.stringify(msg);
    console.log(`[sdk] ${msg.type}${msg.subtype ? `:${msg.subtype}` : ""} ${msgJson.slice(0, 300)}${msgJson.length > 300 ? "..." : ""}`);

    switch (msg.type) {
      case "system": {
        if (msg.subtype === "init") {
          sessionId = msg.session_id;
          const mcpServers = msg.mcp_servers ?? [];
          console.log(`[sdk] MCP servers: ${JSON.stringify(mcpServers)}`);
          emit({ type: "init", sessionId, model: msg.model ?? "unknown" });
        }
        break;
      }

      case "assistant": {
        const content = msg.message?.content;
        if (Array.isArray(content)) {
          for (const block of content) {
            if (block.type === "thinking" && block.thinking) {
              emit({ type: "thinking", text: block.thinking });
            } else if (block.type === "text" && block.text) {
              emit({ type: "text", text: block.text });
            } else if (block.type === "tool_use") {
              emit({
                type: "tool_use",
                tool: block.name,
                input: summarizeToolInput(block.name, block.input),
              });
            }
          }
        }
        break;
      }

      case "tool_use_summary": {
        emit({ type: "tool_result", summary: msg.summary ?? "" });
        break;
      }

      case "result": {
        if (msg.subtype === "error_during_execution" || msg.is_error) {
          console.error(`[sdk] ERROR: ${JSON.stringify(msg)}`);
        }
        resultText = msg.result ?? "";
        // Capture session ID from error responses too (SDK may assign a new one)
        if (msg.session_id) sessionId = msg.session_id;
        await emit({
          type: "result",
          text: resultText,
          costUsd: msg.total_cost_usd ?? 0,
          durationMs: msg.duration_ms ?? 0,
          turns: msg.num_turns ?? 0,
        });
        break;
      }
    }
  }

  // Check for handoff signal
  let handoff: HandoffRequest | null = null;
  const handoffMatch = resultText.match(HANDOFF_REGEX);
  if (handoffMatch) {
    handoff = {
      type: "local",
      target: handoffMatch[1]!,
      intent: handoffMatch[2] || null,
    };
    resultText = resultText.replace(HANDOFF_REGEX, "").trim();
  }

  return { text: resultText, sessionId, handoff };
}
