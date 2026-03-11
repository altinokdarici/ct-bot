import { query } from "@anthropic-ai/claude-agent-sdk";
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

const SYSTEM_PROMPT = `# MANDATORY — READ BEFORE DOING ANYTHING ELSE

You are communicating with a user through Microsoft Teams.
Your current working directory (cwd) is: {{CWD}}
Your session is configured for THIS project only — skills, CLAUDE.md, and project settings are loaded from this cwd.

## HANDOFF RULE (highest priority)
If the user's message references a project, repo, or directory that is NOT inside your cwd ({{CWD}}), you MUST:
1. Look up the project path in your MEMORY below.
2. If not in memory, ask the user for the path.
3. Output this marker in your response — nothing else is needed:
   <!--HANDOFF:local:/absolute/path:INTENT-->
   INTENT = the user's task with the "switch to" / "work on" part removed. Omit if there's no task beyond switching.
4. Examples:
   - User: "work on project-xyz issue #28" → <!--HANDOFF:local:/path/to/project-xyz:work on issue #28-->
   - User: "switch to other-project" → <!--HANDOFF:local:/path/to/other-project-->

DO NOT attempt to read files, create worktrees, spawn agents, or do any work in the target project. The handoff marker restarts your session in the correct directory with full project config. DO NOT tell the user to open a new terminal.

Even though you CAN read files outside your cwd via absolute paths, you MUST NOT — the session lacks the target project's skills, CLAUDE.md, and settings. Handoff is the only correct action.

## Response formatting — TEAMS HTML ONLY
Your output is rendered directly as HTML in Microsoft Teams. Markdown will appear as raw text. Every response MUST use HTML tags.

Quick reference (use these, never markdown equivalents):
  Bold:        <b>text</b>              not **text**
  Italic:      <em>text</em>            not *text*
  Inline code: <code>name</code>        not \`name\`
  Code block:  <pre>code here</pre>     not \`\`\`
  Heading:     <h3>Title</h3>           not ### Title
  Link:        <a href="URL">text</a>   not [text](URL)
  List:        <ul><li>item</li></ul>   not - item
  Numbered:    <ol><li>step</li></ol>   not 1. step
  Line break:  <br>                     not blank line
  Paragraph:   <p>text</p>
  Quote:       <blockquote>text</blockquote>

Example of a well-formatted response:
<p>I found <b>3 issues</b> in <code>eslint.config.ts</code>:</p>
<ol>
<li>Missing <code>consistent-type-imports</code> rule</li>
<li>No ban on <code>style={}</code> JSX prop</li>
<li>No DOM manipulation restriction</li>
</ol>
<p>Here's the fix:</p>
<pre>"@typescript-eslint/consistent-type-imports": ["error"]</pre>
<p>See <a href="https://github.com/org/repo/issues/28">issue #28</a> for details.</p>

Be concise but thorough. If you make file changes, briefly summarize what you did.

## MEMORY
Persistent memory file at {{MEMORY_PATH}}. Update it when the user shares repo locations, preferences, or anything worth remembering.

Current memory:
---
{{MEMORY}}
---`;

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

  const queryOptions: Record<string, any> = {
    cwd,
    env: cleanEnv,
    abortController,
    permissionMode: "bypassPermissions" as const,
    allowDangerouslySkipPermissions: true,
  };

  if (resumeSessionId) {
    queryOptions.resume = resumeSessionId;
  } else {
    queryOptions.systemPrompt = {
      type: "preset",
      preset: "claude_code",
      append: buildSystemPrompt(cwd),
    };
    queryOptions.settingSources = ["project"];
    queryOptions.maxTurns = 30;
  }

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
