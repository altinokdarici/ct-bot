import { startOrchestrator } from "./orchestrator.ts";

// Remove CLAUDECODE env var so spawned Claude Code sessions don't think they're nested
delete process.env.CLAUDECODE;

console.log("=== Teams ↔ Claude Code Bridge ===");
console.warn("\x1b[33m⚠  WARNING: Runs with --dangerously-skip-permissions using your Azure credentials.\x1b[0m");
console.warn("\x1b[33m   Claude Code has full filesystem/tool access with no prompts. Use in trusted environments only.\x1b[0m");
startOrchestrator().catch((err) => {
  console.error("Fatal error:", err);
  process.exit(1);
});
