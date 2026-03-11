import { startOrchestrator } from "./orchestrator.ts";

// Remove CLAUDECODE env var so spawned Claude Code sessions don't think they're nested
delete process.env.CLAUDECODE;

console.log("=== Teams ↔ Claude Code Bridge ===");
startOrchestrator().catch((err) => {
  console.error("Fatal error:", err);
  process.exit(1);
});
