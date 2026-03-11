import { execSync } from "node:child_process";
import { existsSync, readFileSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { randomBytes } from "node:crypto";
import { gunzipSync } from "node:zlib";
import { randomUUID } from "node:crypto";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const BOLD = "\x1b[1m";
const GREEN = "\x1b[32m";
const RED = "\x1b[31m";
const YELLOW = "\x1b[33m";
const CYAN = "\x1b[36m";
const DIM = "\x1b[2m";
const RESET = "\x1b[0m";

function ok(msg: string) { console.log(`  ${GREEN}✓${RESET} ${msg}`); }
function fail(msg: string) { console.log(`  ${RED}✗${RESET} ${msg}`); }
function warn(msg: string) { console.log(`  ${YELLOW}!${RESET} ${msg}`); }
function info(msg: string) { console.log(`  ${DIM}${msg}${RESET}`); }
function heading(msg: string) { console.log(`\n${BOLD}${CYAN}${msg}${RESET}\n`); }

function commandExists(cmd: string): boolean {
  try {
    const check = process.platform === "win32" ? `where ${cmd}` : `which ${cmd}`;
    execSync(check, { stdio: "pipe" });
    return true;
  } catch {
    return false;
  }
}

function prompt(question: string): Promise<string> {
  return new Promise((resolve) => {
    process.stdout.write(`  ${question} `);
    let data = "";
    process.stdin.setEncoding("utf8");
    process.stdin.resume();
    process.stdin.once("data", (chunk: string) => {
      process.stdin.pause();
      data = chunk.trim();
      resolve(data);
    });
  });
}

// ---------------------------------------------------------------------------
// Auth helpers (duplicated from teams-bridge to keep setup standalone)
// ---------------------------------------------------------------------------

function listTenantIds(): string[] {
  const out = execSync('az account list --all --query "[].tenantId" -o tsv', {
    encoding: "utf8",
  });
  return [...new Set(out.split("\n").map((t: string) => t.trim()).filter(Boolean))];
}

function azToken(tenantId: string, resource: string): string | null {
  try {
    return execSync(
      `az account get-access-token --tenant ${tenantId} --resource ${resource} --query accessToken -o tsv`,
      { encoding: "utf8", stdio: ["pipe", "pipe", "pipe"] },
    ).trim();
  } catch {
    return null;
  }
}

function getAadToken(): { token: string; tenantId: string } {
  for (const tenantId of listTenantIds()) {
    const token = azToken(tenantId, "https://api.spaces.skype.com");
    if (token) return { token, tenantId };
  }
  throw new Error("Could not get AAD token");
}

function decodeToken(token: string): Record<string, any> {
  return JSON.parse(Buffer.from(token.split(".")[1]!, "base64url").toString());
}

async function getSkypeToken(aadToken: string): Promise<string> {
  const resp = await fetch("https://teams.microsoft.com/api/authsvc/v1.0/authz", {
    method: "POST",
    headers: { Authorization: `Bearer ${aadToken}`, "Content-Type": "application/json" },
    body: "{}",
  });
  if (!resp.ok) throw new Error(`authz failed: ${resp.status}`);
  const data: any = await resp.json();
  const skypeToken = data.tokens?.skypeToken;
  if (!skypeToken) throw new Error("No skypeToken in authz response");
  return skypeToken;
}

// ---------------------------------------------------------------------------
// Trouter (channel discovery)
// ---------------------------------------------------------------------------

const TROUTER_GO = "https://go.trouter.teams.microsoft.com";

async function setupTrouter(skypeToken: string, epid: string) {
  const resp = await fetch(`${TROUTER_GO}/v4/a?epid=${encodeURIComponent(epid)}`, {
    method: "POST",
    headers: { "X-Skypetoken": skypeToken, "Content-Length": "0" },
  });
  if (!resp.ok) throw new Error(`Trouter session failed: ${resp.status}`);
  const data: any = await resp.json();
  return {
    socketio: (data.socketio || `${TROUTER_GO}/`) as string,
    surl: data.surl as string,
    connectparams: data.connectparams as Record<string, string>,
    ccid: data.ccid as string,
  };
}

function buildConnectQuery(connectparams: Record<string, string>, ccid: string, epid: string): string {
  const parts = ["v=v4"];
  for (const [key, value] of Object.entries(connectparams)) {
    parts.push(`${key}=${encodeURIComponent(value)}`);
  }
  const tc = JSON.stringify({ cv: "2024.23.01.2", ua: "TeamsCDL", hr: "", v: "49/25010202142" });
  parts.push(`tc=${encodeURIComponent(tc)}`);
  parts.push(`con_num=${Date.now()}_1`);
  parts.push(`epid=${encodeURIComponent(epid)}`);
  if (ccid) parts.push(`ccid=${encodeURIComponent(ccid)}`);
  parts.push("auth=true");
  parts.push("timeout=40");
  return parts.join("&");
}

async function register(aadToken: string, skypeToken: string, surl: string, epid: string): Promise<void> {
  const url = "https://teams.microsoft.com/registrar/prod/V2/registrations";
  const registrations = [
    { appId: "SkypeSpacesWeb", templateKey: "SkypeSpacesWeb_2.3", path: `${surl}SkypeSpacesWeb` },
    { appId: "TeamsCDLWebWorker", templateKey: "TeamsCDLWebWorker_2.1", path: surl },
  ];
  for (const reg of registrations) {
    const body = {
      clientDescription: {
        appId: reg.appId, aesKey: "", languageId: "en-US", platform: "edge",
        templateKey: reg.templateKey, platformUIVersion: "49/25010202142",
      },
      registrationId: epid, nodeId: "",
      transports: { TROUTER: [{ context: "", path: reg.path, ttl: 86400 }] },
    };
    await fetch(url, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${aadToken}`,
        "X-Skypetoken": skypeToken,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(body),
    });
  }
}

function stripHtml(s: string): string {
  return s
    .replace(/<[^>]+>/g, "")
    .replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">")
    .replace(/&nbsp;/g, " ").replace(/&#39;/g, "'").replace(/&quot;/g, '"')
    .trim();
}

function parseNotification(raw: any): any {
  try {
    const data = typeof raw === "string" ? JSON.parse(raw) : raw;
    let body = data.body ?? data;
    if (typeof body === "string") {
      try { body = JSON.parse(body); } catch { /* leave */ }
    }
    if (body && typeof body === "object" && body.cp) {
      try {
        body = JSON.parse(gunzipSync(Buffer.from(body.cp, "base64")).toString());
      } catch { /* ignore */ }
    }
    if (typeof body === "string") {
      const encoding = data.headers?.["X-Microsoft-Skype-Content-Encoding"] ??
                        data.headers?.["x-microsoft-skype-content-encoding"];
      if (encoding === "gzip") {
        body = JSON.parse(gunzipSync(Buffer.from(body, "base64")).toString());
      }
    }
    return body;
  } catch {
    return null;
  }
}

function extractChannelFromNotification(body: any): { channelId: string; text: string } | null {
  const resource = body.resource ?? body;
  const msgType = resource.messagetype ?? resource.messageType;
  if (!msgType || !["RichText/Html", "Text", "RichText"].includes(msgType)) return null;

  const html = resource.content ?? resource.body ?? "";
  const text = msgType.startsWith("RichText") ? stripHtml(html) : html.trim();
  if (!text) return null;

  const channelId = resource.to ?? "";
  if (!channelId) return null;

  return { channelId, text };
}

/**
 * Listen to all Teams notifications and resolve when a message matching
 * the secret code is found. Returns the channel ID.
 */
function discoverChannel(
  aadToken: string,
  skypeToken: string,
  secretCode: string,
): Promise<string> {
  return new Promise(async (resolve, reject) => {
    const epid = randomUUID();
    const trouter = await setupTrouter(skypeToken, epid);

    const query = buildConnectQuery(trouter.connectparams, trouter.ccid, epid);
    const hsResp = await fetch(`${trouter.socketio}socket.io/1/?${query}`, {
      headers: { "X-Skypetoken": skypeToken },
    });
    if (!hsResp.ok) { reject(new Error(`Handshake failed: ${hsResp.status}`)); return; }
    const sessionId = (await hsResp.text()).split(":")[0]!;

    const wsBase = trouter.socketio.replace(/^http/, "ws");
    const wsUrl = `${wsBase}socket.io/1/websocket/${sessionId}?${query}`;
    const ws = new WebSocket(wsUrl, { headers: { "X-Skypetoken": skypeToken } } as any);

    let heartbeat: ReturnType<typeof setInterval> | null = null;
    let cmdCount = 0;
    const timeout = setTimeout(() => {
      ws.close();
      reject(new Error("Timed out waiting for message (120s). Make sure you posted the code in a Teams channel."));
    }, 120_000);

    const sendEvent = (json: any) => ws.send(`5:${cmdCount++}+::${JSON.stringify(json)}`);
    const sendEphemeral = (json: any) => ws.send(`5:::${JSON.stringify(json)}`);

    function handleNotification(data: any) {
      const items = Array.isArray(data) ? data : [data];
      for (const item of items) {
        const body = parseNotification(item);
        if (!body) continue;
        const msg = extractChannelFromNotification(body);
        if (!msg) continue;
        if (msg.text.includes(secretCode)) {
          clearTimeout(timeout);
          if (heartbeat) clearInterval(heartbeat);
          ws.close();
          resolve(msg.channelId);
          return;
        }
      }
    }

    ws.addEventListener("message", (event) => {
      const raw = typeof event.data === "string" ? event.data : event.data.toString();
      const colonIdx1 = raw.indexOf(":");
      if (colonIdx1 === -1) return;
      const type = raw.slice(0, colonIdx1);

      switch (type) {
        case "1": {
          sendEphemeral({
            name: "user.authenticate",
            args: [{
              headers: { "X-Ms-Test-User": "False", Authorization: `Bearer ${aadToken}`, "X-MS-Migration": "True" },
              connectparams: trouter.connectparams,
            }],
          });
          sendEvent({ name: "user.activity", args: [{ state: "active", cv: "2024.23.01.2.0.1" }] });
          heartbeat = setInterval(() => {
            if (ws.readyState === WebSocket.OPEN) sendEvent({ name: "ping", args: [] });
          }, 30_000);
          register(aadToken, skypeToken, trouter.surl, epid).catch(() => {});
          break;
        }
        case "2": { ws.send("2::"); break; }
        case "3": {
          const rest = raw.slice(colonIdx1 + 1);
          const c2 = rest.indexOf(":");
          const after = rest.slice(c2 + 1);
          const c3 = after.indexOf(":");
          const dataStr = after.slice(c3 + 1);
          try {
            const parsed = JSON.parse(dataStr);
            if (parsed.id != null) ws.send(`3:::{"id":${parsed.id},"status":200}`);
            handleNotification(parsed);
          } catch { /* ignore */ }
          break;
        }
        case "5": {
          const rest5 = raw.slice(colonIdx1 + 1);
          const c2 = rest5.indexOf(":");
          const after5 = rest5.slice(c2 + 1);
          const c3 = after5.indexOf(":");
          const data5 = after5.slice(c3 + 1);
          try {
            const parsed = JSON.parse(data5);
            if (parsed.name === "trouter.request" && parsed.args?.[0]) {
              handleNotification(parsed.args[0]);
            }
          } catch { /* ignore */ }
          break;
        }
      }
    });

    ws.addEventListener("error", () => {
      clearTimeout(timeout);
      if (heartbeat) clearInterval(heartbeat);
      reject(new Error("WebSocket error during channel discovery"));
    });
  });
}

// ---------------------------------------------------------------------------
// Prerequisite checks
// ---------------------------------------------------------------------------

async function checkPrereqs(): Promise<boolean> {
  heading("Checking prerequisites");
  let allGood = true;

  // 1. Node.js
  const nodeVersion = process.version;
  const major = parseInt(nodeVersion.slice(1), 10);
  if (major >= 22) {
    ok(`Node.js ${nodeVersion}`);
  } else {
    fail(`Node.js ${nodeVersion} — need v22+ (native TS, fetch, WebSocket)`);
    allGood = false;
  }

  // 2. az CLI
  if (commandExists("az")) {
    ok("Azure CLI (az) installed");

    // Check if logged in
    try {
      const tenants = listTenantIds();
      if (tenants.length > 0) {
        ok(`az login — ${tenants.length} tenant(s) found`);
      } else {
        fail("az login — no tenants found. Run: az login");
        allGood = false;
      }
    } catch {
      fail("az login — not authenticated. Run: az login");
      allGood = false;
    }
  } else {
    fail("Azure CLI (az) not installed. Install: https://aka.ms/installazurecli");
    allGood = false;
  }

  // 3. Claude Code
  if (commandExists("claude")) {
    ok("Claude Code CLI installed");
  } else {
    fail("Claude Code CLI not found. Install: npm install -g @anthropic-ai/claude-code");
    allGood = false;
  }

  // 4. npm dependencies
  if (existsSync(join(process.cwd(), "node_modules", "@anthropic-ai", "claude-agent-sdk"))) {
    ok("npm dependencies installed");
  } else {
    warn("npm dependencies not installed. Run: npm install");
    allGood = false;
  }

  // 5. gh CLI (optional)
  if (commandExists("gh")) {
    ok("GitHub CLI (gh) installed (optional)");
  } else {
    warn("GitHub CLI (gh) not found (optional — needed for codespace SSH)");
  }

  return allGood;
}

// ---------------------------------------------------------------------------
// Channel setup
// ---------------------------------------------------------------------------

async function setupChannel(): Promise<string | null> {
  heading("Channel discovery");

  const envPath = join(process.cwd(), ".env");

  // Check existing .env
  if (existsSync(envPath)) {
    const existing = readFileSync(envPath, "utf8");
    const match = existing.match(/^CHANNEL_ID=(.+)$/m);
    if (match?.[1]) {
      ok(`Existing CHANNEL_ID found in .env`);
      info(match[1]);
      const answer = await prompt("Keep this channel? (Y/n)");
      if (answer.toLowerCase() !== "n") {
        return match[1];
      }
    }
  }

  console.log();
  info("Let's connect to a Teams channel.");
  info("We'll listen for a unique code you post in any channel.\n");

  // Get tokens
  console.log(`  ${DIM}Authenticating...${RESET}`);
  let aadToken: string;
  let skypeToken: string;
  try {
    const aad = getAadToken();
    aadToken = aad.token;
    const claims = decodeToken(aadToken);
    ok(`Authenticated as ${claims.name ?? claims.upn}`);
    skypeToken = await getSkypeToken(aadToken);
    ok("Skype token obtained");
  } catch (err: any) {
    fail(`Authentication failed: ${err.message}`);
    info("Make sure you've run: az login");
    return null;
  }

  // Generate secret code
  const secretCode = `setup-${randomBytes(4).toString("hex")}`;

  console.log();
  console.log(`  ${YELLOW}${BOLD}⚠  Create a private channel just for yourself.${RESET}`);
  console.log(`  ${YELLOW}Do NOT add anyone else or share this channel — the bot runs${RESET}`);
  console.log(`  ${YELLOW}with your credentials and full permissions. Anyone in the${RESET}`);
  console.log(`  ${YELLOW}channel can send commands that execute on your machine.${RESET}`);
  console.log();
  console.log(`  ${BOLD}Post this exact message in that channel:${RESET}`);
  console.log();
  console.log(`    ${CYAN}${BOLD}${secretCode}${RESET}`);
  console.log();
  info("Listening for your message (timeout: 2 minutes)...\n");

  try {
    const channelId = await discoverChannel(aadToken, skypeToken, secretCode);
    ok(`Channel discovered!`);
    info(channelId);

    // Write to .env
    let envContent = "";
    if (existsSync(envPath)) {
      envContent = readFileSync(envPath, "utf8");
      if (envContent.match(/^CHANNEL_ID=/m)) {
        envContent = envContent.replace(/^CHANNEL_ID=.*$/m, `CHANNEL_ID=${channelId}`);
      } else {
        envContent = envContent.trimEnd() + `\nCHANNEL_ID=${channelId}\n`;
      }
    } else {
      envContent = `CHANNEL_ID=${channelId}\n`;
    }
    writeFileSync(envPath, envContent);
    ok("Written to .env");

    return channelId;
  } catch (err: any) {
    fail(err.message);
    return null;
  }
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

async function main() {
  console.log(`\n${BOLD}╔══════════════════════════════════════╗${RESET}`);
  console.log(`${BOLD}║  Teams ↔ Claude Code Bridge Setup    ║${RESET}`);
  console.log(`${BOLD}╚══════════════════════════════════════╝${RESET}`);

  console.log(`\n  ${RED}${BOLD}⚠  WARNING${RESET}`);
  console.log(`  ${YELLOW}This bot runs Claude Code with ${BOLD}--dangerously-skip-permissions${RESET}${YELLOW}.${RESET}`);
  console.log(`  ${YELLOW}It uses ${BOLD}your Azure credentials${RESET}${YELLOW} (az login) to connect to Teams.${RESET}`);
  console.log(`  ${YELLOW}Claude Code will have ${BOLD}full access${RESET}${YELLOW} to your filesystem and tools${RESET}`);
  console.log(`  ${YELLOW}with no permission prompts. Only use in a trusted environment.${RESET}`);

  const prereqsOk = await checkPrereqs();

  if (!prereqsOk) {
    console.log(`\n  ${RED}Fix the issues above before continuing.${RESET}\n`);
    process.exit(1);
  }

  const channelId = await setupChannel();

  if (!channelId) {
    console.log(`\n  ${RED}Channel setup failed.${RESET}\n`);
    process.exit(1);
  }

  heading("Ready!");
  console.log(`  Start the bridge with: ${CYAN}npm start${RESET}\n`);
}

main().catch((err) => {
  console.error("\nUnexpected error:", err);
  process.exit(1);
});
