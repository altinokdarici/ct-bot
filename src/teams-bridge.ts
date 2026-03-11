import { execSync } from "node:child_process";
import { gunzipSync } from "node:zlib";
import { randomUUID } from "node:crypto";
import type { TeamsMessage } from "./types.ts";

// ---------------------------------------------------------------------------
// Config
// ---------------------------------------------------------------------------

const EPID = randomUUID();
let conCounter = 0;

// ---------------------------------------------------------------------------
// Auth — az CLI tokens
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

function decodeToken(token: string): Record<string, any> {
  return JSON.parse(Buffer.from(token.split(".")[1]!, "base64url").toString());
}

function getAadToken(): { token: string; tenantId: string } {
  for (const tenantId of listTenantIds()) {
    const token = azToken(tenantId, "https://api.spaces.skype.com");
    if (token) return { token, tenantId };
  }
  throw new Error("Could not get AAD token. Run: az login");
}

function getIc3Token(): { token: string; tenantId: string } | null {
  for (const tenantId of listTenantIds()) {
    const token = azToken(tenantId, "https://ic3.teams.office.com");
    if (token) return { token, tenantId };
  }
  return null;
}

async function getSkypeToken(aadToken: string): Promise<string> {
  const resp = await fetch("https://teams.microsoft.com/api/authsvc/v1.0/authz", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${aadToken}`,
      "Content-Type": "application/json",
    },
    body: "{}",
  });
  if (!resp.ok) throw new Error(`authz failed: ${resp.status}`);
  const data: any = await resp.json();
  const skypeToken = data.tokens?.skypeToken;
  if (!skypeToken) throw new Error("No skypeToken in authz response");
  return skypeToken;
}

// ---------------------------------------------------------------------------
// Region discovery (for sending)
// ---------------------------------------------------------------------------

async function probeRegion(token: string): Promise<string | null> {
  const resp = await fetch(
    "https://teams.cloud.microsoft/api/chatsvc/amer/v1/users/ME/properties",
    { headers: { Authorization: `Bearer ${token}`, "x-ms-migration": "True" }, redirect: "manual" },
  );
  if (resp.status === 200) return "amer";
  if (resp.status === 401 || resp.status === 403) return null;
  const location = resp.headers.get("location") || "";
  const m = location.match(/\/api\/chatsvc\/([^/]+)\//);
  if (m) return m[1] ?? null;
  for (const region of ["noam-pilot1", "noam-pilot2", "emea", "apac"]) {
    const r = await fetch(
      `https://teams.cloud.microsoft/api/chatsvc/${region}/v1/users/ME/properties`,
      { headers: { Authorization: `Bearer ${token}`, "x-ms-migration": "True" } },
    );
    if (r.status === 200) return region;
  }
  return null;
}

// ---------------------------------------------------------------------------
// Trouter setup
// ---------------------------------------------------------------------------

const TROUTER_GO = "https://go.trouter.teams.microsoft.com";

interface TrouterSession {
  socketio: string;
  surl: string;
  connectparams: Record<string, string>;
  ccid: string;
}

async function setupTrouter(skypeToken: string): Promise<TrouterSession> {
  const url = `${TROUTER_GO}/v4/a?epid=${encodeURIComponent(EPID)}`;
  const resp = await fetch(url, {
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

function buildConnectQuery(connectparams: Record<string, string>, ccid: string): string {
  conCounter++;
  const parts = ["v=v4"];
  for (const [key, value] of Object.entries(connectparams)) {
    parts.push(`${key}=${encodeURIComponent(value)}`);
  }
  const tc = JSON.stringify({ cv: "2024.23.01.2", ua: "TeamsCDL", hr: "", v: "49/25010202142" });
  parts.push(`tc=${encodeURIComponent(tc)}`);
  parts.push(`con_num=${Date.now()}_${conCounter}`);
  parts.push(`epid=${encodeURIComponent(EPID)}`);
  if (ccid) parts.push(`ccid=${encodeURIComponent(ccid)}`);
  parts.push("auth=true");
  parts.push("timeout=40");
  return parts.join("&");
}

async function socketioHandshake(
  socketio: string,
  connectparams: Record<string, string>,
  ccid: string,
  skypeToken: string,
): Promise<string> {
  const query = buildConnectQuery(connectparams, ccid);
  const resp = await fetch(`${socketio}socket.io/1/?${query}`, {
    headers: { "X-Skypetoken": skypeToken },
  });
  if (!resp.ok) throw new Error(`Socket.IO handshake failed: ${resp.status}`);
  const text = await resp.text();
  return text.split(":")[0]!;
}

async function register(aadToken: string, skypeToken: string, surl: string): Promise<void> {
  const url = "https://teams.microsoft.com/registrar/prod/V2/registrations";
  const registrations = [
    { appId: "NextGenCalling", templateKey: "DesktopNgc_2.3:SkypeNgc", path: `${surl}NGCallManagerWin` },
    { appId: "SkypeSpacesWeb", templateKey: "SkypeSpacesWeb_2.3", path: `${surl}SkypeSpacesWeb` },
    { appId: "TeamsCDLWebWorker", templateKey: "TeamsCDLWebWorker_2.1", path: surl },
  ];
  for (const reg of registrations) {
    const body = {
      clientDescription: {
        appId: reg.appId, aesKey: "", languageId: "en-US", platform: "edge",
        templateKey: reg.templateKey, platformUIVersion: "49/25010202142",
      },
      registrationId: EPID, nodeId: "",
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

// ---------------------------------------------------------------------------
// Notification parsing
// ---------------------------------------------------------------------------

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

function extractTeamsMessage(body: any, channelId: string): TeamsMessage | null {
  const resource = body.resource ?? body;
  const msgType = resource.messagetype ?? resource.messageType;

  const to = resource.to ?? "";

  if (!msgType || !["RichText/Html", "Text", "RichText"].includes(msgType)) return null;

  const html = resource.content ?? resource.body ?? "";
  const text = msgType.startsWith("RichText") ? stripHtml(html) : html.trim();
  if (!text) return null;

  // Channel filter: resource.to must match our channel
  if (to !== channelId) return null;

  // Self-filter: skip messages the bot sent (prefixed with "AI:")
  if (text.startsWith("AI:")) return null;

  const messageId = resource.id ?? resource.version ?? "";
  const parentMessageId = String(body.parentmessageid ?? "");

  // For new top-level messages, use the message's own ID as thread ID
  const threadId = parentMessageId || messageId;
  const isNewThread = !parentMessageId || parentMessageId === messageId;

  const fromId = resource.from?.split("/").pop() ?? "";
  const from = resource.imdisplayname ?? resource.fromDisplayNameInToken ?? "?";
  const arrivalTime = resource.originalarrivaltime ?? resource.composetime;
  const time = arrivalTime ? new Date(arrivalTime).toISOString() : new Date().toISOString();

  return { from, fromId, channelId, threadId, messageId, isNewThread, text, html, time };
}

// ---------------------------------------------------------------------------
// WebSocket connection
// ---------------------------------------------------------------------------

interface WsCallbacks {
  onConnected?: () => void;
  onNotification?: (data: any) => void;
  onClose?: () => void;
  onError?: () => void;
}

function connectWebSocket(
  socketio: string,
  sessionId: string,
  connectparams: Record<string, string>,
  ccid: string,
  skypeToken: string,
  aadToken: string,
  callbacks: WsCallbacks,
): WebSocket {
  const query = buildConnectQuery(connectparams, ccid);
  const wsBase = socketio.replace(/^http/, "ws");
  const wsUrl = `${wsBase}socket.io/1/websocket/${sessionId}?${query}`;

  const ws = new WebSocket(wsUrl, { headers: { "X-Skypetoken": skypeToken } } as any);
  let heartbeatInterval: ReturnType<typeof setInterval> | null = null;
  let commandCount = 0;

  const sendEvent = (json: any) => ws.send(`5:${commandCount++}+::${JSON.stringify(json)}`);
  const sendEphemeral = (json: any) => ws.send(`5:::${JSON.stringify(json)}`);

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
            connectparams,
          }],
        });
        sendEvent({ name: "user.activity", args: [{ state: "active", cv: "2024.23.01.2.0.1" }] });
        heartbeatInterval = setInterval(() => {
          if (ws.readyState === WebSocket.OPEN) sendEvent({ name: "ping", args: [] });
        }, 30_000);
        callbacks.onConnected?.();
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
          callbacks.onNotification?.(parsed);
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
            callbacks.onNotification?.(parsed.args[0]);
          }
        } catch { /* ignore */ }
        break;
      }
      case "6": case "8": break;
    }
  });

  ws.addEventListener("close", () => {
    if (heartbeatInterval) clearInterval(heartbeatInterval);
    callbacks.onClose?.();
  });
  ws.addEventListener("error", () => {
    if (heartbeatInterval) clearInterval(heartbeatInterval);
    callbacks.onError?.();
  });

  return ws;
}

// ---------------------------------------------------------------------------
// Public API: Listener
// ---------------------------------------------------------------------------

export async function startListener(
  channelId: string,
  onMessage: (msg: TeamsMessage) => void,
): Promise<void> {
  console.log("[teams] Getting AAD token...");
  let { token: aadToken } = getAadToken();
  const claims = decodeToken(aadToken);
  console.log(`[teams] Authenticated as: ${claims.name ?? claims.upn}`);

  console.log("[teams] Getting Skype token...");
  let skypeToken = await getSkypeToken(aadToken);

  console.log("[teams] Setting up Trouter...");
  let trouter = await setupTrouter(skypeToken);
  let wsSessionId = await socketioHandshake(trouter.socketio, trouter.connectparams, trouter.ccid, skypeToken);

  const seenMessages = new Map<string, number>();

  const TOKEN_REFRESH_MS = 45 * 60 * 1000;
  let lastTokenRefresh = Date.now();

  const refreshInterval = setInterval(async () => {
    if (Date.now() - lastTokenRefresh > TOKEN_REFRESH_MS) {
      try {
        const newAad = getAadToken();
        aadToken = newAad.token;
        skypeToken = await getSkypeToken(aadToken);
        lastTokenRefresh = Date.now();
        console.log("[teams] Tokens refreshed");
      } catch (err: any) {
        console.error("[teams] Token refresh error:", err.message);
      }
    }
  }, 60_000);

  function connect() {
    console.log("[teams] Connecting WebSocket...");
    connectWebSocket(trouter.socketio, wsSessionId, trouter.connectparams, trouter.ccid, skypeToken, aadToken, {
      async onConnected() {
        await register(aadToken, skypeToken, trouter.surl);
        console.log("[teams] Listening for messages...");
      },
      onNotification(data) {
        const items = Array.isArray(data) ? data : [data];
        for (const item of items) {
          const body = parseNotification(item);
          if (!body) continue;

          const msg = extractTeamsMessage(body, channelId);
          if (!msg) continue;

          // Deduplicate
          if (seenMessages.has(msg.messageId)) continue;
          seenMessages.set(msg.messageId, Date.now());

          // Prune old entries
          const now = Date.now();
          for (const [id, ts] of seenMessages) {
            if (now - ts > 60_000) seenMessages.delete(id);
          }

          onMessage(msg);
        }
      },
      onClose() {
        console.log("[teams] Connection lost, reconnecting in 2s...");
        setTimeout(reconnect, 2000);
      },
      onError() {
        console.log("[teams] Connection error, reconnecting in 2s...");
        setTimeout(reconnect, 2000);
      },
    });
  }

  async function reconnect() {
    try {
      trouter = await setupTrouter(skypeToken);
      wsSessionId = await socketioHandshake(trouter.socketio, trouter.connectparams, trouter.ccid, skypeToken);
      connect();
    } catch (err: any) {
      console.error("[teams] Reconnect failed:", err.message, "— retrying in 5s...");
      setTimeout(reconnect, 5000);
    }
  }

  connect();

  process.on("SIGINT", () => {
    clearInterval(refreshInterval);
    console.log("[teams] Shutting down...");
    process.exit(0);
  });
}

// ---------------------------------------------------------------------------
// Public API: Send thread reply
// ---------------------------------------------------------------------------

let cachedSendAuth: { token: string; region: string; myOid: string; displayName: string } | null = null;

async function getSendAuth() {
  if (cachedSendAuth) return cachedSendAuth;

  const ic3 = getIc3Token();
  if (!ic3) throw new Error("Could not get ic3 token for sending");

  const region = await probeRegion(ic3.token);
  if (!region) throw new Error("Could not discover region for sending");

  const claims = decodeToken(ic3.token);
  cachedSendAuth = {
    token: ic3.token,
    region,
    myOid: claims.oid,
    displayName: claims.name || "",
  };
  return cachedSendAuth;
}

// Refresh send auth periodically
setInterval(() => { cachedSendAuth = null; }, 40 * 60 * 1000);

// SAFETY: Only these channels are allowed to receive messages
const ALLOWED_CHANNELS = new Set(
  process.env.CHANNEL_ID ? [process.env.CHANNEL_ID] : [],
);

export async function sendThreadReply(
  channelId: string,
  threadRootId: string,
  message: string,
): Promise<string | null> {
  // GUARDRAIL: Never send to an unallowed channel
  if (!ALLOWED_CHANNELS.has(channelId)) {
    console.error(`[teams] BLOCKED: Attempted to send to unauthorized channel: ${channelId}`);
    return null;
  }

  const auth = await getSendAuth();
  const now = new Date().toISOString().replace(/\.\d{3}Z$/, ".000Z");
  const clientMsgId = String(Math.floor(Math.random() * 9e18) + 1e18);

  const convWithThread = `${channelId};messageid=${threadRootId}`;
  const encodedConv = encodeURIComponent(convWithThread);
  const url = `https://teams.cloud.microsoft/api/chatsvc/${auth.region}/v1/users/ME/conversations/${encodedConv}/messages`;

  const body = {
    id: "-1",
    type: "Message",
    conversationid: channelId,
    conversationLink: `https://teams.cloud.microsoft/api/chatsvc/${auth.region}/v1/users/ME/conversations/${convWithThread}`,
    from: `8:orgid:${auth.myOid}`,
    fromUserId: `8:orgid:${auth.myOid}`,
    composetime: now,
    originalarrivaltime: now,
    content: `<b>AI:</b> ${message}`,
    messagetype: "RichText/Html",
    contenttype: "Text",
    imdisplayname: auth.displayName,
    clientmessageid: clientMsgId,
    callId: "",
    state: 0,
    version: "0",
    amsreferences: [],
    properties: {
      importance: "",
      subject: "",
      title: "",
      cards: "[]",
      links: "[]",
      mentions: "[]",
      onbehalfof: null,
      files: "[]",
      policyViolation: null,
      formatVariant: "TEAMS",
    },
    crossPostChannels: [],
  };

  const resp = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${auth.token}`,
      "Content-Type": "application/json",
      behavioroverride: "redirectAs404",
      "x-ms-migration": "True",
    },
    body: JSON.stringify(body),
  });

  if (resp.status !== 201) {
    const text = await resp.text();
    console.error(`[teams] Send failed: ${resp.status} ${text}`);
    return null;
  }

  // Return the server-assigned message ID for later editing
  try {
    const data: any = await resp.json();
    const msgId = data.OriginalArrivalTime ?? data.id ?? data.version ?? null;
    console.log(`[teams] Sent message, server id=${data.id}, version=${data.version}, OriginalArrivalTime=${data.OriginalArrivalTime}, using=${msgId ?? clientMsgId}`);
    return msgId ?? clientMsgId;
  } catch {
    console.log(`[teams] Sent message, no JSON body, using clientMsgId=${clientMsgId}`);
    return clientMsgId;
  }
}

export async function editThreadReply(
  channelId: string,
  threadRootId: string,
  messageId: string,
  newContent: string,
): Promise<void> {
  if (!ALLOWED_CHANNELS.has(channelId)) return;

  const auth = await getSendAuth();
  const convWithThread = `${channelId};messageid=${threadRootId}`;
  const encodedConv = encodeURIComponent(convWithThread);
  const url = `https://teams.cloud.microsoft/api/chatsvc/${auth.region}/v1/users/ME/conversations/${encodedConv}/messages/${messageId}`;

  const body = {
    content: `<b>AI:</b> ${newContent}`,
    messagetype: "RichText/Html",
    contenttype: "Text",
    amsreferences: [],
    properties: {
      importance: "",
      subject: "",
      title: "",
      cards: "[]",
      links: "[]",
      mentions: "[]",
      onbehalfof: null,
      files: "[]",
      policyViolation: null,
      formatVariant: "TEAMS",
    },
  };

  console.log(`[teams] Editing message ${messageId} (${JSON.stringify(body).length} bytes)`);

  const resp = await fetch(url, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${auth.token}`,
      "Content-Type": "application/json",
      behavioroverride: "redirectAs404",
      "x-ms-migration": "True",
    },
    body: JSON.stringify(body),
  });

  const respText = await resp.text();
  console.log(`[teams] Edit response: ${resp.status} ${respText.slice(0, 200)}`);

  if (resp.status !== 200 && resp.status !== 204 && resp.status !== 201) {
    console.error(`[teams] Edit failed (${resp.status}) msgId=${messageId}: ${respText.slice(0, 300)}`);
  }
}
