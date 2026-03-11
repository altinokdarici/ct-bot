export interface TeamsMessage {
  /** Sender display name */
  from: string;
  /** Sender OID (e.g. "8:orgid:xxx") */
  fromId: string;
  /** Channel conversation ID */
  channelId: string;
  /** Thread root message ID — same for all messages in a thread */
  threadId: string;
  /** This message's own ID */
  messageId: string;
  /** Whether this is a new thread (first message) */
  isNewThread: boolean;
  /** Plain text content (HTML stripped) */
  text: string;
  /** Raw HTML content */
  html: string;
  /** Arrival timestamp */
  time: string;
}

export interface Session {
  /** Claude Agent SDK session ID for resumption */
  sessionId: string | null;
  /** Thread root message ID (key) */
  threadId: string;
  /** Working directory */
  cwd: string;
  /** Session status */
  status: "active" | "idle" | "handoff";
  /** Timestamp */
  createdAt: number;
  lastActiveAt: number;
  /** Conversation context carried across handoffs */
  handoffContext: string | null;
}

export interface HandoffRequest {
  type: "local";
  target: string; // local path
  intent: string | null; // extracted actionable task for the new environment
}
