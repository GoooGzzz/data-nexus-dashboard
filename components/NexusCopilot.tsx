"use client";
import {
  useState,
  useRef,
  useEffect,
  forwardRef,
  useImperativeHandle,
  useCallback,
} from "react";
import { regionData, buildDashboardContext, type RegionKey } from "@/lib/dashboard-data";

interface Message {
  id: string;
  role: "user" | "assistant";
  content: string;
  timestamp: Date;
  isError?: boolean;
}

export interface NexusCopilotHandle {
  /** Opens the panel and optionally sends a question immediately. */
  open: (question?: string) => void;
}

interface NexusCopilotProps {
  /** Currently selected/filtered region on the dashboard, if any — grounds the AI's answers. */
  selectedRegion?: RegionKey | null;
}

const SUGGESTED_PROMPTS = [
  "Why is Upper Egypt underperforming?",
  "Which region needs attention today?",
  "Summarize today's anomalies",
  "What would fix the attendance gap?",
];

function formatTime(d: Date) {
  return d.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
}

// Tiny, dependency-free markdown: **bold**, `code`, and line breaks/bullets.
// Markdown rendering: **bold**, `code`, and line breaks/bullets
// instead of dumping raw asterisks into the chat.
function renderInline(text: string, keyPrefix: string) {
  const parts = text.split(/(\*\*[^*]+\*\*|`[^`]+`)/g).filter(Boolean);
  return parts.map((part, i) => {
    if (part.startsWith("**") && part.endsWith("**")) {
      return (
        <strong key={`${keyPrefix}-${i}`} style={{ color: "#fff" }}>
          {part.slice(2, -2)}
        </strong>
      );
    }
    if (part.startsWith("`") && part.endsWith("`")) {
      return (
        <code
          key={`${keyPrefix}-${i}`}
          style={{
            background: "#1a1a1a",
            border: "1px solid #333",
            borderRadius: "4px",
            padding: "1px 5px",
            fontSize: "12px",
            color: "#39ff14",
          }}
        >
          {part.slice(1, -1)}
        </code>
      );
    }
    return <span key={`${keyPrefix}-${i}`}>{part}</span>;
  });
}

// Detects a contiguous block of markdown table lines starting at index i.
// A valid table needs a header row, a separator row (---|---), then >=1 data row.
function isTableSeparator(line: string) {
  return /^\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)+\|?$/.test(line.trim());
}

function parseTableRow(line: string): string[] {
  return line
    .trim()
    .replace(/^\|/, "")
    .replace(/\|$/, "")
    .split("|")
    .map((c) => c.trim());
}

function MarkdownTable({ header, rows, keyPrefix }: { header: string[]; rows: string[][]; keyPrefix: string }) {
  return (
    <table
      style={{
        width: "100%",
        borderCollapse: "collapse",
        margin: "8px 0",
        fontSize: "12.5px",
      }}
    >
      <thead>
        <tr>
          {header.map((h, i) => (
            <th
              key={`${keyPrefix}-h-${i}`}
              style={{
                textAlign: "left",
                color: "#39ff14",
                borderBottom: "1px solid #2a2a2a",
                padding: "5px 8px",
                fontWeight: 600,
                whiteSpace: "nowrap",
              }}
            >
              {renderInline(h, `${keyPrefix}-h-${i}`)}
            </th>
          ))}
        </tr>
      </thead>
      <tbody>
        {rows.map((row, ri) => (
          <tr key={`${keyPrefix}-r-${ri}`} style={{ background: ri % 2 === 0 ? "transparent" : "rgba(255,255,255,0.03)" }}>
            {row.map((cell, ci) => (
              <td
                key={`${keyPrefix}-r-${ri}-c-${ci}`}
                style={{
                  padding: "5px 8px",
                  borderBottom: "1px solid #1f1f1f",
                  color: "#e5e5e5",
                  verticalAlign: "top",
                }}
              >
                {renderInline(cell, `${keyPrefix}-r-${ri}-c-${ci}`)}
              </td>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
  );
}

function MessageBody({ content }: { content: string }) {
  const lines = content.split("\n");
  const blocks: React.ReactNode[] = [];
  let i = 0;

  while (i < lines.length) {
    const line = lines[i];
    const trimmed = line.trim();

    // Table block: header row, separator row, then data rows
    if (trimmed.includes("|") && i + 1 < lines.length && isTableSeparator(lines[i + 1])) {
      const header = parseTableRow(trimmed);
      let j = i + 2;
      const rows: string[][] = [];
      while (j < lines.length && lines[j].trim().includes("|")) {
        rows.push(parseTableRow(lines[j]));
        j++;
      }
      blocks.push(<MarkdownTable key={`t${i}`} header={header} rows={rows} keyPrefix={`t${i}`} />);
      i = j;
      continue;
    }

    const isBullet = /^[-*]\s+/.test(trimmed);
    const display = isBullet ? line.replace(/^[-*]\s+/, "") : line;
    blocks.push(
      <div key={i} style={{ marginLeft: isBullet ? "14px" : 0, marginTop: i === 0 ? 0 : "4px" }}>
        {isBullet && "• "}
        {display ? renderInline(display, `l${i}`) : "\u00A0"}
      </div>
    );
    i++;
  }

  return <>{blocks}</>;
}

const NexusCopilot = forwardRef<NexusCopilotHandle, NexusCopilotProps>(function NexusCopilot(
  { selectedRegion = null },
  ref
) {
  const [open, setOpen] = useState(false);
  const [messages, setMessages] = useState<Message[]>([]);
  const [input, setInput] = useState("");
  const [isLoading, setIsLoading] = useState(false);
  const [isStreaming, setIsStreaming] = useState(false);

  const messagesEndRef = useRef<HTMLDivElement>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const abortRef = useRef<AbortController | null>(null);

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages]);

  // Close on Escape
  useEffect(() => {
    if (!open) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") setOpen(false);
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [open]);

  const sendMessage = useCallback(
    async (text: string) => {
      if (!text.trim() || isLoading) return;

      const userMessage: Message = {
        id: `${Date.now()}-u`,
        role: "user",
        content: text.trim(),
        timestamp: new Date(),
      };

      const history = [...messages, userMessage];
      setMessages(history);
      setInput("");
      setIsLoading(true);
      setIsStreaming(true);

      const assistantId = `${Date.now()}-a`;
      // Seed an empty assistant message we'll fill in as tokens arrive.
      setMessages((prev) => [
        ...prev,
        { id: assistantId, role: "assistant", content: "", timestamp: new Date() },
      ]);

      const controller = new AbortController();
      abortRef.current = controller;

      try {
        const response = await fetch("/api/nexus", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          signal: controller.signal,
          body: JSON.stringify({
            system: [
              "You are Nexus Prime, elite retail intelligence AI for 4000+ mobile shops in Egypt.",
              "Ground every answer in the live dashboard snapshot provided below — cite real numbers from it exactly as given, never invent or round differently.",
              "",
              "FORMAT RULES (follow exactly):",
              "1. When answering about a region's metrics, present them as a markdown table with columns: Metric | Value. Rows must use these exact labels in this order: Performance, Sales, Shops, Attendance Gap, Trend — using the exact figures from the snapshot below.",
              "2. After the table, add a short '**Quick read**' line (1 sentence).",
              "3. If you give recommendations/actions, list them as a markdown table with columns: Action | Why it matters | Expected impact — 2-4 rows max.",
              "4. If comparing multiple regions, use one markdown table with regions as columns and metrics as rows.",
              "5. Never output a wall of bullets for numeric data — numbers always go in a table. Bullets are only for non-numeric commentary.",
              "6. Keep total reply under ~150 words excluding table cells. No filler intro sentences before the table.",
              "",
              buildDashboardContext(selectedRegion),
            ].join("\n"),
            messages: history.map((m) => ({ role: m.role, content: m.content })),
          }),
        });

        if (!response.ok || !response.body) {
          throw new Error(`Nexus API error (${response.status})`);
        }

        // The route streams Server-Sent Events in the OpenAI-compatible
        // format: lines like `data: {"choices":[{"delta":{"content":"..."}}]}`
        // ending in `data: [DONE]`. We read it manually here instead of
        // calling response.json() — that was the bug that made every
        // reply fall through to the offline mock message.
        const reader = response.body.getReader();
        const decoder = new TextDecoder();
        let buffer = "";
        let accumulated = "";

        while (true) {
          const { value, done } = await reader.read();
          if (done) break;

          buffer += decoder.decode(value, { stream: true });
          const lines = buffer.split("\n");
          buffer = lines.pop() ?? ""; // keep the last (possibly partial) line

          for (const rawLine of lines) {
            const line = rawLine.trim();
            if (!line.startsWith("data:")) continue;
            const data = line.slice(5).trim();
            if (data === "[DONE]") continue;

            try {
              const parsed = JSON.parse(data);
              const delta: string | undefined = parsed?.choices?.[0]?.delta?.content;
              if (delta) {
                accumulated += delta;
                setMessages((prev) =>
                  prev.map((m) => (m.id === assistantId ? { ...m, content: accumulated } : m))
                );
              }
            } catch {
              // Partial/non-JSON chunk — safe to skip, more data is coming.
            }
          }
        }

        if (!accumulated) {
          throw new Error("Empty response from Nexus");
        }
      } catch (err) {
        if ((err as Error).name === "AbortError") return;
        console.error("Nexus copilot error:", err);
        setMessages((prev) =>
          prev.map((m) =>
            m.id === assistantId
              ? {
                  ...m,
                  isError: true,
                  content:
                    "Nexus couldn't reach the intelligence engine just now. The intelligence engine encountered an issue. Please try again.",
                }
              : m
          )
        );
      } finally {
        setIsLoading(false);
        setIsStreaming(false);
        abortRef.current = null;
      }
    },
    [messages, isLoading, selectedRegion]
  );

  useImperativeHandle(ref, () => ({
    open: (question?: string) => {
      setOpen(true);
      if (question) {
        // let the panel mount before sending
        setTimeout(() => sendMessage(question), 50);
      }
    },
  }));

  const handleKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      sendMessage(input);
    }
  };

  return (
    <>
      {/* Floating Button */}
      <button
        onClick={() => setOpen((o) => !o)}
        aria-label={open ? "Close Nexus copilot" : "Open Nexus copilot"}
        style={{
          position: "fixed",
          bottom: "30px",
          right: "30px",
          zIndex: 10000,
          width: "66px",
          height: "66px",
          borderRadius: "50%",
          background: "linear-gradient(135deg, #0a0a0a, #1a1a1a)",
          border: "2px solid #39ff14",
          color: "#39ff14",
          fontWeight: "bold",
          fontSize: "11px",
          boxShadow: open
            ? "0 0 18px rgba(57,255,20,0.4)"
            : "0 0 30px rgba(57,255,20,0.6)",
          cursor: "pointer",
          transition: "box-shadow 0.2s ease, transform 0.2s ease",
        }}
        onMouseEnter={(e) => (e.currentTarget.style.transform = "scale(1.06)")}
        onMouseLeave={(e) => (e.currentTarget.style.transform = "scale(1)")}
      >
        {isLoading ? (
          <span
            style={{
              display: "inline-block",
              width: 16,
              height: 16,
              border: "2px solid #39ff14",
              borderTopColor: "transparent",
              borderRadius: "50%",
              animation: "nexus-spin 0.8s linear infinite",
            }}
          />
        ) : (
          "NEXUS"
        )}
      </button>

      {open && (
        <div
          onClick={() => setOpen(false)}
          style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", zIndex: 10000 }}
        />
      )}

      {/* Sidebar */}
      <div
        style={{
          position: "fixed",
          top: 0,
          right: 0,
          bottom: 0,
          width: "480px",
          maxWidth: "100vw",
          background: "#0a0a0a",
          borderLeft: "2px solid #39ff14",
          zIndex: 10001,
          display: "flex",
          flexDirection: "column",
          boxShadow: "-10px 0 40px rgba(0,0,0,0.8)",
          transform: open ? "translateX(0)" : "translateX(100%)",
          transition: "transform 0.25s ease",
        }}
      >
        <div style={{ padding: "20px", borderBottom: "1px solid #39ff14", display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
          <div>
            <h2 style={{ color: "#39ff14", margin: 0, letterSpacing: "0.05em" }}>NEXUS PRIME</h2>
            <p style={{ color: "#666", fontSize: "12px", margin: "4px 0 0" }}>
              <span
                style={{
                  display: "inline-block",
                  width: 6,
                  height: 6,
                  borderRadius: "50%",
                  background: isStreaming ? "#39ff14" : "#444",
                  marginRight: 6,
                  boxShadow: isStreaming ? "0 0 6px #39ff14" : "none",
                }}
              />
              {isStreaming ? "ANALYZING..." : "RETAIL INTELLIGENCE ACTIVE"}
              {selectedRegion && (
                <span style={{ color: "#39ff14" }}> · {selectedRegion}</span>
              )}
            </p>
          </div>
          <button
            onClick={() => setOpen(false)}
            aria-label="Close"
            style={{ background: "none", border: "none", color: "#666", fontSize: "20px", cursor: "pointer", lineHeight: 1 }}
          >
            ×
          </button>
        </div>

        <div style={{ flex: 1, padding: "20px", overflowY: "auto" }}>
          {messages.length === 0 && (
            <div style={{ marginTop: "30px" }}>
              <p style={{ color: "#666", textAlign: "center", fontSize: "13px" }}>
                Ask me anything about sales, attendance, or regions —
                <br />I'm reading live numbers off this dashboard.
              </p>
              <div style={{ display: "flex", flexWrap: "wrap", gap: "8px", marginTop: "18px", justifyContent: "center" }}>
                {SUGGESTED_PROMPTS.map((p) => (
                  <button
                    key={p}
                    onClick={() => sendMessage(p)}
                    style={{
                      background: "#111",
                      border: "1px solid #2a2a2a",
                      color: "#aaa",
                      borderRadius: "16px",
                      padding: "7px 12px",
                      fontSize: "12px",
                      cursor: "pointer",
                      transition: "border-color 0.15s, color 0.15s",
                    }}
                    onMouseEnter={(e) => {
                      e.currentTarget.style.borderColor = "#39ff14";
                      e.currentTarget.style.color = "#39ff14";
                    }}
                    onMouseLeave={(e) => {
                      e.currentTarget.style.borderColor = "#2a2a2a";
                      e.currentTarget.style.color = "#aaa";
                    }}
                  >
                    {p}
                  </button>
                ))}
              </div>
            </div>
          )}

          {messages.map((msg, idx) => {
            const isLastAssistant =
              msg.role === "assistant" && idx === messages.length - 1 && isStreaming;
            return (
              <div
                key={msg.id}
                style={{
                  marginBottom: "18px",
                  textAlign: msg.role === "user" ? "right" : "left",
                }}
              >
                <div style={{ display: "flex", alignItems: "baseline", gap: "8px", justifyContent: msg.role === "user" ? "flex-end" : "flex-start" }}>
                  <strong style={{ color: msg.role === "user" ? "#fff" : msg.isError ? "#f87171" : "#39ff14", fontSize: "11px", letterSpacing: "0.05em" }}>
                    {msg.role === "user" ? "YOU" : "NEXUS"}
                  </strong>
                  <span style={{ color: "#444", fontSize: "10px" }}>{formatTime(msg.timestamp)}</span>
                </div>
                <div
                  style={{
                    marginTop: "4px",
                    fontSize: "13.5px",
                    lineHeight: 1.5,
                    color: msg.isError ? "#f87171" : "#e5e5e5",
                    display: msg.role === "assistant" ? "block" : "inline-block",
                    textAlign: "left",
                    width: msg.role === "assistant" ? "100%" : undefined,
                    maxWidth: "100%",
                  }}
                >
                  {msg.content ? (
                    <MessageBody content={msg.content} />
                  ) : (
                    <span style={{ color: "#555" }}>thinking…</span>
                  )}
                  {isLastAssistant && msg.content && (
                    <span
                      style={{
                        display: "inline-block",
                        width: "7px",
                        height: "13px",
                        background: "#39ff14",
                        marginLeft: "3px",
                        verticalAlign: "text-bottom",
                        animation: "nexus-blink 1s step-end infinite",
                      }}
                    />
                  )}
                </div>
              </div>
            );
          })}
          <div ref={messagesEndRef} />
        </div>

        <div style={{ padding: "16px", borderTop: "1px solid #333" }}>
          <div style={{ display: "flex", gap: "8px", alignItems: "flex-end" }}>
            <textarea
              ref={textareaRef}
              value={input}
              onChange={(e) => setInput(e.target.value)}
              onKeyDown={handleKeyDown}
              placeholder="Ask Nexus anything..."
              rows={1}
              style={{
                flex: 1,
                padding: "12px 14px",
                background: "#111",
                border: "1px solid #39ff14",
                color: "white",
                borderRadius: "8px",
                resize: "none",
                fontFamily: "inherit",
                fontSize: "13.5px",
                maxHeight: "120px",
              }}
              disabled={isLoading}
            />
            <button
              onClick={() => sendMessage(input)}
              disabled={isLoading || !input.trim()}
              aria-label="Send"
              style={{
                background: isLoading || !input.trim() ? "#1a1a1a" : "#39ff14",
                color: isLoading || !input.trim() ? "#555" : "#000",
                border: "none",
                borderRadius: "8px",
                width: "42px",
                height: "42px",
                fontWeight: "bold",
                cursor: isLoading || !input.trim() ? "not-allowed" : "pointer",
                flexShrink: 0,
              }}
            >
              →
            </button>
          </div>
        </div>
      </div>

      <style>{`
        @keyframes nexus-blink { 50% { opacity: 0; } }
        @keyframes nexus-spin { to { transform: rotate(360deg); } }
      `}</style>
    </>
  );
});

export default NexusCopilot;
