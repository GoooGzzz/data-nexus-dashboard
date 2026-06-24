import { NextRequest, NextResponse } from "next/server";

export async function POST(request: NextRequest) {
  try {
    const body = await request.json();

    if (!Array.isArray(body.messages) || body.messages.length === 0) {
      return NextResponse.json({ error: "messages array is required" }, { status: 400 });
    }

    const response = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "anthropic-version": "2023-06-01",
        "anthropic-dangerous-direct-browser-access": "true",
      },
      body: JSON.stringify({
        model: "claude-sonnet-4-6",
        max_tokens: 1024,
        system: body.system || "You are Nexus, elite retail intelligence AI.",
        messages: body.messages,
        stream: true,
      }),
    });

    if (!response.ok || !response.body) {
      const detail = await response.text().catch(() => "");
      console.error("Anthropic API error", response.status, detail);
      return NextResponse.json(
        { error: "Anthropic API error", status: response.status, detail },
        { status: response.status }
      );
    }

    // Anthropic streams using SSE with event: content_block_delta
    // We transform it to OpenAI-compatible format for the client
    const encoder = new TextEncoder();
    const transformedStream = new ReadableStream({
      async start(controller) {
        const reader = response.body!.getReader();
        const decoder = new TextDecoder();
        let buffer = "";

        try {
          while (true) {
            const { value, done } = await reader.read();
            if (done) break;

            buffer += decoder.decode(value, { stream: true });
            const lines = buffer.split("\n");
            buffer = lines.pop() ?? "";

            for (const rawLine of lines) {
              const line = rawLine.trim();
              if (!line.startsWith("data:")) continue;
              const data = line.slice(5).trim();
              if (!data) continue;

              try {
                const parsed = JSON.parse(data);
                // Anthropic delta events
                if (parsed.type === "content_block_delta" && parsed.delta?.type === "text_delta") {
                  const text = parsed.delta.text;
                  // Re-emit as OpenAI-compatible SSE so the client works unchanged
                  const openAIChunk = JSON.stringify({
                    choices: [{ delta: { content: text } }],
                  });
                  controller.enqueue(encoder.encode(`data: ${openAIChunk}\n\n`));
                }
                if (parsed.type === "message_stop") {
                  controller.enqueue(encoder.encode("data: [DONE]\n\n"));
                }
              } catch {
                // partial chunk — safe to skip
              }
            }
          }
        } finally {
          controller.close();
        }
      },
    });

    return new Response(transformedStream, {
      headers: {
        "Content-Type": "text/event-stream",
        "Cache-Control": "no-cache, no-transform",
        Connection: "keep-alive",
      },
    });
  } catch (error) {
    console.error("Nexus proxy error", error);
    return NextResponse.json({ error: "Proxy error" }, { status: 500 });
  }
}
