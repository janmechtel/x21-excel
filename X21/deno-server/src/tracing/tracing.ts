// tracing.ts
import {
  Langfuse,
  LangfuseGenerationClient,
  LangfuseSpanClient,
  LangfuseTraceClient,
} from "langfuse";
import { createLogger } from "../utils/logger.ts";

const logger = createLogger("Tracing");

class TracingService {
  private langfuse: Langfuse;
  public traces = new Map<string, LangfuseTraceClient>(); // <traceId, trace>
  private generationCounter = new Map<string, number>(); // <traceId, generationCount>
  private generations = new Map<string, LangfuseGenerationClient>(); // <generationId, generation>
  private spans = new Map<string, LangfuseSpanClient>(); // <spanId, span>

  constructor() {
    const secretKey = Deno.env.get("LANGFUSE_SECRET_KEY");
    const publicKey = Deno.env.get("LANGFUSE_PUBLIC_KEY");
    const baseUrl = Deno.env.get("LANGFUSE_BASE_URL") ||
      "https://cloud.langfuse.com";
    const environment = Deno.env.get("LANGFUSE_TRACING_ENVIRONMENT") ||
      "development";

    if (!secretKey || !publicKey) {
      console.warn("[Tracing] LANGFUSE_API_KEY not set. Using no-op tracing.");
      this.langfuse = null as any;
    } else {
      this.langfuse = new Langfuse({
        secretKey,
        publicKey,
        baseUrl,
        environment,
      });
    }
  }

  startTrace(
    traceId: string,
    metadata: Record<string, any> = {},
    workbookName: string,
  ): any {
    if (!this.langfuse) {
      return null;
    }

    if (this.hasTrace(traceId)) {
      return this.getTrace(traceId);
    }

    this.generationCounter.set(traceId, 0);

    const sessionId = metadata.sessionId ?? "";
    const sessionIdMetadata = workbookName + "-" + sessionId;

    const trace: LangfuseTraceClient = this.langfuse.trace({
      id: traceId,
      name: metadata.name ?? "",
      userId: metadata.userEmail ?? "",
      input: metadata.input ?? "",
      sessionId: sessionIdMetadata,
      metadata: {
        ...metadata,
        timestamp: new Date().toISOString(),
      },
    });

    logger.info("trace created", traceId);
    this.traces.set(traceId, trace);
    return trace;
  }

  hasTrace(traceId: string): boolean {
    return this.traces.has(traceId);
  }

  getTrace(traceId: string): LangfuseTraceClient | null {
    if (!this.hasTrace(traceId)) {
      logger.warn("Trace not found", { traceId });
      return null;
    }
    return this.traces.get(traceId) as LangfuseTraceClient;
  }

  logGenerationStart(
    traceId: string,
    params: any,
    additionalMetadata?: any,
  ): string {
    const trace = this.getTrace(traceId);
    if (!trace) return "";

    const counter: number = this.generationCounterGetAndIncrease(traceId);

    // Determine the provider name for the generation name
    const providerName = additionalMetadata?.provider === "azure-openai"
      ? "Azure OpenAI"
      : "Anthropic";

    const lastMessage = Array.isArray(params?.messages)
      ? params.messages[params.messages.length - 1]
      : undefined;
    const generationInput = extractInputForGeneration(lastMessage?.content);
    const toolNames = extractToolNames(params?.tools);
    const systemMessage = extractSystemMessage(params?.system);
    const modelName = additionalMetadata?.model || params?.model || "";

    logger.info("Langfuse generation start", {
      traceId,
      provider: providerName,
      model: modelName,
    });

    const generation: LangfuseGenerationClient = trace.generation({
      name: `${providerName} API Call #${counter}`,
      input: generationInput,
      model: modelName,
      metadata: {
        lastMessageRole: lastMessage?.role,
        messageCount: params.messages.length,
        systemMessage,
        toolCount: toolNames.length,
        toolNames: toolNames.join(", "),
        modelName,
        ...additionalMetadata,
      },
    });

    this.generations.set(generation.id, generation);
    return generation.id;
  }

  logGenerationEnd(generationId: string, params: any) {
    const generation = this.generations.get(generationId);
    if (!generation) return;
    generation.end({
      output: params.output,
      metadata: params.metadata,
      usage: {
        input: params.metadata.inputTokens,
        output: params.metadata.outputTokens,
      },
    });
  }

  private generationCounterGetAndIncrease(traceId: string) {
    const generationCount = this.generationCounter.get(traceId) || 0;
    this.generationCounter.set(traceId, generationCount + 1);
    return generationCount;
  }

  logToolCallStart(traceId: string, tool: string, payload: any): string {
    logger.info("reading trace", { traceId });
    const trace = this.getTrace(traceId);

    if (!trace) return "";

    logger.info("creating span", { tool: tool });
    const span = trace.span({
      name: tool,
      input: payload,
    });

    this.spans.set(span.id, span);
    return span.id;
  }

  logToolCallEnd(spanId: string, payload: any) {
    const span = this.getSpan(spanId);
    if (!span) return;
    span.update({
      output: payload,
    });
    span.end();
    this.spans.delete(spanId);
  }

  logUserCancellationSpan(traceId: string): string {
    logger.info("Creating user cancellation span", { traceId });
    const trace = this.getTrace(traceId);
    if (!trace) return "";
    const span = trace.span({
      name: "user_cancellation",
    });
    return span.id;
  }

  logToolRejectionSpan(
    traceId: string,
    toolId: string,
    toolName: string,
    userMessage: string,
  ): string {
    logger.info("Creating tool rejection span", { traceId, toolId, toolName });
    const trace = this.getTrace(traceId);
    if (!trace) return "";
    const span = trace.span({
      name: "tool_rejection",
      input: {
        toolId: toolId,
        toolName: toolName,
        userMessage: userMessage,
      },
    });
    this.spans.set(span.id, span);
    return span.id;
  }

  logToolRejectionSpanEnd(spanId: string, metadata: Record<string, any> = {}) {
    if (!spanId) return;
    try {
      const span = this.getSpan(spanId);
      if (!span) return;
      span.update({
        output: metadata,
      });
      span.end();
      this.spans.delete(spanId);
    } catch (error) {
      logger.error("Error ending tool rejection span", error);
    }
  }

  logToolApprovalSpan(
    traceId: string,
    toolId: string,
    toolName: string,
  ): string {
    logger.info("Creating tool approval span", { traceId, toolId, toolName });
    const trace = this.getTrace(traceId);
    if (!trace) return "";
    const span = trace.span({
      name: "tool_approval",
      input: {
        toolId: toolId,
        toolName: toolName,
      },
    });
    this.spans.set(span.id, span);
    return span.id;
  }

  logToolApprovalSpanEnd(spanId: string, metadata: Record<string, any> = {}) {
    if (!spanId) return;
    try {
      const span = this.getSpan(spanId);
      if (!span) return;
      span.update({
        output: metadata,
      });
      span.end();
      this.spans.delete(spanId);
    } catch (error) {
      logger.error("Error ending tool approval span", error);
    }
  }

  logCompactingStart(traceId: string, metadata: Record<string, any>): string {
    logger.info("Starting compacting span", { traceId });
    const trace = this.getTrace(traceId);

    if (!trace) return "";

    logger.info("Creating compacting span");
    const span = trace.span({
      name: "conversation_compacting",
      input: metadata,
    });

    this.spans.set(span.id, span);
    return span.id;
  }

  logCompactingEnd(spanId: string, metadata: Record<string, any>) {
    const span = this.getSpan(spanId);
    if (!span) return;

    span.end({
      output: metadata,
    });
    this.spans.delete(spanId);
  }

  getSpan(spanId: string): LangfuseSpanClient | null {
    if (!this.spans.has(spanId)) {
      logger.warn("Span not found", { spanId });
      return null;
    }
    return this.spans.get(spanId) as LangfuseSpanClient;
  }

  startSpan(
    traceId: string,
    params: { name: string; input?: any; metadata?: any },
  ): string {
    logger.info("creating span", { name: params.name, traceId });
    const trace = this.getTrace(traceId);

    if (!trace) return "";

    const span = trace.span({
      name: params.name,
      input: params.input,
      metadata: params.metadata,
    });

    this.spans.set(span.id, span);
    return span.id;
  }

  endSpan(
    spanId: string,
    params: { output?: any; success?: boolean; error?: string },
  ) {
    const span = this.getSpan(spanId);
    if (!span) return;

    span.update({
      output: params.output,
      metadata: {
        success: params.success,
        error: params.error,
      },
    });
    span.end();
    this.spans.delete(spanId);
  }

  endTrace(traceId: string, metadata: Record<string, any> = {}) {
    const trace = this.getTrace(traceId);
    if (!trace) return;

    trace.update({
      metadata: metadata,
    });

    this.traces.delete(traceId);
  }

  sendScore(traceId: string, score: number): void {
    if (!this.langfuse) {
      return;
    }

    const name = "score";

    this.langfuse.score({
      traceId,
      name,
      value: score,
    });
  }

  sendFeedback(traceId: string, feedback: string): void {
    if (!this.langfuse) {
      return;
    }

    const name = "feedback";

    this.langfuse.score({
      traceId,
      name,
      value: feedback,
    });
  }
}

// Anthropic system can be a string or an array of text blocks; normalize for tracing.
function extractSystemMessage(system: unknown): string {
  if (typeof system === "string") return system;
  if (Array.isArray(system)) {
    const parts: string[] = [];
    let unknownCount = 0;

    for (const item of system) {
      if (typeof item === "string") {
        parts.push(item);
        continue;
      }
      if (item && typeof item === "object") {
        const text = (item as { text?: unknown }).text;
        if (typeof text === "string") {
          parts.push(text);
          continue;
        }
      }
      unknownCount += 1;
    }

    if (unknownCount > 0) {
      logger.warn("Unexpected system message parts", {
        unknownCount,
        total: system.length,
      });
    }

    return parts.join("\n").trim();
  }

  if (system !== null && system !== undefined) {
    logger.warn("Unexpected system message type", { type: typeof system });
  }
  return "";
}

function extractToolNames(tools: unknown): string[] {
  if (!Array.isArray(tools)) return [];
  return tools
    .map((tool) => {
      if (!tool || typeof tool !== "object") return "";
      const name = (tool as { name?: unknown }).name;
      return typeof name === "string" ? name : "";
    })
    .filter((name) => name.length > 0);
}

function extractInputForGeneration(content: unknown): unknown {
  // Avoid logging giant payloads (e.g., base64 documents). Prefer a concise summary.
  if (typeof content === "string") return content;

  if (Array.isArray(content)) {
    const textParts: string[] = [];
    const attachmentTypes: string[] = [];

    for (const block of content) {
      if (!block || typeof block !== "object") continue;
      const blockType = (block as { type?: unknown }).type;
      if (blockType === "text") {
        const text = (block as { text?: unknown }).text;
        if (typeof text === "string") {
          textParts.push(text);
        }
      } else if (blockType === "document" || blockType === "image") {
        const mediaType = (block as { source?: { media_type?: unknown } })
          ?.source?.media_type;
        attachmentTypes.push(
          typeof mediaType === "string" ? mediaType : String(blockType),
        );
      } else if (blockType === "tool_use") {
        const name = (block as { name?: unknown }).name;
        const id = (block as { id?: unknown }).id;
        textParts.push(`[tool_use:${String(name || id || "unknown")}]`);
      }
    }

    const summary: Record<string, unknown> = {};
    const text = textParts.join("\n").trim();
    if (text.length > 0) summary.text = text;
    if (attachmentTypes.length > 0) {
      summary.attachments = {
        count: attachmentTypes.length,
        types: Array.from(new Set(attachmentTypes)),
      };
    }

    if (Object.keys(summary).length === 0) {
      return "[non-text content]";
    }
    return summary;
  }

  if (content === null || content === undefined) return "";
  try {
    return JSON.stringify(content);
  } catch (_err) {
    return String(content);
  }
}

export const tracing = new TracingService();
