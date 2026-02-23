# Centralized Status Management

**Last Updated:** January 5, 2026

## Overview

Status updates flow through the **server as the single source of truth**, with the **client filtering stale messages** to prevent race conditions.

```
VSTO/Operations → Server (sends all status) → Client (filters stale) → UI
                     ↑
                User actions
```

## Who Does What

### Server (Source of Truth)

- ✅ Sends all status updates for operations
- ✅ Controls operation lifecycle
- ❌ Cannot filter based on client state (network lag)

**Sends:**

- Completion: `stream:end` then `idle`
- Cancellation: `stream:cancelled` then `idle`
- Errors: `stream:error` then `idle`
- Tool rejection: `idle` (after revert completes)
- Tool approval: `generating_llm` (LLM analyzes results)
- Operations: `reading_excel`, `writing_excel_format`, etc. (via VSTO progress)

### Client (Smart Filter)

- ✅ Filters stale "idle" during active operations
- ✅ Provides immediate feedback for user actions (ESC, Cancel, New Chat)
- ✅ Applies valid status updates to UI
- ❌ **Never sets status directly** - always waits for backend updates

**Rule:** Reject `idle` if streaming operation is active. Accept all other statuses.

## VSTO Progress Update Architecture

**Problem:** VSTO was sending progress updates asynchronously (fire-and-forget), so they arrived AFTER the HTTP response completed.

**Solution:** VSTO now awaits all progress updates before returning HTTP response.

**Before (Fire-and-Forget):**

```csharp
private void SendProgressUpdate(...) {
    _ = Task.Run(async () => {  // ❌ Fire-and-forget
        await _progressHttpClient.SendAsync(...);
    });
}
```

**After (Awaitable):**

```csharp
private async Task SendProgressUpdate(...) {
    await _progressHttpClient.SendAsync(...);  // ✅ Properly awaited
}
```

**Impact:**

- HTTP 200 response **truly** means all work is complete
- Server can immediately send "idle" without delays
- Backend remains Single Source of Truth
- No timing hacks or arbitrary delays needed

**Key Principle:** HTTP response completion = operation complete (including all progress updates)

## The Race Condition Problem

**Scenario:** User rapidly starts operations while network messages are delayed.

```
1. Operation A completes → Server sends "idle"
2. User immediately starts Operation B
3. Operation A's delayed "idle" arrives → Would clear B's status! ❌
```

**Why Client Must Filter:**

- Server doesn't know client's real-time state (network lag)
- Messages in flight can't be recalled
- Client is at the network edge with immediate local knowledge

## The Solution: Message Ordering + Race Protection

### 1. Message Ordering (Backend)

**Backend sends messages in specific order:**

```
1. Send "stream:end"   → Frontend clears currentAssistantMessageRef
2. Send "idle" status  → Frontend receives with ref=null ✅
```

**Why This Order Matters:**

- `stream:end` clears the ref that tracks active operations
- `idle` arrives after ref is cleared
- Simple boolean check: if ref is set, idle is stale

**Implementation:**

```typescript
// Safe helper methods ensure correct ordering
socket.endStream(workbookName, usage, model);
socket.cancelStream(workbookName, message, requestId);
socket.errorStream(workbookName, payload);

// Internally:
// endStream():
//   1. send("stream:end", ...)
//   2. sendStatus("idle")
// cancelStream():
//   1. send("stream:cancelled", ...)
//   2. sendStatus("idle")
// errorStream():
//   1. send("stream:error", ...)
//   2. sendStatus("idle")
```

### 2. Race Protection (Frontend)

**What:** A React ref tracking the current streaming operation's message ID.

```typescript
const currentAssistantMessageRef = useRef<string | null>(null);
// Lifecycle: null → "msg-ABC" (active) → null (idle)
```

**Protection Logic** (`useWebSocketStream.ts`):

```typescript
onStatusUpdate: (payload) => {
  if (isCancellingRef.current) return; // Ignore during cancellation

  // Simple race condition protection:
  // Backend sends stream:end BEFORE idle, so ref is always cleared first.
  // If idle arrives while ref is set, it's from a previous operation.
  if (
    payload.status === "idle" &&
    currentAssistantMessageRef.current !== null
  ) {
    return; // Stale "idle" from previous operation
  }

  setOperationStatus(payload.status); // Accept and display
}
```

**Result:**

- Legitimate "idle" always arrives when `ref === null` → Accept ✅
- Stale "idle" always arrives when `ref !== null` → Reject ❌
- Simple boolean check, no complex conditions needed

## Why Both Are Needed

**Message Ordering Alone:** Not enough because user can start Operation B before Operation A's messages arrive due to network lag.

**Race Protection Alone:** Would need complex checks to determine if idle is legitimate or stale.

**Together:** Simple and robust solution that handles all edge cases.

## User Actions: Immediate Feedback

User actions call `resetStatusIndicator()` for instant UI response:

- ESC key, Cancel button, New Chat → Immediate idle display
- Server confirms with its own "idle" shortly after

**Location:** `X21/web-ui/src/App.tsx`

## Testing Checklist

**Basic:**

- [ ] Normal completion shows idle
- [ ] ESC/Cancel shows immediate idle
- [ ] Stream errors show idle before error

**Race Condition (Critical):**

- [ ] Rapid operation switching - no stale statuses
- [ ] Check console for "Ignoring stale idle" message
- [ ] Network lag simulation (throttle in DevTools)

**Multi-Tab:**

- [ ] Multiple tabs don't interfere with each other
- [ ] Note: Current implementation may have limitations

## Key Takeaways

1. **Server = Source of Truth** - Sends all legitimate status updates
2. **Client = Smart Filter** - Validates which updates apply to current operation
3. **Message Ordering** - `stream:end` before `idle` simplifies protection
4. **VSTO Awaits Progress** - HTTP 200 means truly complete
5. **Network lag is real** - Messages arrive out of order, client handles this

## Related Files

**Server:**

- `X21/deno-server/src/services/websocket-manager.ts` - `endStream()` helper
- `X21/deno-server/src/router/index.ts` - WebSocket handler, cancellation, errors
- `X21/deno-server/src/stream/tool-logic.ts` - Streaming, tool execution
- `X21/vsto-addin/Services/ExcelApiService.cs` - Progress update implementation

**Client:**

- `X21/web-ui/src/hooks/useWebSocketStream.ts` - Status handler, race protection
- `X21/web-ui/src/App.tsx` - Manual resets

---
*January 5, 2026*
