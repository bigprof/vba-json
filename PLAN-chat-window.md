# Plan: Chat Window with Tool-Calling Loop

## What Already Exists

The codebase already has all the building blocks:

- **`OpenAI.cls`** — Full Chat Completions & Responses API client
- **`OpenAIHelpers.bas`** — Message builders, JSON helpers, response extractors (including `OpenAIExtractToolCalls`, `OpenAIExtractFinishReason`)
- **`JsonData.cls`** / **`Json.bas`** — JSON parsing & traversal
- **`OpenAITester.bas`** — A working demo of the exact tool-calling loop in `TestOpenAI_FunctionToolCall_RequestResponseLoop`

What's missing is the **GUI layer** and **background processing** to drive the loop without freezing the UI.

---

## Architecture

```
┌──────────────────────────────────────────────────┐
│  frmChat.frm (Chat Window)                       │
│  ┌────────────────────────────────────────────┐  │
│  │ RichTextBox (chat history, read-only)      │  │
│  │  User: "What's the weather in Boston?"     │  │
│  │  🤖: "Let me check..."                     │  │
│  │  🔧 get_weather("Boston") → 72°F           │  │
│  │  🤖: "It's 72°F and partly cloudy."        │  │
│  └────────────────────────────────────────────┘  │
│  [Text Input___________________] [Send]          │
│  Status: Thinking...                             │
│  Timer: tmrLoop (100ms)                          │
└──────────────┬───────────────────────────────────┘
               │ owns & drives
               ▼
┌──────────────────────────────────────────────────┐
│  ChatEngine.cls (State Machine)                  │
│  ┌────────────────────────────────────────────┐  │
│  │ States: Idle → Calling → Processing → Done │  │
│  │ Holds: Messages Collection                 │  │
│  │ Holds: ToolsJson string                    │  │
│  │ Events: StatusChanged, MessageReceived,    │  │
│  │         ToolExecuted, ErrorOccurred, Done  │  │
│  └──────────────┬─────────────────────────────┘  │
│                 │ uses                             │
│  ┌──────────────▼─────────────────────────────┐  │
│  │ ToolRegistry.cls                           │  │
│  │  Register(toolName, handler)               │  │
│  │  Execute(toolName, args) → result          │  │
│  │  BuildToolsJson() → "[{...},{...}]"        │  │
│  └────────────────────────────────────────────┘  │
└──────────────────────────────────────────────────┘
```

---

## New Files to Create

| File | Purpose |
|---|---|
| **`ChatEngine.cls`** | Core loop logic — state machine that calls API, detects tool calls, invokes `ToolRegistry`, and feeds results back until the model returns a final text response |
| **`ToolRegistry.cls`** | Maps tool names → handler objects; generates the `tools` JSON array; dispatches execution |
| **`IToolHandler.cls`** | Abstract interface: `Execute(args As String) As String` + `GetName() As String` |
| **`AsyncHttpRequest.cls`** | Non-blocking HTTP wrapper using `WithEvents` on `MSXML2.XMLHTTP60`; fires `Completed` event when response arrives (Layer 2 mitigation) |
| **`frmChat.frm`** | VB6 Form with chat display (`RichTextBox`), input box, Send button, and a `Timer` control |

**Modified files:**

| File | Change |
|---|---|
| **`OpenAI.cls`** | Extract `BuildRequestBody` from `PostJson` so the body+headers can be built synchronously and sent asynchronously. Old `PostJson` retained for backward compat. |

---

## Background Processing Strategy

VB6 is single-threaded — the HTTP call in `OpenAI.PostJson` blocks. To keep the UI alive:

**Use a Timer-driven state machine (recommended):**

1. User clicks **Send** → `ChatEngine.StartConversation(userMessage)` → enables `tmrLoop`
2. `tmrLoop` fires every 100ms → calls `ChatEngine.ContinueLoop()`
3. `ContinueLoop()` advances one step at a time:
   - **If state = Idle**: do nothing (Timer disabled)
   - **If state = ReadyToCall**: make the sync API call (brief freeze, typically 2-5s for tool calls — acceptable). Extract response.
   - **If finish_reason = `"tool_calls"`**: extract tool calls, execute each via `ToolRegistry`, add results to Messages, raise `ToolExecuted` event for UI, stay in ReadyToCall (Timer will trigger next iteration)
   - **If finish_reason = `"stop"`**: extract text, add assistant message, raise `MessageReceived`, disable Timer → Idle
4. The Timer keeps the UI responsive between API calls (you can type, scroll, etc.)

```
                 ┌──────────┐
    Send clicked │   Idle   │
         ───────►│ (timer   │◄────────────── Done ──────┐
                 │  off)    │                            │
                 └────┬─────┘                            │
                      │ enable timer                     │
                 ┌────▼─────┐                            │
                 │ ReadyTo   │◄──── tool calls found ────┤
                 │  Call     │                           │
                 └────┬─────┘                            │
                      │ API call (sync, 2-5s)            │
                 ┌────▼─────┐                            │
                 │ Evaluate  │                            │
                 │ Response  │                            │
                 └────┬─────┘                            │
                      │                                  │
          ┌───────────┴───────────┐                      │
          │ finish_reason?        │                      │
          ▼                       ▼                      │
    "tool_calls"              "stop"/other               │
    - extract tools           - extract text             │
    - execute each            - fire MessageReceived ────┘
    - add results to Messages
    - fire ToolExecuted
    - stay in ReadyToCall
```

---

## Implementation Steps (in order)

### Phase 1 — Core classes (no UI)

1. **`IToolHandler.cls`** — Stub class with `Execute(argsJson As String) As String` and `GetName() As String`

2. **`ToolRegistry.cls`** — 
   - Internal `Scripting.Dictionary` mapping names → handlers
   - `Register(handler As IToolHandler)`
   - `Execute(name, args) As String` — dispatches
   - `BuildToolsJson() As String` — iterates registered tools, builds the OpenAI `tools` array in Chat Completions format (`{"type":"function","function":{...}}`)

3. **`ChatEngine.cls`** — 
   - Properties: `OpenAI` client, `ToolRegistry`, `SystemPrompt`, `UseAsync` (Boolean)
   - Internal: `Messages` collection, state enum (`Idle | ReadyToCall | WaitingForResponse | Done`), `mLastResponse`
   - `WithEvents` on `AsyncHttpRequest` (created when `UseAsync = True`)
   - `StartConversation(userMessage)` — clears history, adds system+user messages, transitions to `ReadyToCall`
   - `ContinueLoop()` — the single-step function called by the Timer:
     - `ReadyToCall`: if async, fire HTTP request → transition to `WaitingForResponse`; if sync, call `OpenAI.CreateChatCompletion` → transition to `EvaluateResponse`
     - `WaitingForResponse`: call `DoEvents` only (response will arrive via event)
     - `Done` / `Idle`: no-op
   - `EvaluateResponse(resp)` — extracts `finish_reason`, dispatches to tool-call processing or text delivery
   - Events: `StatusChanged(ByVal Status As String)`, `MessageReceived(ByVal Text As String)`, `ToolExecuted(ByVal Name As String, ByVal Args As String, ByVal Result As String)`, `ErrorOccurred(ByVal Msg As String)`, `Done()`

### Phase 2 — Async HTTP (Layer 2)

4. **`AsyncHttpRequest.cls`** — 
   - `Private WithEvents mHttp As MSXML2.XMLHTTP60`
   - `Start(ByVal method As String, ByVal url As String, ByVal body As String, ByVal headers As Collection)`
   - `Cancel()` — abort in-flight request
   - `IsBusy As Boolean` — true while request is in flight
   - Events: `Completed(ByVal status As Long, ByVal responseText As String)`, `Error(ByVal number As Long, ByVal description As String)`

5. **Refactor `OpenAI.cls`** — Extract `BuildRequestBody` method:
   - `BuildRequestBody(RelativeUrl As String, body As String, ByRef outUrl As String, ByRef outHeaders As Collection)` — builds URL with base, sets auth/org/project headers, returns them for `AsyncHttpRequest` to use. Does NOT call `send`.
   - Old `PostJson` delegates to `BuildRequestBody` + its existing sync `send` — zero behavior change for existing callers.
   - `CreateChatCompletion` calls `BuildRequestBody` + `PostJson` as before.

### Phase 3 — UI

6. **`frmChat.frm`** — 
   - `RichTextBox` (or `TextBox` multiline) for chat history (read-only)
   - `TextBox` for user input, `CommandButton` for Send
   - `Timer` control (`tmrLoop`, `Interval=100`, `Enabled=False`)
   - `Timer` control (`tmrWatchdog`, `Interval=500`) — pumps `DoEvents` once when engine is busy, to prevent OS ghosting
   - `Form_Load`: create `ChatEngine`, register example tools, set system prompt
   - `cmdSend_Click`: disable input, call `ChatEngine.StartConversation`, enable `tmrLoop` and `tmrWatchdog`
   - `tmrLoop_Timer`: call `ChatEngine.ContinueLoop()`. If engine state is `Done` or `Idle`, disable both timers and re-enable input.
   - Handle `ChatEngine` events:
     - `StatusChanged` → update `lblStatus` caption + `Me.Refresh` + `DoEvents`
     - `MessageReceived` → append `"🤖 " & text` to chat display, scroll to end
     - `ToolExecuted` → append `"🔧 " & name & "(" & args & ") → " & result` in muted color
     - `ErrorOccurred` → append `"❌ " & msg` in red, disable timers, re-enable input

7. **Example Tools** — Two `IToolHandler` implementations (weather, coordinates) using the mock data from `ExecuteToolFunction` in the existing tests. These demonstrate the pattern; real tools can be swapped in later.

8. **Entry Point** — Add `TestChatWindow()` to `OpenAITester.bas` that does `frmChat.Show vbModal`

---

## API Choice

Use the **Chat Completions API** (`CreateChatCompletion`) rather than the Responses API because:
- The message-based model (user → assistant → tool → assistant) maps naturally to a chat UI
- `TestOpenAI_FunctionToolCall_RequestResponseLoop` already proves the pattern works
- The Responses API's stateful `previous_response_id` chaining is harder to interrupt mid-loop for UI updates

### Side note: What changes if we switch to the Responses API later

The Responses API is semantically different from Chat Completions in ways that touch several files. Here is exactly what would need to change:

**1. `ChatEngine.cls` — history model flips entirely**

| Concept | Chat Completions (current) | Responses API (future) |
|---|---|---|
| Conversation history | `Collection` of message JSON strings built up manually | The server holds history via `previous_response_id`; the client only sends the *new input items* |
| Multi-turn | Append assistant + tool messages to `Messages`, resend entire history each call | Only send `function_call_output` items for the current tool calls; the server remembers prior turns |
| Start new chat | Clear `Messages`, add developer + user | Omit `previous_response_id` (first call); optionally set `instructions` |
| Tool call format | `{"type":"function","function":{"name":"…","arguments":"…"}}` with `tool_call_id` | `{"type":"function_call","call_id":"…","name":"…","arguments":"…"}` (flat, no wrapper) |
| Tool result format | `{"role":"tool","tool_call_id":"…","content":"…"}` | `{"type":"function_call_output","call_id":"…","output":"…"}` |
| Extracting text | `choices.0.message.content` | `output_text` convenience field, or `output[].content[].text` |
| Extracting tool calls | `choices.0.message.tool_calls` array | Iterate `output[]` items where `type = "function_call"` |
| Finish detection | `choices.0.finish_reason = "tool_calls"` vs `"stop"` | `status = "completed"` and no `function_call` items in `output[]` |
| Reasoning | `reasoning_effort` parameter in Chat Completions | `reasoning` object `{"effort":"…","summary":"…"}` |

**2. `OpenAI.cls` — already done**

`CreateResponse` and `CreateResponseSimple` are already implemented. No new endpoint work needed.

**3. `OpenAIHelpers.bas` — already done**

`ResponsesExtractText`, `ResponsesExtractStatus`, `ResponsesExtractToolCalls`, `ResponsesBuildFunctionCallResult`, `ResponsesTextFormatText/JsonObject/JsonSchema`, `ResponsesReasoning`, `ResponsesToolChoiceAuto/None/Required` are all already written. No new helpers needed.

**4. `ToolRegistry.cls` — no change**

The tool execution interface is API-agnostic. The only difference is `BuildToolsJson()` would generate the Responses flattened format (`{"type":"function","name":"…",…}` without the `"function"` wrapper) instead of the Chat Completions nested format. This is a one-line toggle — add a `BuildToolsJsonResponses()` method.

**5. `ChatEngine.cls` — state machine changes**

- Drop the `Messages` collection; store `mPreviousResponseId As String` instead.
- `StartConversation`: first call sends user input as `InputItems` (string) + `instructions`; capture the returned `id`.
- Loop iterations: each subsequent call passes `PreviousResponseId:=mPreviousResponseId` and `InputItems:=newToolResults` (a `Collection` of `function_call_output` items). Re-capture `id`.
- `EvaluateResponse`: iterate `output[]` for `function_call` items instead of checking `finish_reason`. If none found, extract `output_text`.
- The `Done` / `ReadyToCall` / `WaitingForResponse` states remain identical.
- The `UseAsync` flag and `AsyncHttpRequest` integration work identically — HTTP is HTTP regardless of endpoint.

**6. `frmChat.frm` — zero changes**

The form only sees events (`MessageReceived`, `ToolExecuted`, `StatusChanged`, `ErrorOccurred`, `Done`). The API choice is invisible to the UI.

**Estimated effort to switch:** ~1 hour. The heavy lifting (async HTTP, tool registry, UI) doesn't change. Only the history management in `ChatEngine` and the JSON format in `ToolRegistry.BuildToolsJson` need updating, and the test `TestResponsesToolCalling` in `OpenAITester.bas` already proves the Responses loop works.

---

## Key Risk: Blocking HTTP — Mitigation

### Problem

`OpenAI.PostJson` uses `MSXML2.XMLHTTP60.send` which blocks the calling thread until the HTTP response arrives. API calls to OpenAI (especially when the model is "thinking" with reasoning) can take 5–30 seconds. During that time the VB6 message pump is frozen: the form stops repainting, appears "(Not Responding)", and user input is lost.

### Mitigation strategy

Two layers — a **short-term safety net** (always present, zero architectural change) and a **proper async refactor** (the real fix).

---

### Layer 1 — Immediate: `DoEvents` pulse + status heartbeat

Before and after every HTTP operation, the form pumps pending messages so the UI surface stays painted.

**What changes:**
- `frmChat` exposes a `PumpMessages()` helper that calls `DoEvents` in a tight guard (`Do While DoEvents(): Loop` is NOT used — that causes reentrancy).
- Before `ChatEngine` calls `OpenAI.CreateChatCompletion`, it raises `StatusChanged("Calling API…")` and the form handler calls `Me.Refresh` / `DoEvents` (once).
- After the call returns, `StatusChanged` fires again so the status label updates.
- `frmChat` also runs a 500 ms "watchdog" Timer that calls `DoEvents` unconditionally while the engine is busy — this prevents the OS from ghosting the window even if the main Timer tick is consumed by the engine.

**Net effect:** The UI still freezes during the actual `send` call, but the window is painted correctly when the freeze starts and ends, and the OS won't flag it as hung unless the call exceeds ~5 seconds. For most Chat Completions calls (especially cheap models like `gpt-4o-mini`), response time is 1–4 seconds, so this is sufficient for v1.

---

### Layer 2 — Proper: Async HTTP with `WithEvents`

For production use (or when using reasoning models that take 10–30 s), `PostJson` is refactored to support asynchronous `XMLHTTP` requests.

**Architecture change:**

```
┌──────────────────────────────────────────┐
│  AsyncHttpRequest.cls  (NEW)             │
│  ┌────────────────────────────────────┐  │
│  │ Private WithEvents mHttp As        │  │
│  │   MSXML2.XMLHTTP60                │  │
│  │                                    │  │
│  │ Sub Start(method, url, body)       │  │
│  │   mHttp.Open method, url, True     │  │  ← async!
│  │   mHttp.send body                  │  │
│  │ End Sub                            │  │
│  │                                    │  │
│  │ Event Completed(status, body)      │  │
│  │ Event Error(number, desc)          │  │
│  │                                    │  │
│  │ Private Sub mHttp_onReadyState…    │  │
│  │   If mHttp.readyState = 4 Then     │  │
│  │     RaiseEvent Completed(…)        │  │
│  │   End If                           │  │
│  │ End Sub                            │  │
│  └────────────────────────────────────┘  │
└──────────────┬───────────────────────────┘
               │ WithEvents in ChatEngine
               ▼
┌──────────────────────────────────────────┐
│  ChatEngine.cls (updated)                │
│  ┌────────────────────────────────────┐  │
│  │ Private WithEvents mAsyncHttp As   │  │
│  │   AsyncHttpRequest                │  │
│  │                                    │  │
│  │ State "WaitingForResponse"         │  │  ← new state
│  │ Timer sees this state → skips      │  │
│  │ (just pumps DoEvents)             │  │
│  │                                    │  │
│  │ mAsyncHttp_Completed:              │  │
│  │   parse response → Evaluate state  │  │
│  │   set state = ReadyToCall          │  │
│  │   (Timer picks up next iteration)  │  │
│  └────────────────────────────────────┘  │
└──────────────────────────────────────────┘
```

**How it works step-by-step:**

1. `ChatEngine` enters `ReadyToCall` state.
2. Instead of calling `OpenAI.PostJson` (sync), it calls `OpenAI.BuildRequestBody()` to get the JSON body, then calls `mAsyncHttp.Start("POST", url, body)`.
3. State transitions to `WaitingForResponse`. Control returns to the message pump immediately.
4. `tmrLoop` fires every 100 ms, sees `WaitingForResponse`, and simply calls `DoEvents` to keep the UI alive. No API work happens on the Timer.
5. When the HTTP response arrives, `mHttp_onReadyStateChange` fires → `AsyncHttpRequest` raises `Completed(status, body)`.
6. `ChatEngine.mAsyncHttp_Completed` parses the JSON, extracts `finish_reason` and text/tool_calls, then either:
   - Transitions to `ReadyToCall` (if tool calls found — Timer drives next iteration)  
   - Transitions to `Done` (if `stop` — raises `MessageReceived`).

**Impact on `OpenAI.cls`:**
`PostJson` is split into two methods:
- `BuildRequestBody(RelativeUrl, body) As String` — builds the URL and sets headers (purely synchronous, no network).
- The old `PostJson` is kept for backward compatibility (existing tests). The new `AsyncHttpRequest` takes over the actual `send`.

**Why not use `WinHttp.WinHttpRequest`?**  
`MSXML2.XMLHTTP60` with `WithEvents` is the most reliable async pattern in VB6. `WinHttpRequest`'s event model has known quirks with VB6's COM event sinks. `XMLHTTP60` is stable and already a dependency of the project.

---

### Updated state machine (with async)

```
                ┌──────────┐
   Send clicked │   Idle   │
        ───────►│ (timer   │◄────────────── Done ─────────┐
                │  off)    │                               │
                └────┬─────┘                               │
                     │ enable timer                        │
                ┌────▼─────┐                               │
                │ ReadyTo   │◄─── tool calls found ────────┤
                │  Call     │                              │
                └────┬─────┘                               │
                     │ fire async HTTP request             │
                ┌────▼──────────┐                          │
                │ WaitingFor     │  ← NEW (non-blocking)   │
                │  Response      │                          │
                └────┬──────────┘                          │
                     │ AsyncHttpRequest.Completed event    │
                ┌────▼─────┐                               │
                │ Evaluate  │                               │
                │ Response  │                               │
                └────┬─────┘                               │
                     │                                     │
         ┌───────────┴───────────┐                         │
         │ finish_reason?        │                         │
         ▼                       ▼                         │
   "tool_calls"             "stop"/other                  │
   - extract tools          - extract text                │
   - execute each           - fire MessageReceived ───────┘
   - add results to Messages
   - fire ToolExecuted
   - set state ReadyToCall
```

### Implementation order for async

1. Build and test **Layer 1** first (the `DoEvents` pulse) — this is 10 lines of code and makes v1 usable immediately.
2. Build `AsyncHttpRequest.cls` with `WithEvents`, test with a simple GET.
3. Add `OpenAI.BuildRequestBody` (refactor `PostJson` without breaking it).
4. Add `WaitingForResponse` state to `ChatEngine`, wire up `WithEvents`.
5. The Timer `tmrLoop` now handles an additional state; no other code changes needed.
6. Old sync path retained behind a `ChatEngine.UseAsync = True/False` flag so you can switch between v1 (sync) and v2 (async) at runtime for comparison.
