# VB6 OpenAI + JsonData

A small VB6 client for working with JSON and the OpenAI Chat Completions and Responses APIs.

---

## JSON Library

The library ships two files that together form a lightweight JSON parser and serializer for VB6 / VBA:

| File | Purpose |
|------|---------|
| `Json.bas` | `ParseJSON()` function – parses a JSON string into a `JsonData` tree |
| `JsonData.cls` | Class representing a single JSON value (object, array, or scalar) |

### Parsing a JSON string

```vb
Dim root As JsonData
root = ParseJSON("{""name"":""Alice"",""age"":30}")

If root.IsValid Then
    Debug.Print root.GetObjectItem("name").ScalarValue  ' Alice
    Debug.Print root.GetObjectItem("age").ScalarValue   ' 30
End If
```

`ParseJSON` returns a `JsonData` object. If the string cannot be parsed, `IsValid` returns `False`.

### JsonData – type constants

| Constant | Value | Meaning |
|----------|-------|---------|
| `JSONDATATYPE_INVALID` | 0 | Parse failed or node not found |
| `JSONDATATYPE_SCALAR` | 1 | A string, number, boolean, or null value |
| `JSONDATATYPE_ARRAY` | 2 | A JSON array |
| `JSONDATATYPE_OBJECT` | 3 | A JSON object |

### JsonData – properties and methods

#### Inspecting the type

| Member | Returns | Description |
|--------|---------|-------------|
| `IsValid` | `Boolean` | `True` when `DataType <> JSONDATATYPE_INVALID` |
| `DataType` | `Long` | One of the `JSONDATATYPE_*` constants above |
| `IsScalar` | `Boolean` | `True` when the node is a scalar value |
| `IsArray` | `Boolean` | `True` when the node is an array |
| `IsObject` | `Boolean` | `True` when the node is an object |

#### Working with scalar values

```vb
Dim node As JsonData
Set node = root.GetChildByPath("user.active")

If node.IsScalar Then
    If IsNull(node.ScalarValue) Then
        Debug.Print "null"
    Else
        Debug.Print CStr(node.ScalarValue)   ' True / False / number / string
    End If
End If
```

| Member | Returns | Description |
|--------|---------|-------------|
| `ScalarValue` | `Variant` | The raw VB `Variant` for a scalar node. May be `Null` for JSON `null`. |

#### Working with arrays

```vb
Dim arr As JsonData
Set arr = root.GetChildByPath("items")

Dim i As Long
For i = 0 To arr.ArrayLength - 1
    Debug.Print arr.GetArrayItem(i).ScalarValue
Next i
```

| Member | Returns | Description |
|--------|---------|-------------|
| `ArrayLength` | `Long` | Number of elements in the array |
| `GetArrayItem(index As Long)` | `JsonData` | Element at the given 0-based index. Returns an invalid `JsonData` when out of range. |

#### Working with objects

```vb
Dim obj As JsonData
Set obj = root.GetChildByPath("user")

Dim key As Variant
For Each key In obj.ObjectKeys
    Debug.Print key & " = " & obj.GetObjectItem(CStr(key)).ScalarValue
Next key
```

| Member | Returns | Description |
|--------|---------|-------------|
| `ObjectHasKeys` | `Boolean` | `True` when the object has at least one key |
| `ObjectKeys` | `String()` | Array of key names in declaration order |
| `GetObjectItem(key As String)` | `JsonData` | Value for the given key. Returns an invalid `JsonData` when the key does not exist. |

#### Navigating with a dot-path

`GetChildByPath` lets you navigate deeply nested structures without intermediate variables. Use `.` as the path separator and integer indices for arrays.

```vb
' JSON: {"user":{"name":"Alice"},"items":[{"value":123},{"value":456},"hello"]}
Debug.Print root.GetChildByPath("user.name").ScalarValue        ' Alice
Debug.Print root.GetChildByPath("items.0.value").ScalarValue    ' 123
Debug.Print root.GetChildByPath("items.2").ScalarValue          ' hello
```

If any part of the path is missing or the type is not traversable, the method returns an invalid `JsonData` (i.e. `IsValid = False`) rather than raising an error.

| Member | Returns | Description |
|--------|---------|-------------|
| `GetChildByPath(path As String)` | `JsonData` | Dot-separated path. Use numeric keys for array indices. Returns `Me` for an empty path. |

### Serializing back to JSON

Call `ToJSON` on any `JsonData` node to produce a JSON string:

```vb
' Pretty-printed with two-space indent
Debug.Print root.ToJSON("  ")

' Compact (no indentation, no newlines)
Debug.Print root.ToJSON("", 0, "")

' Default: tab-indented
Debug.Print root.ToJSON()
```

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `IndentWith` | `String` | `vbTab` | String repeated once per nesting level |
| `Depth` | `Long` | `0` | Starting depth (normally leave at 0) |
| `NewLineSequence` | `String` | `vbNewLine` | Line separator between tokens |

### JSON helper functions (`OpenAIHelpers.bas`)

These functions make it easy to build JSON strings manually without a serializer.

#### `JsonString(value As String) As String`

Encodes a VB string as a properly escaped JSON string literal (including `\n`, `\t`, Unicode escapes, etc.).

```vb
Dim s As String
s = JsonString("Hello ""World""")
' Result: "Hello \"World\""
```

#### `JsonBoolean(value As Boolean) As String`

Returns `"true"` or `"false"`.

```vb
Debug.Print JsonBoolean(True)   ' true
Debug.Print JsonBoolean(False)  ' false
```

#### `JsonNumber(value As Double) As String`

Formats a number as a valid JSON number, guaranteeing a leading zero before the decimal point (e.g. `0.5` instead of `.5`).

```vb
Debug.Print JsonNumber(0.5)   ' 0.5
Debug.Print JsonNumber(-0.2)  ' -0.2
```

#### `CollectionToJsonArray(Items As Collection) As String`

Converts a `Collection` of already-serialized JSON values into a JSON array string.

```vb
Dim parts As New Collection
parts.Add JsonString("red")
parts.Add JsonString("green")
parts.Add JsonString("blue")

Debug.Print CollectionToJsonArray(parts)
' ["red","green","blue"]
```

---

## Notes — Chat Completions API

- OpenAI recommends the **Responses API** for new text-generation builds, but Chat Completions remains available. ([platform.openai.com](https://platform.openai.com/docs/guides/chat-completions?utm_source=openai))
- The examples above use the Chat Completions API. The same client now also supports the Responses API — see the next section.
- For Chat Completions, `response_format` is a JSON object, `tools` is a JSON array, `tool_choice` is a string or JSON object, and `metadata` is a JSON object. `parallel_tool_calls` is an optional boolean, and `temperature` defaults to `1`. ([platform.openai.com](https://platform.openai.com/docs/api-reference/chat/create-chat-completion?utm_source=openai))
- With GPT-5-family Chat Completions usage, some combinations are model-dependent. In practice, unsupported fields should be omitted rather than sent as invalid placeholders. This matches the API reference behavior for optional request fields. ([platform.openai.com](https://platform.openai.com/docs/api-reference/chat/create-chat-completion?utm_source=openai))

## Setup

Set your API key in an environment variable:

```bat
set OPENAI_API_KEY=your_api_key_here
```

Then in VB6:

```vb
Dim ai As OpenAI
Set ai = New OpenAI
ai.ApiKey = Environ$("OPENAI_API_KEY")
```

## Simple text example

```vb
Public Sub TestOpenAISimple()
    Dim client As OpenAI
    Dim result As JsonData
    
    Set client = New OpenAI
    client.ApiKey = Environ$("OPENAI_API_KEY")
    
    Set result = client.CreateChatCompletionSimple( _
        "gpt-5.4", _
        "You are a helpful assistant.", _
        "Write a short haiku about VB6 and APIs." _
    )
    
    Debug.Print OpenAIExtractText(result)
End Sub
```

## Multi-turn chat example

```vb
Public Sub TestOpenAI_MultiTurn()
    Dim ai As OpenAI
    Dim resp As JsonData
    Dim Messages As Collection
    
    On Error GoTo EH
    
    Set ai = New OpenAI
    ai.ApiKey = Environ$("OPENAI_API_KEY")
    
    Set Messages = New Collection
    Messages.Add OpenAIMessageDeveloper("You are a concise assistant.")
    Messages.Add OpenAIMessageUser("What is COM?")
    Messages.Add OpenAIMessageAssistant("COM stands for Component Object Model.")
    Messages.Add OpenAIMessageUser("Now explain it for a beginner in one sentence.")
    
    Set resp = ai.CreateChatCompletion( _
        Model:="gpt-5.4", _
        Messages:=Messages, _
        Temperature:=1, _
        MaxCompletionTokens:=100, _
        Verbosity:="low", _
        ReasoningEffort:="low" _
    )
    
    Debug.Print OpenAIExtractText(resp)
    Exit Sub

EH:
    Debug.Print "[ERROR] "; Err.Number; " - "; Err.Description
End Sub
```

## Structured output with JSON Schema

Chat Completions supports `response_format` with `json_schema`, which enforces a schema-based structured output format. ([platform.openai.com](https://platform.openai.com/docs/api-reference/chat/create-chat-completion?utm_source=openai))

```vb
Public Sub TestOpenAI_JsonSchema()
    Dim ai As OpenAI
    Dim resp As JsonData
    Dim Messages As Collection
    Dim schemaJson As String
    Dim responseFormat As String
    
    On Error GoTo EH
    
    Set ai = New OpenAI
    ai.ApiKey = Environ$("OPENAI_API_KEY")
    
    schemaJson = _
        "{" & _
            """type"":""object""," & _
            """properties"":{" & _
                """title"":{""type"":""string""}," & _
                """year"":{""type"":""integer""}" & _
            "}," & _
            """required"":[""title"",""year""]," & _
            """additionalProperties"":false" & _
        "}"
    
    responseFormat = OpenAIResponseFormatJsonSchema( _
        Name:="movie_info", _
        schemaJson:=schemaJson, _
        Description:="Movie information", _
        Strict:=True _
    )
    
    Set Messages = New Collection
    Messages.Add OpenAIMessageDeveloper("Return only data that matches the schema.")
    Messages.Add OpenAIMessageUser("Provide the title and year for The Matrix.")
    
    Set resp = ai.CreateChatCompletion( _
        Model:="gpt-5.4", _
        Messages:=Messages, _
        MaxCompletionTokens:=100, _
        Verbosity:="low", _
        ResponseFormatJson:=responseFormat _
    )
    
    Debug.Print OpenAIExtractText(resp)
    Exit Sub

EH:
    Debug.Print "[ERROR] "; Err.Number; " - "; Err.Description
End Sub
```

## JSON object mode

Chat Completions also supports the older `json_object` response format, though `json_schema` is preferred when supported. ([platform.openai.com](https://platform.openai.com/docs/api-reference/chat/create-chat-completion?utm_source=openai))

```vb
Public Sub TestOpenAI_JsonObject()
    Dim ai As OpenAI
    Dim resp As JsonData
    Dim Messages As Collection
    Dim responseFormat As String
    
    On Error GoTo EH
    
    Set ai = New OpenAI
    ai.ApiKey = Environ$("OPENAI_API_KEY")
    
    Set Messages = New Collection
    Messages.Add OpenAIMessageDeveloper("Return valid JSON only.")
    Messages.Add OpenAIMessageUser("Return an object with keys title and year for The Matrix.")
    
    responseFormat = OpenAIResponseFormatJsonObject()
    
    Set resp = ai.CreateChatCompletion( _
        Model:="gpt-5.4", _
        Messages:=Messages, _
        MaxCompletionTokens:=100, _
        Verbosity:="low", _
        ResponseFormatJson:=responseFormat _
    )
    
    Debug.Print OpenAIExtractText(resp)
    Exit Sub

EH:
    Debug.Print "[ERROR] "; Err.Number; " - "; Err.Description
End Sub
```

## Function tool call example

Chat Completions supports function tools, and `tool_choice` defaults to `auto` when tools are present. `parallel_tool_calls` is a separate optional boolean. ([platform.openai.com](https://platform.openai.com/docs/api-reference/chat/create-chat-completion?utm_source=openai))

```vb
Public Sub TestOpenAI_FunctionToolCall_RequestOnly()
    Dim ai As OpenAI
    Dim resp As JsonData
    Dim Messages As Collection
    Dim ToolsJson As String
    Dim toolCalls As JsonData
    
    On Error GoTo EH
    
    Set ai = New OpenAI
    ai.ApiKey = Environ$("OPENAI_API_KEY")
    
    ToolsJson = _
        "[" & _
            "{" & _
                """type"":""function""," & _
                """function"":{" & _
                    """name"":""get_current_weather""," & _
                    """description"":""Get the current weather in a given location""," & _
                    """parameters"":{" & _
                        """type"":""object""," & _
                        """properties"":{" & _
                            """location"":{""type"":""string""}," & _
                            """unit"":{""type"":""string"",""enum"":[""celsius"",""fahrenheit""]}" & _
                        "}," & _
                        """required"":[""location"",""unit""]," & _
                        """additionalProperties"":false" & _
                    "}," & _
                    """strict"":true" & _
                "}" & _
            "}" & _
        "]"
    
    Set Messages = New Collection
    Messages.Add OpenAIMessageUser("What is the weather in Boston today?")
    
    Set resp = ai.CreateChatCompletion( _
        Model:="gpt-5.4", _
        Messages:=Messages, _
        ToolsJson:=ToolsJson, _
        ToolChoiceJson:=OpenAIToolChoiceAuto(), _
        ParallelToolCalls:=False _
    )
    
    Set toolCalls = OpenAIExtractToolCalls(resp)
    
    Debug.Print "finish_reason = "; OpenAIExtractFinishReason(resp)
    If Not toolCalls Is Nothing Then
        If toolCalls.IsValid Then
            Debug.Print toolCalls.ToJSON("  ")
        Else
            Debug.Print OpenAIExtractText(resp)
        End If
    Else
        Debug.Print OpenAIExtractText(resp)
    End If
    
    Exit Sub

EH:
    Debug.Print "[ERROR] "; Err.Number; " - "; Err.Description
End Sub
```

## Important VB6 client behavior

Your VB6 wrapper should follow these rules:

- Optional parameters default to `Empty` when not provided, so you only need to pass parameters you actually want to set.
- Omit optional JSON fields entirely when they are not being used.
- Do **not** send `Empty` values as raw JSON fields.
- Serialize floating-point numbers with a leading zero, so `0.2` is emitted as `0.2`, not `.2`.

---

## Responses API

The client now includes `CreateResponse` and `CreateResponseSimple` methods for the [Responses API](https://platform.openai.com/docs/api-reference/responses/create), which OpenAI recommends for new text-generation builds. The Responses API uses `input` + `instructions` instead of `messages`, has a flat tools format (no `function` wrapper), and returns `output[]` instead of `choices[]`.

### Simple text generation

```vb
Public Sub TestResponsesSimple()
    Dim ai As OpenAI
    Dim resp As JsonData

    Set ai = New OpenAI
    ai.ApiKey = Environ$("OPENAI_API_KEY")

    Set resp = ai.CreateResponseSimple( _
        "gpt-5.4", _
        "You are a helpful assistant.", _
        "Write a short haiku about VB6 and APIs." _
    )

    Debug.Print ResponsesExtractText(resp)
    Debug.Print ResponsesExtractStatus(resp)   ' "completed"
End Sub
```

### Multi-turn conversation via `previous_response_id`

The Responses API simplifies multi-turn by letting you chain responses. Pass the previous response's `id` as `PreviousResponseId` — the API carries forward context automatically.

```vb
Public Sub TestResponsesMultiTurn()
    Dim ai As OpenAI
    Dim resp1 As JsonData, resp2 As JsonData
    Dim firstId As String

    Set ai = New OpenAI
    ai.ApiKey = Environ$("OPENAI_API_KEY")

    ' Turn 1
    Set resp1 = ai.CreateResponse( _
        Model:="gpt-5.4", _
        InputItems:="Tell me a short joke about programming.", _
        Instructions:="You are a witty assistant." _
    )
    Debug.Print ResponsesExtractText(resp1)

    ' Turn 2 — chain via previous_response_id
    firstId = CStr(resp1.GetChildByPath("id").ScalarValue)
    Set resp2 = ai.CreateResponse( _
        Model:="gpt-5.4", _
        InputItems:="Now explain why that joke is funny.", _
        Instructions:="You are a witty assistant.", _
        PreviousResponseId:=firstId _
    )
    Debug.Print ResponsesExtractText(resp2)
End Sub
```

### Structured output with JSON Schema

The Responses API uses `text.format` instead of `response_format`. Use `ResponsesTextFormatJsonSchema` to build the format parameter.

```vb
Public Sub TestResponsesJsonSchema()
    Dim ai As OpenAI
    Dim resp As JsonData
    Dim schemaJson As String
    Dim textFormat As String

    Set ai = New OpenAI
    ai.ApiKey = Environ$("OPENAI_API_KEY")

    schemaJson = _
        "{" & _
            """type"":""object""," & _
            """properties"":{" & _
                """title"":{""type"":""string""}," & _
                """year"":{""type"":""integer""}" & _
            "}," & _
            """required"":[""title"",""year""]," & _
            """additionalProperties"":false" & _
        "}"

    textFormat = ResponsesTextFormatJsonSchema( _
        Name:="movie_info", _
        schemaJson:=schemaJson, _
        Strict:=True _
    )

    Set resp = ai.CreateResponse( _
        Model:="gpt-5.4", _
        InputItems:="Provide the title and year for The Matrix.", _
        Instructions:="Return only data that matches the schema.", _
        TextFormatJson:=textFormat _
    )

    Debug.Print ResponsesExtractText(resp)
End Sub
```

### Agentic tool-calling loop

The Responses API supports function tools (flat format — no `function` wrapper). When building agentic loops, the critical pattern is:

1. On the first call, send the user's message.
2. On subsequent calls, pass `PreviousResponseId` and send **only** new `function_call_output` items as `InputItems`. The API carries forward prior context automatically.

```vb
Public Sub TestResponsesToolCalling()
    Dim ai As OpenAI
    Dim resp As JsonData
    Dim InputItems As Collection
    Dim toolsJson As String
    Dim outputItems As JsonData
    Dim status As String
    Dim i As Long
    Dim item As JsonData
    Dim itemType As String
    Dim previousId As String
    Dim callId As String
    Dim toolResult As String
    Dim loopCounter As Long, maxLoops As Long

    Set ai = New OpenAI
    ai.ApiKey = Environ$("OPENAI_API_KEY")

    ' Tools in flat Responses format — no "function" wrapper
    toolsJson = "[{""type"":""function""," & _
        """name"":""get_current_weather""," & _
        """description"":""Get the current weather in a given location""," & _
        """parameters"":{""type"":""object""," & _
        """properties"":{""location"":{""type"":""string""}}," & _
        """required"":[""location""],""additionalProperties"":false}," & _
        """strict"":true}]"

    previousId = ""
    Set InputItems = New Collection
    InputItems.Add OpenAIMessageUser("What's the weather in Boston?")

    loopCounter = 0
    maxLoops = 10

    Do While loopCounter < maxLoops
        loopCounter = loopCounter + 1

        If LenB(previousId) = 0 Then
            ' First call: send the initial user message
            Set resp = ai.CreateResponse( _
                Model:="gpt-4.1", _
                InputItems:=InputItems, _
                Instructions:="You are a helpful weather assistant. Use tools to answer questions.", _
                ToolsJson:=toolsJson, _
                ToolChoiceJson:=ResponsesToolChoiceAuto(), _
                ParallelToolCalls:=True, _
                MaxOutputTokens:=1024 _
            )
        Else
            ' Subsequent calls: chain via previous_response_id,
            ' send only new function_call_output items
            Set resp = ai.CreateResponse( _
                Model:="gpt-4.1", _
                InputItems:=InputItems, _
                Instructions:="You are a helpful weather assistant. Use tools to answer questions.", _
                ToolsJson:=toolsJson, _
                ToolChoiceJson:=ResponsesToolChoiceAuto(), _
                ParallelToolCalls:=True, _
                MaxOutputTokens:=1024, _
                PreviousResponseId:=previousId _
            )
        End If

        status = ResponsesExtractStatus(resp)

        ' Capture the response id for chaining the next iteration
        previousId = CStr(resp.GetChildByPath("id").ScalarValue)

        ' Reset InputItems — only queue new function_call_output items
        Set InputItems = New Collection

        ' Check output items for function calls
        Set outputItems = ResponsesExtractOutputItems(resp)
        Dim hasToolCalls As Boolean
        hasToolCalls = False

        If Not outputItems Is Nothing Then
            If outputItems.IsArray Then
                For i = 0 To outputItems.ArrayLength - 1
                    Set item = outputItems.GetArrayItem(i)
                    itemType = CStr(item.GetChildByPath("type").ScalarValue)

                    If StrComp(itemType, "function_call", vbTextCompare) = 0 Then
                        hasToolCalls = True
                        callId = CStr(item.GetChildByPath("call_id").ScalarValue)
                        Debug.Print "  Tool: " & CStr(item.GetChildByPath("name").ScalarValue)

                        ' Execute the tool and queue the result
                        toolResult = ExecuteToolFunction( _
                            CStr(item.GetChildByPath("name").ScalarValue), _
                            CStr(item.GetChildByPath("arguments").ScalarValue))
                        InputItems.Add ResponsesBuildFunctionCallResult(callId, toolResult)
                    End If
                Next i
            End If
        End If

        If Not hasToolCalls Then Exit Do
    Loop

    Debug.Print "Final: " & ResponsesExtractText(resp)
End Sub

' Helper: build a function_call_output item for the Responses API
Private Function ResponsesBuildFunctionCallResult(ByVal callId As String, ByVal output As String) As String
    ResponsesBuildFunctionCallResult = _
        "{" & _
            """type"":""function_call_output""," & _
            """call_id"":" & JsonString(callId) & "," & _
            """output"":" & JsonString(output) & _
        "}"
End Function
```

**Key points for agentic loops:**

- **Always chain with `previous_response_id`** after the first call. The `function_call` items from the prior response must be in context for `function_call_output` items to match.
- **Reset `InputItems` each iteration** to contain only the new `function_call_output` items. Prior messages are carried forward automatically.
- Tools use a **flat format**: `{"type":"function","name":"...","description":"...","parameters":{...},"strict":true}` — no `"function"` wrapper.
- Tool results use `{"type":"function_call_output","call_id":"...","output":"..."}` — the `call_id` must match the `function_call` item's `call_id`.

### Responses API helpers (`OpenAIHelpers.bas`)

| Helper | Returns | Description |
|--------|---------|-------------|
| `ResponsesExtractText(resp)` | `String` | Reads `output_text` field; falls back to `output[0].content[0].text` |
| `ResponsesExtractStatus(resp)` | `String` | Reads `status` field (`completed`, `failed`, `in_progress`, …) |
| `ResponsesExtractOutputItems(resp)` | `JsonData` | Returns the `output[]` array node |
| `ResponsesExtractToolCalls(resp)` | `JsonData` | Returns the `output[]` array (iterate for `function_call` items) |
| `ResponsesTextFormatText()` | `String` | `{"format":{"type":"text"}}` |
| `ResponsesTextFormatJsonObject()` | `String` | `{"format":{"type":"json_object"}}` |
| `ResponsesTextFormatJsonSchema(name, schema, ...)` | `String` | Builds `text.format` with `json_schema` |
| `ResponsesReasoning(effort, [summary])` | `String` | Builds `{"effort":...,"summary":...}` |
| `ResponsesToolChoiceAuto()` / `None` / `Required` | `String` | Tool choice constants |

### `CreateResponse` parameters

`CreateResponse(Model, InputItems, [optional...])`

| Parameter | Type | Description |
|-----------|------|-------------|
| `Model` | `String` | **Required.** Model ID, e.g. `"gpt-4.1"` |
| `InputItems` | `String` or `Collection` | **Required.** A user message string or a `Collection` of message/function_call_output JSON objects |
| `Instructions` | `String` | System/developer prompt (replaces the old `messages[0].role:"developer"` pattern) |
| `Temperature` | `Double` | Sampling temperature (0–2) |
| `MaxOutputTokens` | `Long` | Upper bound on output tokens |
| `TopP` | `Double` | Nucleus sampling (0–1) |
| `ToolsJson` | `String` | JSON array of tool definitions (flat format) |
| `ToolChoiceJson` | `String` | Tool choice (use `ResponsesToolChoiceAuto` etc.) |
| `TextFormatJson` | `String` | Text format config (use `ResponsesTextFormat*` helpers) |
| `ReasoningJson` | `String` | Reasoning config (use `ResponsesReasoning`) |
| `PreviousResponseId` | `String` | ID of the previous response for multi-turn chaining |
| `ParallelToolCalls` | `Boolean` | Allow parallel tool calls |
| `Store` | `Boolean` | Store the response for later retrieval |
| `MaxToolCalls` | `Long` | Maximum number of tool calls |

### Chat Completions vs. Responses — quick reference

| Aspect | Chat Completions | Responses |
|--------|-----------------|-----------|
| Endpoint | `/v1/chat/completions` | `/v1/responses` |
| Client method | `CreateChatCompletion` / `CreateChatCompletionSimple` | `CreateResponse` / `CreateResponseSimple` |
| User input | `messages` array | `input` (string or Collection) |
| System prompt | `role:"developer"` message in array | `instructions` parameter |
| Text helper | `OpenAIExtractText` | `ResponsesExtractText` |
| Status helper | `OpenAIExtractFinishReason` | `ResponsesExtractStatus` |
| Tool helper | `OpenAIExtractToolCalls` → `choices[0].message.tool_calls` | `ResponsesExtractOutputItems` → `output[]` |
| Structured output | `response_format` parameter | `text.format` parameter |
| Reasoning | `reasoning_effort` string | `reasoning` object (`{effort, summary}`) |
| Multi-turn | Rebuild `messages` array manually | `previous_response_id` chains automatically |
| Tools format | `{type:"function", function:{name, ...}}` | `{type:"function", name:"...", ...}` (flat) |
| Tool result format | `{role:"tool", tool_call_id, content}` | `{type:"function_call_output", call_id, output}` |

---
Learn more:
1. [Text generation - OpenAI API](https://platform.openai.com/docs/guides/chat-completions?utm_source=openai)
2. [Chat Completions | OpenAI API Reference](https://platform.openai.com/docs/api-reference/chat/create-chat-completion?utm_source=openai)
3. [Responses API | OpenAI API Reference](https://platform.openai.com/docs/api-reference/responses/create?utm_source=openai)