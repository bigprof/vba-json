# VB6 OpenAI + JsonData

A small VB6 client for working with JSON and the OpenAI Chat Completions API.

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

## Notes

- These examples use the **Chat Completions API**.
- OpenAI recommends the **Responses API** for new text-generation builds, but Chat Completions remains available and is what this VB6 client currently targets. ([platform.openai.com](https://platform.openai.com/docs/guides/chat-completions?utm_source=openai))
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
- `ReasoningEffort` defaults to `"low"` and `Store` defaults to `False`; override them explicitly if needed.
- Omit optional JSON fields entirely when they are not being used.
- Do **not** send `Empty` values as raw JSON fields.
- Serialize floating-point numbers with a leading zero, so `0.2` is emitted as `0.2`, not `.2`.
- Only send `parallel_tool_calls` when you are actually sending tools, since it only applies during tool use. ([platform.openai.com](https://platform.openai.com/docs/api-reference/chat/create-chat-completion?utm_source=openai))
- `metadata` must be a JSON object if sent. ([platform.openai.com](https://platform.openai.com/docs/api-reference/chat/create-chat-completion?utm_source=openai))
- `response_format` must be a JSON object if sent. ([platform.openai.com](https://platform.openai.com/docs/api-reference/chat/create-chat-completion?utm_source=openai))

## Recommendation

If you continue evolving this client, consider adding a separate `CreateResponse` wrapper for the newer Responses API, since OpenAI recommends that API for new applications. ([platform.openai.com](https://platform.openai.com/docs/guides/chat-completions?utm_source=openai))

---
Learn more:
1. [Text generation - OpenAI API](https://platform.openai.com/docs/guides/chat-completions?utm_source=openai)
2. [Chat Completions | OpenAI API Reference](https://platform.openai.com/docs/api-reference/chat/create-chat-completion?utm_source=openai)