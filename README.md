```md
# VB6 OpenAI + JsonData

A small VB6 client for working with JSON and the OpenAI Chat Completions API.

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
        Temperature:=Empty, _
        MaxCompletionTokens:=100, _
        Verbosity:="low", _
        ResponseFormatJson:=Empty, _
        ToolsJson:=Empty, _
        ToolChoiceJson:=Empty, _
        ReasoningEffort:="low", _
        Store:=Empty, _
        MetadataJson:=Empty, _
        ParallelToolCalls:=Empty _
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
        Temperature:=Empty, _
        MaxCompletionTokens:=100, _
        Verbosity:="low", _
        ResponseFormatJson:=responseFormat, _
        ToolsJson:=Empty, _
        ToolChoiceJson:=Empty, _
        ReasoningEffort:="low", _
        Store:=Empty, _
        MetadataJson:=Empty, _
        ParallelToolCalls:=Empty _
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
        Temperature:=Empty, _
        MaxCompletionTokens:=100, _
        Verbosity:="low", _
        ResponseFormatJson:=responseFormat, _
        ToolsJson:=Empty, _
        ToolChoiceJson:=Empty, _
        ReasoningEffort:="low", _
        Store:=Empty, _
        MetadataJson:=Empty, _
        ParallelToolCalls:=Empty _
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
        Temperature:=Empty, _
        MaxCompletionTokens:=Empty, _
        Verbosity:=Empty, _
        ResponseFormatJson:=Empty, _
        ToolsJson:=ToolsJson, _
        ToolChoiceJson:=OpenAIToolChoiceAuto(), _
        ReasoningEffort:=Empty, _
        Store:=Empty, _
        MetadataJson:=Empty, _
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

- Omit optional JSON fields entirely when they are not being used.
- Do **not** send `Empty` values as raw JSON fields.
- Serialize floating-point numbers with a leading zero, so `0.2` is emitted as `0.2`, not `.2`.
- Only send `parallel_tool_calls` when you are actually sending tools, since it only applies during tool use. ([platform.openai.com](https://platform.openai.com/docs/api-reference/chat/create-chat-completion?utm_source=openai))
- `metadata` must be a JSON object if sent. ([platform.openai.com](https://platform.openai.com/docs/api-reference/chat/create-chat-completion?utm_source=openai))
- `response_format` must be a JSON object if sent. ([platform.openai.com](https://platform.openai.com/docs/api-reference/chat/create-chat-completion?utm_source=openai))

## Recommendation

If you continue evolving this client, consider adding a separate `CreateResponse` wrapper for the newer Responses API, since OpenAI recommends that API for new applications. ([platform.openai.com](https://platform.openai.com/docs/guides/chat-completions?utm_source=openai))
```

---
Learn more:
1. [Text generation - OpenAI API](https://platform.openai.com/docs/guides/chat-completions?utm_source=openai)
2. [Chat Completions | OpenAI API Reference](https://platform.openai.com/docs/api-reference/chat/create-chat-completion?utm_source=openai)