Attribute VB_Name = "OpenAITester"
Option Explicit

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
    
    'Debug.Print result.json
End Sub

Public Sub TestOpenAIAdvanced()
    Dim client As OpenAI
    Dim msgs As Collection
    Dim result As JsonData
    
    Set client = New OpenAI
    client.ApiKey = Environ$("OPENAI_API_KEY")
    
    Set msgs = New Collection
    msgs.Add "{""role"":""developer"",""content"":""You are a concise assistant.""}"
    msgs.Add "{""role"":""user"",""content"":""What is COM?""}"
    msgs.Add "{""role"":""assistant"",""content"":""COM stands for Component Object Model.""}"
    msgs.Add "{""role"":""user"",""content"":""Now explain it for a beginner in one sentence.""}"
    
    Set result = client.CreateChatCompletion( _
        "gpt-5.4", _
        msgs, _
        Empty, _
        "100", _
        "low", _
        Empty, _
        "[{""type"":""function"",""function"":{""name"":""dummy_tool"",""description"":""Dummy tool for testing."",""parameters"":{""type"":""object"",""properties"":{}}}}]", _
        Empty, _
        Empty, _
        True, _
        Empty, _
        False _
    )
End Sub

Public Sub TestOpenAI_SimpleText()
    Dim ai As OpenAI
    Dim resp As JsonData
    Dim text As String
    
    On Error GoTo EH
    
    Set ai = New OpenAI
    ai.ApiKey = Environ$("OPENAI_API_KEY")
    
    Set resp = ai.CreateChatCompletionSimple( _
        Model:="gpt-5.4", _
        DeveloperPrompt:="You are a helpful assistant.", _
        UserPrompt:="Write a short haiku about VB6 and APIs." _
    )
    
    Debug.Print String(70, "=")
    Debug.Print "Simple text response"
    Debug.Print String(70, "=")
    Debug.Print "finish_reason = "; OpenAIExtractFinishReason(resp)
    Debug.Print "content:"
    Debug.Print OpenAIExtractText(resp)
    Debug.Print
    Debug.Print "raw response JSON:"
    Debug.Print resp.ToJSON("  ")
    
    Exit Sub

EH:
    Debug.Print "[ERROR] "; Err.Number; " - "; Err.Description
End Sub

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
        ResponseFormatJson:=Empty, _
        ToolsJson:=Empty, _
        ToolChoiceJson:=Empty, _
        MetadataJson:=Empty, _
        ReasoningEffort:="low", _
        ParallelToolCalls:=Empty _
    )
    
    Debug.Print String(70, "=")
    Debug.Print "Multi-turn response"
    Debug.Print String(70, "=")
    Debug.Print OpenAIExtractText(resp)
    
    Exit Sub

EH:
    Debug.Print "[ERROR] "; Err.Number; " - "; Err.Description
End Sub

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
    
    Debug.Print String(70, "=")
    Debug.Print "JSON object response"
    Debug.Print String(70, "=")
    Debug.Print OpenAIExtractText(resp)
    
    Exit Sub

EH:
    Debug.Print "[ERROR] "; Err.Number; " - "; Err.Description
End Sub

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
    
    Debug.Print String(70, "=")
    Debug.Print "JSON schema response"
    Debug.Print String(70, "=")
    Debug.Print OpenAIExtractText(resp)
    
    Exit Sub

EH:
    Debug.Print "[ERROR] "; Err.Number; " - "; Err.Description
End Sub
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
    
    Debug.Print String(70, "=")
    Debug.Print "Tool call request-only response"
    Debug.Print String(70, "=")
    Debug.Print "finish_reason = "; OpenAIExtractFinishReason(resp)
    
    If Not toolCalls Is Nothing Then
        If toolCalls.IsValid Then
            Debug.Print "tool_calls detected"
            Debug.Print toolCalls.ToJSON("  ")
        Else
            Debug.Print "no tool_calls"
            Debug.Print OpenAIExtractText(resp)
        End If
    Else
        Debug.Print "no tool_calls"
        Debug.Print OpenAIExtractText(resp)
    End If
    
    Exit Sub

EH:
    Debug.Print "[ERROR] "; Err.Number; " - "; Err.Description
End Sub

