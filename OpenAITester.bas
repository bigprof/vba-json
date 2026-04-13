Attribute VB_Name = "OpenAITester"
Option Explicit

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
        UserPrompt:="Write a short haiku about VB6 and APIs.", _
        Temperature:=0.7, _
        MaxCompletionTokens:=120, _
        Verbosity:="low" _
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
    Dim messages As Collection
    
    On Error GoTo EH
    
    Set ai = New OpenAI
    ai.ApiKey = Environ$("OPENAI_API_KEY")
    
    Set messages = New Collection
    messages.Add OpenAIMessageDeveloper("You are a concise assistant.")
    messages.Add OpenAIMessageUser("What is COM?")
    messages.Add OpenAIMessageAssistant("COM stands for Component Object Model.")
    messages.Add OpenAIMessageUser("Now explain it for a beginner in one sentence.")
    
    Set resp = ai.CreateChatCompletion( _
        Model:="gpt-5.4", _
        messages:=messages, _
        Temperature:=0.2, _
        MaxCompletionTokens:=100, _
        Verbosity:="low" _
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
    Dim messages As Collection
    Dim responseFormat As String
    
    On Error GoTo EH
    
    Set ai = New OpenAI
    ai.ApiKey = Environ$("OPENAI_API_KEY")
    
    Set messages = New Collection
    messages.Add OpenAIMessageDeveloper("Return valid JSON only.")
    messages.Add OpenAIMessageUser("Return an object with keys title and year for The Matrix.")
    
    responseFormat = OpenAIResponseFormatJsonObject()
    
    Set resp = ai.CreateChatCompletion( _
        Model:="gpt-5.4", _
        messages:=messages, _
        MaxCompletionTokens:=100, _
        Verbosity:="low", _
        ResponseFormatJson:=responseFormat _
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
    Dim messages As Collection
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
    
    Set messages = New Collection
    messages.Add OpenAIMessageDeveloper("Return only data that matches the schema.")
    messages.Add OpenAIMessageUser("Provide the title and year for The Matrix.")
    
    Set resp = ai.CreateChatCompletion( _
        Model:="gpt-5.4", _
        messages:=messages, _
        MaxCompletionTokens:=100, _
        Verbosity:="low", _
        ResponseFormatJson:=responseFormat _
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
    Dim messages As Collection
    Dim toolsJson As String
    Dim toolCalls As JsonData
    
    On Error GoTo EH
    
    Set ai = New OpenAI
    ai.ApiKey = Environ$("OPENAI_API_KEY")
    
    toolsJson = _
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
                        """required"":[""location""]," & _
                        """additionalProperties"":false" & _
                    "}," & _
                    """strict"":true" & _
                "}" & _
            "}" & _
        "]"
    
    Set messages = New Collection
    messages.Add OpenAIMessageUser("What is the weather in Boston today?")
    
    Set resp = ai.CreateChatCompletion( _
        Model:="gpt-5.4", _
        messages:=messages, _
        toolsJson:=toolsJson, _
        ToolChoiceJson:=OpenAIToolChoiceAuto(), _
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

