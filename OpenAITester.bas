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
    
    Debug.Print OpenAIExtractText(result)
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
        ReasoningEffort:="low" _
    )
    
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
        MaxCompletionTokens:=100, _
        Verbosity:="low", _
        ResponseFormatJson:=responseFormat _
    )
    
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
        MaxCompletionTokens:=100, _
        Verbosity:="low", _
        ResponseFormatJson:=responseFormat _
    )
    
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
        Model:="gpt-5.2", _
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

Public Sub TestOpenAI_MultiTurn_UsesPriorAnswer()
    Dim ai As OpenAI
    Dim resp1 As JsonData
    Dim resp2 As JsonData
    Dim Messages As Collection
    Dim assistantText1 As String
    Dim followUp As String

    On Error GoTo EH

    Set ai = New OpenAI
    ai.ApiKey = Environ$("OPENAI_API_KEY")

    ' Turn 1: ask for something we can build on
    Set Messages = New Collection
    Messages.Add OpenAIMessageDeveloper("You are a concise assistant. Use plain language.")
    Messages.Add OpenAIMessageUser( _
        "Give me 3 bullet points explaining COM in VB6, plus a short analogy." _
    )

    Set resp1 = ai.CreateChatCompletion( _
        Model:="gpt-5.4", _
        Messages:=Messages, _
        Temperature:=1, _
        MaxCompletionTokens:=200, _
        Verbosity:="low", _
        ReasoningEffort:="low" _
    )

    assistantText1 = OpenAIExtractText(resp1)

    Debug.Print String(70, "=")
    Debug.Print "TURN 1 (assistant)"
    Debug.Print String(70, "=")
    Debug.Print assistantText1
    Debug.Print

    ' Turn 2: use the prior assistant output as the basis for a new request
    ' Important: append the assistant message to the same Messages collection
    Messages.Add OpenAIMessageAssistant(assistantText1)

    followUp = _
        "Using your explanation above:" & vbCrLf & _
        "1) write a tiny VB6 pseudo-example of creating an object and calling a method," & vbCrLf & _
        "2) then rewrite your analogy in one sentence for a 10-year-old."

    Messages.Add OpenAIMessageUser(followUp)

    Set resp2 = ai.CreateChatCompletion( _
        Model:="gpt-5.4", _
        Messages:=Messages, _
        Temperature:=1, _
        MaxCompletionTokens:=250, _
        Verbosity:="low", _
        ReasoningEffort:="low" _
    )

    Debug.Print String(70, "=")
    Debug.Print "TURN 2 (assistant)"
    Debug.Print String(70, "=")
    Debug.Print OpenAIExtractText(resp2)
    Debug.Print

    Exit Sub

EH:
    Debug.Print "[ERROR] "; Err.Number; " - "; Err.Description
End Sub

