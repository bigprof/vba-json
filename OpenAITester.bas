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

' ============================================================================
' NEW TEST: Full tool calling request-response loop
' ============================================================================
' This test demonstrates a complete tool call loop where:
' 1. The client requests tool calls
' 2. Tool results are collected and added to the message history
' 3. The loop continues until finish_reason is no longer "tool_calls"
' ============================================================================

Public Sub TestOpenAI_FunctionToolCall_RequestResponseLoop()
    Dim ai As OpenAI
    Dim resp As JsonData
    Dim Messages As Collection
    Dim ToolsJson As String
    Dim toolCalls As JsonData
    Dim finishReason As String
    Dim loopCounter As Long
    Dim maxLoops As Long
    
    On Error GoTo EH
    
    Set ai = New OpenAI
    ai.ApiKey = Environ$("OPENAI_API_KEY")
    
    ' Build the tools JSON using a helper function
    ToolsJson = BuildWeatherToolsJson()
    
    ' Initialize the message collection with a user question
    Set Messages = New Collection
    Messages.Add OpenAIMessageDeveloper( _
        "You are a helpful weather assistant. Use the provided tools to answer questions." _
    )
    Messages.Add OpenAIMessageUser( _
        "What's the weather in Boston and what are its coordinates?" _
    )
    
    Debug.Print String(80, "=")
    Debug.Print "TOOL CALL REQUEST-RESPONSE LOOP TEST"
    Debug.Print String(80, "=")
    Debug.Print
    
    loopCounter = 0
    maxLoops = 10 ' Prevent infinite loops
    
    ' Main tool calling loop
    Do While loopCounter < maxLoops
        loopCounter = loopCounter + 1
        Debug.Print "--- Iteration " & loopCounter & " ---"
        
        ' Send request to API
        Set resp = ai.CreateChatCompletion( _
            Model:="gpt-4o-mini", _
            Messages:=Messages, _
            MaxCompletionTokens:=1024, _
            ToolsJson:=ToolsJson, _
            ToolChoiceJson:=OpenAIToolChoiceAuto(), _
            ParallelToolCalls:=True _
        )
        
        ' Extract finish reason
        finishReason = OpenAIExtractFinishReason(resp)
        Debug.Print "Finish reason: " & finishReason
        
        ' Check if we have tool calls to process
        If StrComp(finishReason, "tool_calls", vbTextCompare) = 0 Then
            ' Extract tool calls from the response
            Set toolCalls = OpenAIExtractToolCalls(resp)
            
            If Not toolCalls Is Nothing Then
                If toolCalls.IsValid Then
                    ' Add the assistant's response to message history, including tool_calls
                    ' so that subsequent 'tool' role messages are valid per the API
                    Messages.Add OpenAIMessageAssistantWithToolCalls(OpenAIExtractText(resp), toolCalls.ToJSON())
                    
                    ' Process each tool call and add results
                    ProcessToolCalls toolCalls, Messages
                    
                    Debug.Print "Tool calls processed and results added to messages."
                Else
                    Debug.Print "Tool calls could not be parsed."
                    Exit Do
                End If
            Else
                Debug.Print "No tool calls found in response."
                Exit Do
            End If
        Else
            ' No more tool calls - exit loop
            Debug.Print "Finish reason is not 'tool_calls', exiting loop."
            Debug.Print vbCrLf & "Final assistant response:"
            Debug.Print OpenAIExtractText(resp)
            Exit Do
        End If
        
        Debug.Print
    Loop
    
    If loopCounter >= maxLoops Then
        Debug.Print "Maximum iterations reached (" & maxLoops & ")"
    End If
    
    Debug.Print String(80, "=")
    Debug.Print "LOOP COMPLETED"
    Debug.Print String(80, "=")
    
    Exit Sub

EH:
    Debug.Print "[ERROR] "; Err.Number; " - "; Err.Description
End Sub

' ============================================================================
' Helper: Build weather tools JSON
' ============================================================================

Private Function BuildWeatherToolsJson() As String
    Dim s1 As String
    Dim s2 As String
    Dim result As String
    
    ' Tool 1: get_current_weather
    s1 = "{""type"":""function"","
    s1 = s1 & """function"":{""name"":""get_current_weather"","
    s1 = s1 & """description"":""Get the current weather in a given location"","
    s1 = s1 & """parameters"":{""type"":""object"","
    s1 = s1 & """properties"":{""location"":{""type"":""string"","
    s1 = s1 & """description"":""The city and state, e.g. San Francisco, CA""},"
    s1 = s1 & """unit"":{""type"":""string"","
    s1 = s1 & """enum"":[""celsius"",""fahrenheit""],"
    s1 = s1 & """description"":""Temperature unit""}},"
    s1 = s1 & """required"":[""location"",""unit""],"
    s1 = s1 & """additionalProperties"":false},"
    s1 = s1 & """strict"":true}}"
    
    ' Tool 2: get_location_coordinates
    s2 = "{""type"":""function"","
    s2 = s2 & """function"":{""name"":""get_location_coordinates"","
    s2 = s2 & """description"":""Get geographic coordinates for a location"","
    s2 = s2 & """parameters"":{""type"":""object"","
    s2 = s2 & """properties"":{""location"":{""type"":""string"","
    s2 = s2 & """description"":""The city and state, e.g. San Francisco, CA""}},"
    s2 = s2 & """required"":[""location""],"
    s2 = s2 & """additionalProperties"":false},"
    s2 = s2 & """strict"":true}}"
    
    result = "[" & s1 & "," & s2 & "]"
    Debug.Print ""
    Debug.Print "ToolsJson:"
    Debug.Print result
    Debug.Print
    BuildWeatherToolsJson = result
End Function

' ============================================================================
' Helper: Process tool calls and add results to message history
' ============================================================================

Private Sub ProcessToolCalls(ByVal toolCalls As JsonData, ByRef Messages As Collection)
    Dim i As Long, ii As String
    Dim toolCount As Long
    Dim toolName As String
    Dim toolInput As String
    Dim toolResult As String
    Dim toolId As String
    
    ' Get the count of tool calls
    ' Assuming toolCalls is a JSON array
    toolCount = toolCalls.ArrayLength
    
    Debug.Print "Processing " & toolCount & " tool call(s)..."
    
    For i = 1 To toolCount
        ' Extract tool information (adjust based on actual JSON structure)
        ' This assumes: toolCalls[i].id, toolCalls[i].function.name, toolCalls[i].function.arguments
        
        ii = CStr(i - 1) ' JSON arrays are typically 0-indexed

        toolId = toolCalls.GetChildByPath(ii & ".id").ScalarValue
        toolName = toolCalls.GetChildByPath(ii & ".function.name").ScalarValue
        toolInput = toolCalls.GetChildByPath(ii & ".function.arguments").ScalarValue
        
        Debug.Print "  Tool #" & i & ": " & toolName
        Debug.Print "    Input: " & toolInput
        
        ' Execute the tool and get the result
        toolResult = ExecuteToolFunction(toolName, toolInput)
        Debug.Print "    Result: " & toolResult
        
        ' Add the tool result to the message history
        ' Format: {"role": "tool", "tool_call_id": "...", "content": "..."}
        Messages.Add OpenAIMessageTool(toolId, toolResult)
    Next i
End Sub

' ============================================================================
' Helper: Simulate tool execution
' ============================================================================

Private Function ExecuteToolFunction(ByVal functionName As String, ByVal arguments As String) As String
    ' This is a mock implementation. In a real scenario, you would:
    ' 1. Parse the JSON arguments
    ' 2. Call the actual function or external service
    ' 3. Return the result as a string
    
    Select Case LCase$(functionName)
        Case "get_current_weather"
            ' Mock weather data
            ExecuteToolFunction = "{""location"":""Boston, MA"",""temperature"":72,""unit"":""fahrenheit"",""description"":""Partly cloudy""}"
        
        Case "get_location_coordinates"
            ' Mock coordinates data
            ExecuteToolFunction = "{""location"":""Boston, MA"",""latitude"":42.3601,""longitude"":-71.0589}"
        
        Case Else
            ExecuteToolFunction = "{""error"":""Unknown function: " & functionName & """}"
    End Select
End Function

' ============================================================================
' Responses API tests
' ============================================================================

Public Sub TestResponsesSimple()
    Dim ai As OpenAI
    Dim resp As JsonData

    On Error GoTo EH

    Set ai = New OpenAI
    ai.ApiKey = Environ$("OPENAI_API_KEY")

    Set resp = ai.CreateResponseSimple( _
        "gpt-5.4", _
        "You are a helpful assistant.", _
        "Write a short haiku about VB6 and APIs." _
    )

    Debug.Print String(70, "=")
    Debug.Print "RESPONSES SIMPLE TEST"
    Debug.Print String(70, "=")
    Debug.Print "Status: "; ResponsesExtractStatus(resp)
    Debug.Print "Output: "; ResponsesExtractText(resp)
    Debug.Print

    Exit Sub

EH:
    Debug.Print "[ERROR] "; Err.Number; " - "; Err.Description
End Sub

Public Sub TestResponsesMultiTurn()
    Dim ai As OpenAI
    Dim resp1 As JsonData
    Dim resp2 As JsonData
    Dim firstId As String

    On Error GoTo EH

    Set ai = New OpenAI
    ai.ApiKey = Environ$("OPENAI_API_KEY")

    ' Turn 1: ask a question
    Set resp1 = ai.CreateResponse( _
        Model:="gpt-5.4", _
        InputItems:="Tell me a short joke about programming.", _
        Instructions:="You are a witty assistant. Keep responses under 2 sentences." _
    )

    Debug.Print String(70, "=")
    Debug.Print "RESPONSES MULTI-TURN TEST"
    Debug.Print String(70, "=")
    Debug.Print "Turn 1 status: "; ResponsesExtractStatus(resp1)
    Debug.Print "Turn 1 output: "; ResponsesExtractText(resp1)
    Debug.Print

    ' Extract the response id for chaining
    firstId = CStr(resp1.GetChildByPath("id").ScalarValue)
    Debug.Print "Response ID: "; firstId
    Debug.Print

    ' Turn 2: follow-up using previous_response_id
    Set resp2 = ai.CreateResponse( _
        Model:="gpt-5.4", _
        InputItems:="Now explain why that joke is funny.", _
        Instructions:="You are a witty assistant. Keep responses under 2 sentences.", _
        PreviousResponseId:=firstId _
    )

    Debug.Print "Turn 2 status: "; ResponsesExtractStatus(resp2)
    Debug.Print "Turn 2 output: "; ResponsesExtractText(resp2)
    Debug.Print

    Exit Sub

EH:
    Debug.Print "[ERROR] "; Err.Number; " - "; Err.Description
End Sub

Public Sub TestResponsesJsonSchema()
    Dim ai As OpenAI
    Dim resp As JsonData
    Dim schemaJson As String
    Dim textFormat As String

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

    textFormat = ResponsesTextFormatJsonSchema( _
        Name:="movie_info", _
        schemaJson:=schemaJson, _
        Description:="Movie information", _
        Strict:=True _
    )

    Set resp = ai.CreateResponse( _
        Model:="gpt-5.4", _
        InputItems:="Provide the title and year for The Matrix.", _
        Instructions:="Return only data that matches the schema.", _
        TextFormatJson:=textFormat _
    )

    Debug.Print String(70, "=")
    Debug.Print "RESPONSES JSON SCHEMA TEST"
    Debug.Print String(70, "=")
    Debug.Print "Status: "; ResponsesExtractStatus(resp)
    Debug.Print "Output: "; ResponsesExtractText(resp)
    Debug.Print

    Exit Sub

EH:
    Debug.Print "[ERROR] "; Err.Number; " - "; Err.Description
End Sub

Public Sub TestResponsesToolCalling()
    Dim ai As OpenAI
    Dim resp As JsonData
    Dim InputItems As Collection
    Dim toolsJson As String
    Dim outputItems As JsonData
    Dim status As String
    Dim loopCounter As Long
    Dim maxLoops As Long
    Dim i As Long
    Dim item As JsonData
    Dim itemType As String
    Dim functionName As String
    Dim functionArgs As String
    Dim callId As String
    Dim toolResult As String
    Dim assistantOutput As String

    On Error GoTo EH

    Set ai = New OpenAI
    ai.ApiKey = Environ$("OPENAI_API_KEY")

    ' Build tools in flat Responses API format (no "function" wrapper)
    toolsJson = BuildResponsesWeatherToolsJson()

    Debug.Print String(80, "=")
    Debug.Print "RESPONSES TOOL CALLING TEST"
    Debug.Print String(80, "=")
    Debug.Print

    loopCounter = 0
    maxLoops = 10

    ' Start with a simple string input
    Set InputItems = New Collection
    InputItems.Add OpenAIMessageUser("What's the weather in Boston and what are its coordinates?")

    Do While loopCounter < maxLoops
        loopCounter = loopCounter + 1
        Debug.Print "--- Iteration " & loopCounter & " ---"

        Set resp = ai.CreateResponse( _
            Model:="gpt-4.1", _
            InputItems:=InputItems, _
            Instructions:="You are a helpful weather assistant. Use tools to answer questions.", _
            ToolsJson:=toolsJson, _
            ToolChoiceJson:=ResponsesToolChoiceAuto(), _
            ParallelToolCalls:=True, _
            MaxOutputTokens:=1024 _
        )

        status = ResponsesExtractStatus(resp)
        Debug.Print "Status: " & status

        ' Check output items for function calls
        Set outputItems = ResponsesExtractOutputItems(resp)

        Dim hasToolCalls As Boolean
        hasToolCalls = False

        If Not outputItems Is Nothing Then
            If outputItems.IsValid Then
                If outputItems.IsArray Then
                    assistantOutput = ""
                    For i = 0 To outputItems.ArrayLength - 1
                        Set item = outputItems.GetArrayItem(i)
                        If Not item Is Nothing Then
                            itemType = CStr(item.GetChildByPath("type").ScalarValue)

                            If StrComp(itemType, "function_call", vbTextCompare) = 0 Then
                                hasToolCalls = True
                                callId = CStr(item.GetChildByPath("call_id").ScalarValue)
                                functionName = CStr(item.GetChildByPath("name").ScalarValue)
                                functionArgs = CStr(item.GetChildByPath("arguments").ScalarValue)

                                Debug.Print "  Tool call: " & functionName
                                Debug.Print "    Args: " & functionArgs

                                toolResult = ExecuteToolFunction(functionName, functionArgs)
                                Debug.Print "    Result: " & toolResult

                                ' Add assistant function_call item and tool result to InputItems
                                ' Responses API uses a different format for tool results
                                ' The function_call output from the model and the function_call_output input
                                InputItems.Add ResponsesBuildFunctionCallResult(callId, toolResult)
                            ElseIf StrComp(itemType, "message", vbTextCompare) = 0 Then
                                ' Collect assistant text
                                Dim contentItem As JsonData
                                Set contentItem = item.GetChildByPath("content.0.text")
                                If Not contentItem Is Nothing Then
                                    If contentItem.IsValid Then
                                        assistantOutput = CStr(contentItem.ScalarValue)
                                    End If
                                End If
                            End If
                        End If
                    Next i
                End If
            End If
        End If

        If hasToolCalls Then
            Debug.Print "Tool calls processed, continuing loop..."
        Else
            Debug.Print "No more tool calls."
            Debug.Print vbCrLf & "Final assistant response:"
            Debug.Print ResponsesExtractText(resp)
            Exit Do
        End If

        Debug.Print
    Loop

    If loopCounter >= maxLoops Then
        Debug.Print "Maximum iterations reached (" & maxLoops & ")"
    End If

    Debug.Print String(80, "=")
    Debug.Print "LOOP COMPLETED"
    Debug.Print String(80, "=")

    Exit Sub

EH:
    Debug.Print "[ERROR] "; Err.Number; " - "; Err.Description
End Sub

' ============================================================================
' Responses API helpers for test
' ============================================================================

Private Function BuildResponsesWeatherToolsJson() As String
    ' Tools in Responses API flat format: {type, name, description, parameters, strict}
    ' No "function" wrapper unlike Chat Completions
    Dim s1 As String
    Dim s2 As String
    Dim result As String

    s1 = "{""type"":""function"","
    s1 = s1 & """name"":""get_current_weather"","
    s1 = s1 & """description"":""Get the current weather in a given location"","
    s1 = s1 & """parameters"":{""type"":""object"","
    s1 = s1 & """properties"":{""location"":{""type"":""string"","
    s1 = s1 & """description"":""The city and state, e.g. San Francisco, CA""},"
    s1 = s1 & """unit"":{""type"":""string"","
    s1 = s1 & """enum"":[""celsius"",""fahrenheit""],"
    s1 = s1 & """description"":""Temperature unit""}},"
    s1 = s1 & """required"":[""location"",""unit""],"
    s1 = s1 & """additionalProperties"":false},"
    s1 = s1 & """strict"":true}"

    s2 = "{""type"":""function"","
    s2 = s2 & """name"":""get_location_coordinates"","
    s2 = s2 & """description"":""Get geographic coordinates for a location"","
    s2 = s2 & """parameters"":{""type"":""object"","
    s2 = s2 & """properties"":{""location"":{""type"":""string"","
    s2 = s2 & """description"":""The city and state, e.g. San Francisco, CA""}},"
    s2 = s2 & """required"":[""location""],"
    s2 = s2 & """additionalProperties"":false},"
    s2 = s2 & """strict"":true}"

    result = "[" & s1 & "," & s2 & "]"
    Debug.Print ""
    Debug.Print "Responses ToolsJson:"
    Debug.Print result
    Debug.Print
    BuildResponsesWeatherToolsJson = result
End Function

Private Function ResponsesBuildFunctionCallResult(ByVal callId As String, ByVal output As String) As String
    ' Responses API format for submitting a function call result:
    ' {"type": "function_call_output", "call_id": "...", "output": "..."}
    ResponsesBuildFunctionCallResult = _
        "{" & _
            """type"":""function_call_output""," & _
            """call_id"":" & JsonString(callId) & "," & _
            """output"":" & JsonString(output) & _
        "}"
End Function

