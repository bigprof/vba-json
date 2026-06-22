VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChatEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================================================
' ChatEngine — State-machine driven chat loop for OpenAI tool-calling
' ============================================================================
' Owns the message history, drives the API call loop, and fires events that
' the UI layer (frmChat) handles.
'
' Usage:
'   Dim WithEvents engine As ChatEngine
'   Set engine = New ChatEngine
'   Set engine.OpenAI = myOpenAIClient
'   Set engine.ToolRegistry = myRegistry
'   engine.SystemPrompt = "You are a helpful assistant."
'   engine.UseAsync = True
'   engine.StartConversation "What's the weather in Boston?"
'   ' ... Timer calls engine.ContinueLoop() ...
'   ' Handle engine.MessageReceived, engine.ToolExecuted, etc.

' --- Public Events ---
Public Event StatusChanged(ByVal Status As String)
Public Event MessageReceived(ByVal Text As String)
Public Event ToolExecuted(ByVal Name As String, ByVal Args As String, ByVal Result As String)
Public Event ErrorOccurred(ByVal Msg As String)
Public Event Done()

' --- Public Properties ---
Public OpenAI As OpenAI
Public ToolRegistry As ToolRegistry
Public SystemPrompt As String
Public UseAsync As Boolean
Public Model As String
Public MaxTokens As Long

' --- State Enum ---
Private Enum ChatEngineState
    Idle = 0
    ReadyToCall = 1
    WaitingForResponse = 2
    Done = 3
End Enum

' --- Private State ---
Private mState As ChatEngineState
Private mMessages As Collection
Private mLastResponseText As String
Private mLoopCounter As Long
Private mMaxLoops As Long

' --- Async HTTP ---
Private mAsyncHttp As AsyncHttpRequest
Private mAsyncUrl As String
Private mAsyncBody As String

Private Sub Class_Initialize()
    Set mMessages = New Collection
    mState = Idle
    UseAsync = False
    Model = "gpt-4o-mini"
    MaxTokens = 1024
    mMaxLoops = 10
    mLoopCounter = 0
End Sub

Private Sub Class_Terminate()
    If Not mAsyncHttp Is Nothing Then
        On Error Resume Next
        mAsyncHttp.Cancel
        Set mAsyncHttp = Nothing
    End If
    Set mMessages = Nothing
End Sub

' ============================================================================
' Public API
' ============================================================================

Public Sub StartConversation(ByVal userMessage As String)
    ' Clear history and begin a new conversation.
    ' Adds system prompt (if set) and the user message, then transitions to
    ' ReadyToCall.  The caller is expected to enable a Timer that drives
    ' ContinueLoop().
    
    If OpenAI Is Nothing Then
        RaiseEvent ErrorOccurred("OpenAI client is not set.")
        mState = Done
        RaiseEvent Done
        Exit Sub
    End If
    
    If ToolRegistry Is Nothing Then
        RaiseEvent ErrorOccurred("ToolRegistry is not set.")
        mState = Done
        RaiseEvent Done
        Exit Sub
    End If
    
    Set mMessages = New Collection
    mLoopCounter = 0
    mLastResponseText = ""
    
    ' Add system/developer message if configured
    If LenB(SystemPrompt) <> 0 Then
        mMessages.Add OpenAIMessageDeveloper(SystemPrompt)
    End If
    
    ' Add user message
    mMessages.Add OpenAIMessageUser(userMessage)
    
    mState = ReadyToCall
    RaiseEvent StatusChanged("Starting conversation...")
End Sub

Public Sub ContinueLoop()
    ' Single-step function called by the Timer.  Advances the state machine
    ' by exactly one step.  Designed to be called from a Timer_Tick event
    ' so the UI message pump gets a chance between ticks.
    
    Select Case mState
        Case Idle
            ' Nothing to do
            
        Case ReadyToCall
            DoApiCall
            
        Case WaitingForResponse
            ' Async request is in flight — poll for completion.
            If Not mAsyncHttp Is Nothing Then
                If mAsyncHttp.Poll Then
                    HandleAsyncResponse
                End If
            End If
            DoEvents
            
        Case Done
            ' Conversation finished — caller should disable the Timer.
    End Select
End Sub

Public Property Get IsActive() As Boolean
    IsActive = (mState = ReadyToCall Or mState = WaitingForResponse)
End Property

Public Property Get IsDone() As Boolean
    IsDone = (mState = Done)
End Property

Public Property Get LastResponseText() As String
    LastResponseText = mLastResponseText
End Property

Public Property Get LoopCount() As Long
    LoopCount = mLoopCounter
End Property

Public Property Get MaxLoops() As Long
    MaxLoops = mMaxLoops
End Property

Public Property Let MaxLoops(ByVal value As Long)
    If value > 0 Then mMaxLoops = value
End Property

' ============================================================================
' Internal: API call dispatching
' ============================================================================

Private Sub DoApiCall()
    Dim ToolsJson As String
    
    ' Guard against infinite loops
    mLoopCounter = mLoopCounter + 1
    If mLoopCounter > mMaxLoops Then
        RaiseEvent ErrorOccurred("Maximum loop iterations reached (" & mMaxLoops & ").")
        mState = Done
        RaiseEvent Done
        Exit Sub
    End If
    
    ' Build tools JSON from the registry
    ToolsJson = ToolRegistry.BuildToolsJson()
    
    RaiseEvent StatusChanged("Calling API (iteration " & mLoopCounter & ")...")
    
    If UseAsync Then
        CallApiAsync ToolsJson
    Else
        CallApiSync ToolsJson
    End If
End Sub

Private Sub CallApiSync(ByVal ToolsJson As String)
    Dim resp As JsonData
    
    On Error GoTo ApiErr
    
    Set resp = OpenAI.CreateChatCompletion( _
        Model:=Model, _
        Messages:=mMessages, _
        MaxCompletionTokens:=MaxTokens, _
        ToolsJson:=ToolsJson, _
        ToolChoiceJson:=OpenAIToolChoiceAuto(), _
        ParallelToolCalls:=True _
    )
    
    EvaluateResponse resp
    Exit Sub

ApiErr:
    RaiseEvent ErrorOccurred("API error: " & Err.Description)
    mState = Done
    RaiseEvent Done
End Sub

Private Sub CallApiAsync(ByVal ToolsJson As String)
    Dim body As String
    Dim headers As Object
    Dim url As String
    
    ' Build the request body same way as CreateChatCompletion does
    body = "{" & _
        """model"":" & JsonString(Model) & _
        ",""messages"":" & CollectionToJsonArray(mMessages) & _
        ",""max_completion_tokens"":" & CStr(MaxTokens) & _
        ",""tools"":" & ToolsJson & _
        ",""tool_choice"":" & OpenAIToolChoiceAuto() & _
        ",""parallel_tool_calls"":true" & _
        "}"
    
    ' Use OpenAI.BuildRequestBody to get URL and auth headers
    Set headers = Nothing
    url = OpenAI.BuildRequestBody("/chat/completions", body, headers)
    
    ' Create async HTTP client if needed
    If mAsyncHttp Is Nothing Then
        Set mAsyncHttp = New AsyncHttpRequest
    End If
    
    ' Start the async request
    On Error GoTo AsyncErr
    mAsyncHttp.Start "POST", url, body, headers
    mState = WaitingForResponse
    Exit Sub

AsyncErr:
    RaiseEvent ErrorOccurred("Async HTTP error: " & Err.Description)
    mState = Done
    RaiseEvent Done
End Sub

' ============================================================================
' Internal: Response evaluation
' ============================================================================

Private Sub EvaluateResponse(ByVal resp As JsonData)
    Dim finishReason As String
    Dim toolCalls As JsonData
    
    ' Extract finish reason
    finishReason = OpenAIExtractFinishReason(resp)
    
    ' Check for tool calls
    If StrComp(finishReason, "tool_calls", vbTextCompare) = 0 Then
        Set toolCalls = OpenAIExtractToolCalls(resp)
        
        If Not toolCalls Is Nothing Then
            If toolCalls.IsValid Then
                ' Add the assistant response (with tool_calls) to message history
                mMessages.Add OpenAIMessageAssistantWithToolCalls( _
                    OpenAIExtractText(resp), toolCalls.ToJSON())
                
                ' Process each tool call
                ProcessToolCalls toolCalls
                
                ' Stay in ReadyToCall for the next iteration
                mState = ReadyToCall
                Exit Sub
            End If
        End If
        
        ' tool_calls finish_reason but couldn't parse — treat as error
        RaiseEvent ErrorOccurred("Finish reason was 'tool_calls' but tool calls could not be parsed.")
        mState = Done
        RaiseEvent Done
        Exit Sub
    End If
    
    ' No more tool calls — we have the final response
    mLastResponseText = OpenAIExtractText(resp)
    
    ' Add the assistant text response to message history
    If LenB(mLastResponseText) <> 0 Then
        mMessages.Add OpenAIMessageAssistant(mLastResponseText)
    End If
    
    mState = Done
    RaiseEvent MessageReceived(mLastResponseText)
    RaiseEvent Done
End Sub

' ============================================================================
' Internal: Tool call processing
' ============================================================================

Private Sub ProcessToolCalls(ByVal toolCalls As JsonData)
    Dim i As Long
    Dim ii As String
    Dim toolCount As Long
    Dim toolName As String
    Dim toolInput As String
    Dim toolResult As String
    Dim toolId As String
    
    toolCount = toolCalls.ArrayLength
    
    For i = 1 To toolCount
        ii = CStr(i - 1) ' JSON arrays are 0-indexed
        
        On Error Resume Next
        toolId = CStr(toolCalls.GetChildByPath(ii & ".id").ScalarValue)
        toolName = CStr(toolCalls.GetChildByPath(ii & ".function.name").ScalarValue)
        toolInput = CStr(toolCalls.GetChildByPath(ii & ".function.arguments").ScalarValue)
        On Error GoTo 0
        
        If LenB(toolName) = 0 Then
            RaiseEvent ErrorOccurred("Tool call #" & i & " has no function name.")
        Else
            ' Execute via ToolRegistry
            toolResult = ToolRegistry.Execute(toolName, toolInput)
            
            ' Fire event for UI
            RaiseEvent ToolExecuted(toolName, toolInput, toolResult)
            
            ' Add tool result to message history
            mMessages.Add OpenAIMessageTool(toolId, toolResult)
        End If
    Next i
End Sub

' ============================================================================
' Async HTTP event handlers
' ============================================================================

Private Sub HandleAsyncResponse()
    Dim parsed As JsonData
    Dim apiError As String
    Dim status As Long
    Dim responseText As String
    
    status = mAsyncHttp.StatusCode
    responseText = mAsyncHttp.ResponseText
    
    On Error GoTo ParseErr
    
    If status < 200 Or status >= 300 Then
        Set parsed = ParseJSON(responseText)
        apiError = ""
        If Not parsed Is Nothing Then
            If parsed.IsValid Then
                apiError = OpenAIExtractErrorMessage(parsed)
            End If
        End If
        
        If LenB(apiError) <> 0 Then
            RaiseEvent ErrorOccurred("HTTP " & CStr(status) & ": " & apiError)
        Else
            RaiseEvent ErrorOccurred("HTTP " & CStr(status) & ": " & responseText)
        End If
        
        mState = Done
        RaiseEvent Done
        Exit Sub
    End If
    
    Set parsed = ParseJSON(responseText)
    If parsed Is Nothing Then
        RaiseEvent ErrorOccurred("Response could not be parsed as JSON.")
        mState = Done
        RaiseEvent Done
        Exit Sub
    End If
    
    If Not parsed.IsValid Then
        RaiseEvent ErrorOccurred("Response is not valid JSON.")
        mState = Done
        RaiseEvent Done
        Exit Sub
    End If
    
    EvaluateResponse parsed
    Exit Sub

ParseErr:
    RaiseEvent ErrorOccurred("Error parsing async response: " & Err.Description)
    mState = Done
    RaiseEvent Done
End Sub
