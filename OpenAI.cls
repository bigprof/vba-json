VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OpenAI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const DEFAULT_BASE_URL As String = "https://api.openai.com/v1"
Private Const ENDPOINT_CHAT_COMPLETIONS As String = "/chat/completions"

Private mApiKey As String
Private mBaseUrl As String
Private mOrganization As String
Private mProject As String

Private mTimeoutResolve As Long
Private mTimeoutConnect As Long
Private mTimeoutSend As Long
Private mTimeoutReceive As Long

Private Sub Class_Initialize()
    mBaseUrl = DEFAULT_BASE_URL
    
    ' Timeouts are in milliseconds
    mTimeoutResolve = 30000
    mTimeoutConnect = 30000
    mTimeoutSend = 30000
    mTimeoutReceive = 120000
End Sub

Public Property Get ApiKey() As String
    ApiKey = mApiKey
End Property

Public Property Let ApiKey(ByVal value As String)
    mApiKey = Trim$(value)
End Property

Public Property Get BaseUrl() As String
    BaseUrl = mBaseUrl
End Property

Public Property Let BaseUrl(ByVal value As String)
    value = Trim$(value)
    If LenB(value) <> 0 Then
        If Right$(value, 1) = "/" Then
            value = Left$(value, Len(value) - 1)
        End If
        mBaseUrl = value
    End If
End Property

Public Property Get Organization() As String
    Organization = mOrganization
End Property

Public Property Let Organization(ByVal value As String)
    mOrganization = Trim$(value)
End Property

Public Property Get Project() As String
    Project = mProject
End Property

Public Property Let Project(ByVal value As String)
    mProject = Trim$(value)
End Property

Public Property Get TimeoutResolve() As Long
    TimeoutResolve = mTimeoutResolve
End Property

Public Property Let TimeoutResolve(ByVal value As Long)
    mTimeoutResolve = value
End Property

Public Property Get TimeoutConnect() As Long
    TimeoutConnect = mTimeoutConnect
End Property

Public Property Let TimeoutConnect(ByVal value As Long)
    mTimeoutConnect = value
End Property

Public Property Get TimeoutSend() As Long
    TimeoutSend = mTimeoutSend
End Property

Public Property Let TimeoutSend(ByVal value As Long)
    mTimeoutSend = value
End Property

Public Property Get TimeoutReceive() As Long
    TimeoutReceive = mTimeoutReceive
End Property

Public Property Let TimeoutReceive(ByVal value As Long)
    mTimeoutReceive = value
End Property

Public Function CreateChatCompletionSimple( _
    ByVal Model As String, _
    ByVal DeveloperPrompt As String, _
    ByVal UserPrompt As String, _
    Optional ByVal Temperature As Variant, _
    Optional ByVal MaxCompletionTokens As Variant, _
    Optional ByVal Verbosity As Variant, _
    Optional ByVal ReasoningEffort As Variant _
) As JsonData
    
    Dim Messages As Collection
    
    Set Messages = New Collection
    Messages.Add OpenAIMessageDeveloper(DeveloperPrompt)
    Messages.Add OpenAIMessageUser(UserPrompt)
    
    Set CreateChatCompletionSimple = CreateChatCompletion( _
        Model:=Model, _
        Messages:=Messages, _
        Temperature:=Temperature, _
        MaxCompletionTokens:=MaxCompletionTokens, _
        Verbosity:=Verbosity, _
        ReasoningEffort:=ReasoningEffort _
    )
End Function

Public Function CreateChatCompletion( _
    ByVal Model As String, _
    ByVal Messages As Collection, _
    Optional ByVal Temperature As Variant, _
    Optional ByVal MaxCompletionTokens As Variant, _
    Optional ByVal Verbosity As Variant, _
    Optional ByVal ResponseFormatJson As Variant, _
    Optional ByVal ToolsJson As Variant, _
    Optional ByVal ToolChoiceJson As Variant, _
    Optional ByVal ReasoningEffort As Variant, _
    Optional ByVal Store As Variant, _
    Optional ByVal MetadataJson As Variant, _
    Optional ByVal ParallelToolCalls As Variant _
) As JsonData
    
    Dim Body As String
    
    ValidateConfiguration
    ValidateCreateChatCompletionInputs Model, Messages
    
    Body = "{"
    Body = Body & """model"":" & JsonString(Model)
    Body = Body & ",""messages"":" & CollectionToJsonArray(Messages)
    
    If Not IsEmpty(Temperature) Then
        Body = Body & ",""temperature"":" & JsonNumber(CDbl(Temperature))
    End If
    
    If Not IsEmpty(MaxCompletionTokens) Then
        Body = Body & ",""max_completion_tokens"":" & CStr(CLng(MaxCompletionTokens))
    End If
    
    If Not IsEmpty(Verbosity) Then
        Body = Body & ",""verbosity"":" & JsonString(CStr(Verbosity))
    End If
    
    If Not IsEmpty(ResponseFormatJson) Then
        If LenB(CStr(ResponseFormatJson)) <> 0 Then
            Body = Body & ",""response_format"":" & CStr(ResponseFormatJson)
        End If
    End If
    
    If Not IsEmpty(ToolsJson) Then
        If LenB(CStr(ToolsJson)) <> 0 Then
            Body = Body & ",""tools"":" & CStr(ToolsJson)
        End If
    End If
    
    If Not IsEmpty(ToolChoiceJson) Then
        If LenB(CStr(ToolChoiceJson)) <> 0 Then
            Body = Body & ",""tool_choice"":" & CStr(ToolChoiceJson)
        End If
    End If
    
    If Not IsEmpty(ReasoningEffort) Then
        If LenB(CStr(ReasoningEffort)) <> 0 Then
            Body = Body & ",""reasoning_effort"":" & JsonString(CStr(ReasoningEffort))
        End If
    End If
    
    If Not IsEmpty(Store) Then
        Body = Body & ",""store"":" & JsonBoolean(CBool(Store))
    End If
    
    If Not IsEmpty(MetadataJson) Then
        If LenB(CStr(MetadataJson)) <> 0 Then
            Body = Body & ",""metadata"":" & CStr(MetadataJson)
        End If
    End If
    
    If Not IsEmpty(ParallelToolCalls) Then
        Body = Body & ",""parallel_tool_calls"":" & JsonBoolean(CBool(ParallelToolCalls))
    End If
    
    Body = Body & "}"
    
    Set CreateChatCompletion = PostJson(ENDPOINT_CHAT_COMPLETIONS, Body)
End Function

Private Function PostJson(ByVal RelativeUrl As String, ByVal Body As String) As JsonData
    Dim http As MSXML2.ServerXMLHTTP
    Dim url As String
    Dim responseText As String
    Dim parsed As JsonData
    Dim apiError As String
    
    url = mBaseUrl & RelativeUrl
    
    Set http = New MSXML2.ServerXMLHTTP
    http.setTimeouts mTimeoutResolve, mTimeoutConnect, mTimeoutSend, mTimeoutReceive
    
    On Error GoTo TransportErr
    
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & mApiKey
    
    If LenB(mOrganization) <> 0 Then
        http.setRequestHeader "OpenAI-Organization", mOrganization
    End If
    
    If LenB(mProject) <> 0 Then
        http.setRequestHeader "OpenAI-Project", mProject
    End If
    
    http.send Body
    
    responseText = http.responseText
    
    Set parsed = ParseJSON(responseText)
    
    If http.Status < 200 Or http.Status >= 300 Then
        apiError = ""
        
        If Not parsed Is Nothing Then
            If parsed.IsValid Then
                apiError = OpenAIExtractErrorMessage(parsed)
            End If
        End If
        
        If LenB(apiError) <> 0 Then
            Err.Raise vbObjectError + 1101, "OpenAI.PostJson", _
                "HTTP " & CStr(http.Status) & " " & http.statusText & vbCrLf & apiError
        Else
            Err.Raise vbObjectError + 1102, "OpenAI.PostJson", _
                "HTTP " & CStr(http.Status) & " " & http.statusText & vbCrLf & responseText
        End If
    End If
    
    If parsed Is Nothing Then
        Err.Raise vbObjectError + 1103, "OpenAI.PostJson", _
            "The response body could not be parsed as JSON." & vbCrLf & responseText
    End If
    
    If Not parsed.IsValid Then
        Err.Raise vbObjectError + 1104, "OpenAI.PostJson", _
            "The response body is not valid JSON." & vbCrLf & responseText
    End If
    
    Set PostJson = parsed
    Exit Function

TransportErr:
    Err.Raise Err.Number, "OpenAI.PostJson", Err.Description
End Function

Private Sub ValidateConfiguration()
    If LenB(mApiKey) = 0 Then
        Err.Raise vbObjectError + 1100, "OpenAI", "ApiKey is required."
    End If
End Sub

Private Sub ValidateCreateChatCompletionInputs(ByVal Model As String, ByVal Messages As Collection)
    If LenB(Trim$(Model)) = 0 Then
        Err.Raise vbObjectError + 1110, "OpenAI.CreateChatCompletion", "Model is required."
    End If
    
    If Messages Is Nothing Then
        Err.Raise vbObjectError + 1111, "OpenAI.CreateChatCompletion", "Messages collection is required."
    End If
    
    If Messages.Count = 0 Then
        Err.Raise vbObjectError + 1112, "OpenAI.CreateChatCompletion", "Messages collection must not be empty."
    End If
End Sub

