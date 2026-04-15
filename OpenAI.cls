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

Private Sub Class_Initialize()
    mBaseUrl = DEFAULT_BASE_URL
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

Public Function CreateChatCompletionSimple( _
    ByVal Model As String, _
    ByVal DeveloperPrompt As String, _
    ByVal UserPrompt As String _
) As JsonData
    
    Dim body As String
    
    ValidateConfiguration
    
    body = "{""model"":" & JsonString(Model) & _
           ",""messages"":[" & _
           "{""role"":""developer"",""content"":" & JsonString(DeveloperPrompt) & "}," & _
           "{""role"":""user"",""content"":" & JsonString(UserPrompt) & "}" & _
           "]}"
    
    Set CreateChatCompletionSimple = PostJson(ENDPOINT_CHAT_COMPLETIONS, body)
End Function

Public Function CreateChatCompletion( _
    ByVal Model As String, _
    ByVal Messages As Collection, _
    Optional ByVal Temperature As Variant = Empty, _
    Optional ByVal MaxCompletionTokens As Variant = Empty, _
    Optional ByVal Verbosity As Variant = Empty, _
    Optional ByVal ResponseFormatJson As Variant = Empty, _
    Optional ByVal ToolsJson As Variant = Empty, _
    Optional ByVal ToolChoiceJson As Variant = Empty, _
    Optional ByVal ReasoningEffort As Variant = "low", _
    Optional ByVal Store As Variant = False, _
    Optional ByVal MetadataJson As Variant = Empty, _
    Optional ByVal ParallelToolCalls As Variant = Empty _
) As JsonData
    
    Dim body As String
    
    ValidateConfiguration
    ValidateCreateChatCompletionInputs Model, Messages
    
    body = "{"
    body = body & """model"":" & JsonString(Model)
    body = body & ",""messages"":" & CollectionToJsonArray(Messages)
    
    If Not IsEmpty(Temperature) Then AppendOptionalJsonDouble body, "temperature", Temperature
    If Not IsEmpty(MaxCompletionTokens) Then AppendOptionalJsonLong body, "max_completion_tokens", MaxCompletionTokens
    If Not IsEmpty(Verbosity) Then AppendOptionalJsonString body, "verbosity", Verbosity
    If Not IsEmpty(ResponseFormatJson) Then AppendOptionalRawJson body, "response_format", ResponseFormatJson
    If Not IsEmpty(ToolsJson) Then AppendOptionalRawJson body, "tools", ToolsJson
    If Not IsEmpty(ToolChoiceJson) Then AppendOptionalRawJson body, "tool_choice", ToolChoiceJson
    If Not IsEmpty(ReasoningEffort) Then AppendOptionalJsonString body, "reasoning_effort", ReasoningEffort
    If Not IsEmpty(Store) Then AppendOptionalJsonBoolean body, "store", Store
    If Not IsEmpty(MetadataJson) Then AppendOptionalRawJson body, "metadata", MetadataJson
    If Not IsEmpty(ParallelToolCalls) Then AppendOptionalJsonBoolean body, "parallel_tool_calls", ParallelToolCalls
    
    body = body & "}"
    
    Set CreateChatCompletion = PostJson(ENDPOINT_CHAT_COMPLETIONS, body)
End Function

Private Function PostJson(ByVal RelativeUrl As String, ByVal body As String) As JsonData
    Dim http As MSXML2.XMLHTTP60
    Dim url As String
    Dim responseText As String
    Dim parsed As JsonData
    Dim apiError As String
    
    url = mBaseUrl & RelativeUrl
    
    Set http = New MSXML2.XMLHTTP60
    
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
    
    http.send body
    
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
                "HTTP " & CStr(http.Status) & " " & http.statusText & vbCrLf & _
                apiError
        Else
            Err.Raise vbObjectError + 1102, "OpenAI.PostJson", _
                "HTTP " & CStr(http.Status) & " " & http.statusText & vbCrLf & _
                responseText
        End If
    End If
    
    If parsed Is Nothing Then
        Err.Raise vbObjectError + 1103, "OpenAI.PostJson", _
            "The response body could not be parsed as JSON." & vbCrLf & _
            responseText
    End If
    
    If Not parsed.IsValid Then
        Err.Raise vbObjectError + 1104, "OpenAI.PostJson", _
            "The response body is not valid JSON." & vbCrLf & _
            responseText
    End If
    
    Set PostJson = parsed
    Exit Function

TransportErr:
    Err.Raise Err.Number, "OpenAI.PostJson", _
        "Transport error during HTTPS request." & vbCrLf & _
        "URL: " & url & vbCrLf & _
        "Err.Number: " & CStr(Err.Number) & vbCrLf & _
        "Hex: 0x" & Right$("00000000" & Hex$(Err.Number), 8) & vbCrLf & _
        "Description: " & Err.Description
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

Private Sub AppendOptionalRawJson(ByRef body As String, ByVal fieldName As String, ByVal value As Variant)
    If IsEmpty(value) Then Exit Sub
    If IsNull(value) Then Exit Sub
    
    On Error GoTo BadValue
    
    If VarType(value) = vbString Then
        If LenB(CStr(value)) = 0 Then Exit Sub
        body = body & ",""" & fieldName & """:" & CStr(value)
        Exit Sub
    End If
    
    Err.Raise vbObjectError + 2001, "OpenAI.AppendOptionalRawJson", _
        "Optional field '" & fieldName & "' must be a JSON string."
    Exit Sub

BadValue:
    Err.Raise vbObjectError + 2002, "OpenAI.AppendOptionalRawJson", _
        "Invalid optional field '" & fieldName & "': " & Err.Description
End Sub

Private Sub AppendOptionalJsonString(ByRef body As String, ByVal fieldName As String, ByVal value As Variant)
    If IsEmpty(value) Then Exit Sub
    If IsNull(value) Then Exit Sub
    
    On Error GoTo BadValue
    
    If LenB(Trim$(CStr(value))) = 0 Then Exit Sub
    body = body & ",""" & fieldName & """:" & JsonString(CStr(value))
    Exit Sub

BadValue:
    Err.Raise vbObjectError + 2003, "OpenAI.AppendOptionalJsonString", _
        "Invalid optional string field '" & fieldName & "': " & Err.Description
End Sub

Private Sub AppendOptionalJsonBoolean(ByRef body As String, ByVal fieldName As String, ByVal value As Variant)
    If IsEmpty(value) Then Exit Sub
    If IsNull(value) Then Exit Sub
    
    On Error GoTo BadValue
    
    body = body & ",""" & fieldName & """:" & JsonBoolean(CBool(value))
    Exit Sub

BadValue:
    Err.Raise vbObjectError + 2004, "OpenAI.AppendOptionalJsonBoolean", _
        "Invalid optional boolean field '" & fieldName & "': " & Err.Description
End Sub

Private Sub AppendOptionalJsonLong(ByRef body As String, ByVal fieldName As String, ByVal value As Variant)
    If IsEmpty(value) Then Exit Sub
    If IsNull(value) Then Exit Sub
    
    On Error GoTo BadValue
    
    body = body & ",""" & fieldName & """:" & CStr(CLng(value))
    Exit Sub

BadValue:
    Err.Raise vbObjectError + 2005, "OpenAI.AppendOptionalJsonLong", _
        "Invalid optional numeric field '" & fieldName & "': " & Err.Description
End Sub

Private Sub AppendOptionalJsonDouble(ByRef body As String, ByVal fieldName As String, ByVal value As Variant)
    If IsEmpty(value) Then Exit Sub
    If IsNull(value) Then Exit Sub
    
    On Error GoTo BadValue
    
    body = body & ",""" & fieldName & """:" & JsonNumber(CDbl(value))
    Exit Sub

BadValue:
    Err.Raise vbObjectError + 2006, "OpenAI.AppendOptionalJsonDouble", _
        "Invalid optional numeric field '" & fieldName & "': " & Err.Description
End Sub

