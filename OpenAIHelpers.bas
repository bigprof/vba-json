Attribute VB_Name = "OpenAIHelpers"
Option Explicit

Public Function OpenAIMessageDeveloper(ByVal Text As String) As String
    OpenAIMessageDeveloper = BuildMessageText("developer", Text)
End Function

Public Function OpenAIMessageUser(ByVal Text As String) As String
    OpenAIMessageUser = BuildMessageText("user", Text)
End Function

Public Function OpenAIMessageAssistant(ByVal Text As String) As String
    OpenAIMessageAssistant = BuildMessageText("assistant", Text)
End Function

Public Function OpenAIMessageText(ByVal Role As String, ByVal Text As String) As String
    OpenAIMessageText = BuildMessageText(Role, Text)
End Function

Public Function BuildMessageText(ByVal Role As String, ByVal Text As String) As String
    BuildMessageText = _
        "{" & _
            """role"":" & JsonString(Role) & "," & _
            """content"":" & JsonString(Text) & _
        "}"
End Function

Public Function JsonString(ByVal Value As String) As String
    Dim i As Long
    Dim ch As String
    Dim code As Long
    Dim s As String
    
    s = """"
    
    For i = 1 To Len(Value)
        ch = Mid$(Value, i, 1)
        Select Case ch
            Case "\"
                s = s & "\\"
            Case """"
                s = s & "\"""
            Case vbBack
                s = s & "\b"
            Case vbFormFeed
                s = s & "\f"
            Case vbLf
                s = s & "\n"
            Case vbCr
                s = s & "\r"
            Case vbTab
                s = s & "\t"
            Case Else
                code = Asc(ch)
                If code >= 0 And code <= 31 Then
                    s = s & "\u" & Right$("0000" & Hex$(code), 4)
                Else
                    s = s & ch
                End If
        End Select
    Next
    
    s = s & """"
    JsonString = s
End Function

Public Function JsonBoolean(ByVal Value As Boolean) As String
    If Value Then
        JsonBoolean = "true"
    Else
        JsonBoolean = "false"
    End If
End Function

Public Function JsonNumber(ByVal Value As Double) As String
    Dim s As String
    
    s = Trim$(str$(Value))
    s = Replace$(s, ",", ".")
    
    JsonNumber = s
End Function

Public Function CollectionToJsonArray(ByVal Items As Collection) As String
    Dim s As String
    Dim i As Long
    
    If Items Is Nothing Then
        CollectionToJsonArray = "[]"
        Exit Function
    End If
    
    s = "["
    For i = 1 To Items.Count
        If i > 1 Then s = s & ","
        s = s & CStr(Items.Item(i))
    Next
    s = s & "]"
    
    CollectionToJsonArray = s
End Function

Public Function OpenAIExtractText(ByVal Response As JsonData) As String
    Dim node As JsonData
    
    Set node = Response.GetChildByPath("choices.0.message.content")
    If node Is Nothing Then Exit Function
    If Not node.IsValid Then Exit Function
    If Not node.IsScalar Then Exit Function
    If IsNull(node.ScalarValue) Then Exit Function
    
    OpenAIExtractText = CStr(node.ScalarValue)
End Function

Public Function OpenAIExtractFinishReason(ByVal Response As JsonData) As String
    Dim node As JsonData
    
    Set node = Response.GetChildByPath("choices.0.finish_reason")
    If node Is Nothing Then Exit Function
    If Not node.IsValid Then Exit Function
    If Not node.IsScalar Then Exit Function
    If IsNull(node.ScalarValue) Then Exit Function
    
    OpenAIExtractFinishReason = CStr(node.ScalarValue)
End Function

Public Function OpenAIExtractErrorMessage(ByVal Response As JsonData) As String
    Dim node As JsonData
    Dim typ As JsonData
    Dim code As JsonData
    Dim msg As JsonData
    Dim s As String
    
    Set msg = Response.GetChildByPath("error.message")
    If Not msg Is Nothing Then
        If msg.IsValid Then
            If msg.IsScalar Then
                If Not IsNull(msg.ScalarValue) Then
                    s = CStr(msg.ScalarValue)
                End If
            End If
        End If
    End If
    
    Set typ = Response.GetChildByPath("error.type")
    If Not typ Is Nothing Then
        If typ.IsValid Then
            If typ.IsScalar Then
                If Not IsNull(typ.ScalarValue) Then
                    If LenB(s) <> 0 Then
                        s = s & vbCrLf
                    End If
                    s = s & "type: " & CStr(typ.ScalarValue)
                End If
            End If
        End If
    End If
    
    Set code = Response.GetChildByPath("error.code")
    If Not code Is Nothing Then
        If code.IsValid Then
            If code.IsScalar Then
                If Not IsNull(code.ScalarValue) Then
                    If LenB(s) <> 0 Then
                        s = s & vbCrLf
                    End If
                    s = s & "code: " & CStr(code.ScalarValue)
                End If
            End If
        End If
    End If
    
    OpenAIExtractErrorMessage = s
End Function

Public Function OpenAIExtractToolCalls(ByVal Response As JsonData) As JsonData
    Set OpenAIExtractToolCalls = Response.GetChildByPath("choices.0.message.tool_calls")
End Function

Public Function OpenAIResponseFormatJsonObject() As String
    OpenAIResponseFormatJsonObject = "{""type"":""json_object""}"
End Function

Public Function OpenAIResponseFormatText() As String
    OpenAIResponseFormatText = "{""type"":""text""}"
End Function

Public Function OpenAIResponseFormatJsonSchema( _
    ByVal Name As String, _
    ByVal SchemaJson As String, _
    Optional ByVal Description As String = "", _
    Optional ByVal Strict As Boolean = False _
) As String
    
    Dim s As String
    
    s = "{"
    s = s & """type"":""json_schema"""
    s = s & ",""json_schema"":{"
    s = s & """name"":" & JsonString(Name)
    
    If LenB(Description) <> 0 Then
        s = s & ",""description"":" & JsonString(Description)
    End If
    
    s = s & ",""schema"":" & SchemaJson
    s = s & ",""strict"":" & JsonBoolean(Strict)
    s = s & "}"
    s = s & "}"
    
    OpenAIResponseFormatJsonSchema = s
End Function

Public Function OpenAIToolChoiceAuto() As String
    OpenAIToolChoiceAuto = """auto"""
End Function

Public Function OpenAIToolChoiceNone() As String
    OpenAIToolChoiceNone = """none"""
End Function

Public Function OpenAIToolChoiceRequired() As String
    OpenAIToolChoiceRequired = """required"""
End Function

