Attribute VB_Name = "Tester"
Option Explicit

Public Sub Main()
    Dim root As JsonData
    Dim JsonText As String
    
    JsonText = _
        "{" & _
        """user"":{" & _
            """name"":""Alice""," & _
            """age"":30," & _
            """active"":true," & _
            """middleName"":null" & _
        "}," & _
        """items"":[" & _
            "{""value"":123}," & _
            "{""value"":456}," & _
            """hello""" & _
        "]" & _
        "}"
    
    Debug.Print String(60, "=")
    Debug.Print "JsonData VB6 test starting"
    Debug.Print String(60, "=")
    Debug.Print "JSON:"
    Debug.Print JsonText
    Debug.Print ""
    
    On Error GoTo ParseErr
    Set root = ParseJSON(JsonText)
    On Error GoTo 0
    
    If root Is Nothing Then
        Debug.Print "[FAIL] ParseJSON returned Nothing"
        Exit Sub
    End If
    
    If root.IsValid Then
        Debug.Print "[PASS] Root is valid"
    Else
        Debug.Print "[FAIL] Root is invalid"
        Exit Sub
    End If
    
    If root.IsObject Then
        Debug.Print "[PASS] Root is an object"
    Else
        Debug.Print "[FAIL] Root is not an object"
    End If
    
    Debug.Print ""
    Debug.Print "Running tests..."
    Debug.Print ""
    
    TestScalar root, "user.name", "Alice", "String value"
    TestScalar root, "user.age", "30", "Number value"
    TestScalar root, "user.active", "True", "Boolean value"
    TestNull root, "user.middleName", "Null value"
    
    TestScalar root, "items.0.value", "123", "Array object item 0"
    TestScalar root, "items.1.value", "456", "Array object item 1"
    TestScalar root, "items.2", "hello", "Array scalar item"
    
    TestMissing root, "user.lastname", "Missing object property"
    TestMissing root, "items.3", "Missing array element"
    TestMissing root, "items.2.value", "Path continues through scalar"
    TestMissing root, "does.not.exist", "Missing nested path"
    
    Debug.Print ""
    Debug.Print "Serialized back to JSON:"
    Debug.Print root.ToJSON("  ")
    Debug.Print ""
    Debug.Print "Done."
    
    TestOpenAISimple
    TestOpenAI_MultiTurn
    TestOpenAI_JsonSchema
    TestOpenAI_JsonObject
    TestOpenAI_FunctionToolCall_RequestOnly
    TestOpenAI_MultiTurn_UsesPriorAnswer
    
    MsgBox "Done. Check the Immediate Window.", vbInformation
    Exit Sub

ParseErr:
    Debug.Print "[FAIL] Parse error: " & Err.Number & " - " & Err.Description
    MsgBox "Parse error: " & Err.Description, vbExclamation
End Sub

Private Sub TestScalar(ByVal root As JsonData, ByVal path As String, ByVal expected As String, ByVal label As String)
    Dim node As JsonData
    Dim actual As String
    
    On Error GoTo EH
    
    Set node = root.GetChildByPath(path)
    
    If node Is Nothing Then
        Debug.Print "[FAIL] " & label & " - node is Nothing for path: " & path
        Exit Sub
    End If
    
    If Not node.IsValid Then
        Debug.Print "[FAIL] " & label & " - invalid node for path: " & path
        Exit Sub
    End If
    
    If Not node.IsScalar Then
        Debug.Print "[FAIL] " & label & " - node is not scalar for path: " & path
        Exit Sub
    End If
    
    actual = ScalarToString(node)
    
    If StrComp(actual, expected, vbBinaryCompare) = 0 Then
        Debug.Print "[PASS] " & label & " - " & path & " = " & actual
    Else
        Debug.Print "[FAIL] " & label & " - " & path & " expected [" & expected & "] but got [" & actual & "]"
    End If
    
    Exit Sub

EH:
    Debug.Print "[FAIL] " & label & " - error at path " & path & ": " & Err.Number & " - " & Err.Description
End Sub

Private Sub TestNull(ByVal root As JsonData, ByVal path As String, ByVal label As String)
    Dim node As JsonData
    
    On Error GoTo EH
    
    Set node = root.GetChildByPath(path)
    
    If node Is Nothing Then
        Debug.Print "[FAIL] " & label & " - node is Nothing for path: " & path
        Exit Sub
    End If
    
    If Not node.IsValid Then
        Debug.Print "[FAIL] " & label & " - invalid node for path: " & path
        Exit Sub
    End If
    
    If Not node.IsScalar Then
        Debug.Print "[FAIL] " & label & " - node is not scalar for path: " & path
        Exit Sub
    End If
    
    If IsNull(node.ScalarValue) Then
        Debug.Print "[PASS] " & label & " - " & path & " is Null"
    Else
        Debug.Print "[FAIL] " & label & " - " & path & " expected Null but got [" & ScalarToString(node) & "]"
    End If
    
    Exit Sub

EH:
    Debug.Print "[FAIL] " & label & " - error at path " & path & ": " & Err.Number & " - " & Err.Description
End Sub

Private Sub TestMissing(ByVal root As JsonData, ByVal path As String, ByVal label As String)
    Dim node As JsonData
    
    On Error GoTo EH
    
    Set node = root.GetChildByPath(path)
    
    If node Is Nothing Then
        Debug.Print "[PASS] " & label & " - node is Nothing for missing path: " & path
        Exit Sub
    End If
    
    If Not node.IsValid Then
        Debug.Print "[PASS] " & label & " - invalid node returned for missing path: " & path
    Else
        Debug.Print "[FAIL] " & label & " - expected missing path, but node is valid: " & path
    End If
    
    Exit Sub

EH:
    Debug.Print "[FAIL] " & label & " - error at path " & path & ": " & Err.Number & " - " & Err.Description
End Sub

Private Function ScalarToString(ByVal node As JsonData) As String
    If IsNull(node.ScalarValue) Then
        ScalarToString = "Null"
    Else
        ScalarToString = CStr(node.ScalarValue)
    End If
End Function

