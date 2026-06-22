VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ToolRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================================================
' ToolRegistry — Maps tool names to IToolHandler implementations
' ============================================================================
' Generates the OpenAI tools JSON array and dispatches tool execution.
'
' Usage:
'   Dim reg As ToolRegistry
'   Set reg = New ToolRegistry
'   reg.Register New WeatherToolHandler
'   reg.Register New CoordinatesToolHandler
'   Dim toolsJson As String
'   toolsJson = reg.BuildToolsJson()
'   Dim result As String
'   result = reg.Execute("get_current_weather", "{""location"":""Boston""}")

Private mHandlers As Object ' Scripting.Dictionary

Private Sub Class_Initialize()
    Set mHandlers = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    If Not mHandlers Is Nothing Then
        mHandlers.RemoveAll
        Set mHandlers = Nothing
    End If
End Sub

Public Sub Register(ByVal handler As IToolHandler)
    Dim name As String
    
    If handler Is Nothing Then
        Err.Raise vbObjectError + 3001, "ToolRegistry.Register", "Handler must not be Nothing."
    End If
    
    name = handler.GetName()
    
    If LenB(Trim$(name)) = 0 Then
        Err.Raise vbObjectError + 3002, "ToolRegistry.Register", "Handler GetName() returned an empty string."
    End If
    
    If mHandlers.Exists(name) Then
        Err.Raise vbObjectError + 3003, "ToolRegistry.Register", _
            "A handler named '" & name & "' is already registered."
    End If
    
    mHandlers.Add handler, name
End Sub

Public Function Execute(ByVal toolName As String, ByVal argsJson As String) As String
    Dim handler As IToolHandler
    
    If Not mHandlers.Exists(toolName) Then
        Execute = "{""error"":""Unknown tool: " & toolName & """}"
        Exit Function
    End If
    
    Set handler = mHandlers.Item(toolName)
    
    On Error GoTo ExecErr
    Execute = handler.Execute(argsJson)
    Exit Function

ExecErr:
    Execute = "{""error"":""Tool execution failed: " & Replace$(Err.Description, """", "'") & """}"
End Function

Public Function BuildToolsJson() As String
    ' Build the OpenAI Chat Completions "tools" JSON array.
    ' Format: [{"type":"function","function":{"name":"...","description":"...","parameters":{...},"strict":true}}, ...]
    Dim s As String
    Dim name As Variant
    Dim handler As IToolHandler
    Dim first As Boolean
    Dim schemaJson As String
    Dim strictVal As Boolean
    
    first = True
    s = "["
    
    For Each name In mHandlers.Keys
        Set handler = mHandlers.Item(name)
        
        If Not first Then
            s = s & ","
        End If
        first = False
        
        ' Get the parameters schema; if empty, use a minimal default
        schemaJson = handler.GetParametersSchemaJson()
        If LenB(schemaJson) = 0 Then
            schemaJson = "{""type"":""object"",""properties"":{},""additionalProperties"":false}"
        End If
        
        ' Strict defaults to True
        strictVal = True
        On Error Resume Next
        strictVal = handler.Strict
        On Error GoTo 0
        
        s = s & "{"
        s = s & """type"":""function"","
        s = s & """function"":{"
        s = s & """name"":" & JsonString(CStr(name)) & ","
        s = s & """description"":" & JsonString(handler.GetDescription()) & ","
        s = s & """parameters"":" & schemaJson & ","
        s = s & """strict"":" & JsonBoolean(strictVal)
        s = s & "}"
        s = s & "}"
    Next name
    
    s = s & "]"
    
    BuildToolsJson = s
End Function

Public Property Get Count() As Long
    Count = mHandlers.Count
End Property

Public Function GetHandler(ByVal toolName As String) As IToolHandler
    If mHandlers.Exists(toolName) Then
        Set GetHandler = mHandlers.Item(toolName)
    End If
End Function

Public Sub Clear()
    mHandlers.RemoveAll
End Sub
