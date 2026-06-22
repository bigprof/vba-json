VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IToolHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================================================
' IToolHandler — Abstract interface for tool function handlers
' ============================================================================
' Implement this interface for each tool you want to register with ToolRegistry.
'
' Usage pattern:
'   Implements IToolHandler   ' in your tool class
'   Private Function IToolHandler_GetName() As String
'       IToolHandler_GetName = "my_tool_name"
'   End Function
'   Private Function IToolHandler_Execute(ByVal argsJson As String) As String
'       ' parse argsJson, do work, return result JSON string
'       IToolHandler_Execute = "{""result"":""done""}"
'   End Function

Public Function GetName() As String
    ' Return the OpenAI function name (e.g. "get_current_weather")
End Function

Public Function Execute(ByVal argsJson As String) As String
    ' Execute the tool with the given JSON arguments string.
    ' Return the result as a JSON string (can be a simple value or object).
End Function

Public Function GetDescription() As String
    ' Return a human-readable description of what this tool does.
    ' Used for ToolRegistry.BuildToolsJson() to construct the OpenAI tools array.
End Function

Public Function GetParametersSchemaJson() As String
    ' Return the JSON Schema for the tool's parameters, as a JSON string.
    ' Must be the value of the "parameters" field in the function definition:
    '   {"type":"object","properties":{...},"required":[...],"additionalProperties":false}
End Function

Public Property Get Strict() As Boolean
    ' Whether strict mode is enabled for this tool's function definition.
    ' Default True for OpenAI structured outputs.
End Property
