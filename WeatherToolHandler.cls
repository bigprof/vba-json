VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WeatherToolHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IToolHandler

' ============================================================================
' WeatherToolHandler — Example IToolHandler for get_current_weather
' ============================================================================
' Returns mock weather data. Replace with a real API call in production.

Private Function IToolHandler_GetName() As String
    IToolHandler_GetName = "get_current_weather"
End Function

Private Function IToolHandler_Execute(ByVal argsJson As String) As String
    ' In production, parse argsJson, call a weather API, return real data.
    ' For now, return mock data.
    IToolHandler_Execute = "{""location"":""Boston, MA"",""temperature"":72,""unit"":""fahrenheit"",""description"":""Partly cloudy""}"
End Function

Private Function IToolHandler_GetDescription() As String
    IToolHandler_GetDescription = "Get the current weather in a given location"
End Function

Private Function IToolHandler_GetParametersSchemaJson() As String
    IToolHandler_GetParametersSchemaJson = _
        "{" & _
            """type"":""object""," & _
            """properties"":{" & _
                """location"":{" & _
                    """type"":""string""," & _
                    """description"":""The city and state, e.g. San Francisco, CA""" & _
                "}," & _
                """unit"":{" & _
                    """type"":""string""," & _
                    """enum"":[""celsius"",""fahrenheit""]," & _
                    """description"":""Temperature unit""" & _
                "}" & _
            "}," & _
            """required"":[""location"",""unit""]," & _
            """additionalProperties"":false" & _
        "}"
End Function

Private Property Get IToolHandler_Strict() As Boolean
    IToolHandler_Strict = True
End Property
