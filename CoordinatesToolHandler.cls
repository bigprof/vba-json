VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CoordinatesToolHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IToolHandler

' ============================================================================
' CoordinatesToolHandler — Example IToolHandler for get_location_coordinates
' ============================================================================
' Returns mock coordinates data. Replace with a real geocoding API in production.

Private Function IToolHandler_GetName() As String
    IToolHandler_GetName = "get_location_coordinates"
End Function

Private Function IToolHandler_Execute(ByVal argsJson As String) As String
    ' In production, parse argsJson, call a geocoding API, return real data.
    ' For now, return mock data.
    IToolHandler_Execute = "{""location"":""Boston, MA"",""latitude"":42.3601,""longitude"":-71.0589}"
End Function

Private Function IToolHandler_GetDescription() As String
    IToolHandler_GetDescription = "Get geographic coordinates for a location"
End Function

Private Function IToolHandler_GetParametersSchemaJson() As String
    IToolHandler_GetParametersSchemaJson = _
        "{" & _
            """type"":""object""," & _
            """properties"":{" & _
                """location"":{" & _
                    """type"":""string""," & _
                    """description"":""The city and state, e.g. San Francisco, CA""" & _
                "}" & _
            "}," & _
            """required"":[""location""]," & _
            """additionalProperties"":false" & _
        "}"
End Function

Private Property Get IToolHandler_Strict() As Boolean
    IToolHandler_Strict = True
End Property
