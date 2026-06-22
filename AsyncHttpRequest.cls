VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AsyncHttpRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================================================
' AsyncHttpRequest — Non-blocking HTTP wrapper using WithEvents on XMLHTTP60
' ============================================================================
' Fires Completed or Error events when the response arrives.
' The caller's message pump (Timer + DoEvents) keeps the UI alive while waiting.
'
' Usage:
'   Private WithEvents mAsync As AsyncHttpRequest
'   Set mAsync = New AsyncHttpRequest
'   mAsync.Start "POST", "https://api.openai.com/v1/chat/completions", body, headers
'   ' ... Timer pumps DoEvents ...
'   Private Sub mAsync_Completed(ByVal status As Long, ByVal responseText As String)
'       ' process response
'   End Sub

Public Event Completed(ByVal Status As Long, ByVal ResponseText As String)
Public Event ErrorOccurred(ByVal Number As Long, ByVal Description As String)

Private WithEvents mHttp As MSXML2.XMLHTTP60
Attribute mHttp.VB_VarHelpID = -1

Private mIsBusy As Boolean
Private mRequestHeaders As Object ' Scripting.Dictionary of header name -> value

Private Sub Class_Initialize()
    Set mHttp = New MSXML2.XMLHTTP60
    Set mRequestHeaders = CreateObject("Scripting.Dictionary")
    mIsBusy = False
End Sub

Private Sub Class_Terminate()
    If Not mHttp Is Nothing Then
        On Error Resume Next
        If mIsBusy Then
            mHttp.abort
        End If
        Set mHttp = Nothing
    End If
    If Not mRequestHeaders Is Nothing Then
        mRequestHeaders.RemoveAll
        Set mRequestHeaders = Nothing
    End If
End Sub

Public Sub Start(ByVal method As String, ByVal url As String, ByVal body As String, Optional ByVal Headers As Variant = Empty)
    Dim key As Variant
    
    If mIsBusy Then
        Err.Raise vbObjectError + 4001, "AsyncHttpRequest.Start", "A request is already in progress. Call Cancel() first."
    End If
    
    mRequestHeaders.RemoveAll
    
    ' Default Content-Type
    mRequestHeaders.Add "Content-Type", "application/json"
    
    ' Copy caller-provided headers (Collection or Dictionary)
    If Not IsEmpty(Headers) Then
        If Not IsNull(Headers) Then
            If TypeOf Headers Is Collection Then
                ' Collection of 2-element arrays or key-value pairs
                Dim col As Collection
                Dim item As Variant
                Dim i As Long
                Set col = Headers
                For i = 1 To col.Count
                    Set item = col.Item(i)
                    If IsArray(item) Then
                        mRequestHeaders.Add CStr(item(0)), CStr(item(1))
                    End If
                Next i
            ElseIf TypeOf Headers Is Object Then
                ' Assume Scripting.Dictionary
                Dim dict As Object
                Set dict = Headers
                For Each key In dict.Keys
                    mRequestHeaders.Add CStr(key), CStr(dict.Item(key))
                Next key
            End If
        End If
    End If
    
    On Error GoTo TransportErr
    
    mHttp.Open method, url, True ' Async!
    
    ' Set all headers
    For Each key In mRequestHeaders.Keys
        mHttp.setRequestHeader CStr(key), CStr(mRequestHeaders.Item(key))
    Next key
    
    mIsBusy = True
    
    If LenB(body) = 0 Then
        mHttp.send
    Else
        mHttp.send body
    End If
    
    Exit Sub

TransportErr:
    mIsBusy = False
    RaiseEvent ErrorOccurred(Err.Number, "Transport error: " & Err.Description & vbCrLf & "URL: " & url)
End Sub

Public Sub Cancel()
    If mIsBusy Then
        On Error Resume Next
        mHttp.abort
        On Error GoTo 0
        mIsBusy = False
    End If
End Sub

Public Property Get IsBusy() As Boolean
    IsBusy = mIsBusy
End Property

Public Property Get StatusCode() As Long
    On Error Resume Next
    StatusCode = mHttp.Status
    On Error GoTo 0
End Property

Private Sub mHttp_onReadyStateChange()
    Dim status As Long
    Dim responseText As String
    
    If mHttp.readyState <> 4 Then Exit Sub ' Not done yet
    
    mIsBusy = False
    status = mHttp.Status
    responseText = mHttp.responseText
    
    RaiseEvent Completed(status, responseText)
End Sub
