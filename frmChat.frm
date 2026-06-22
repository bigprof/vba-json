VERSION 5.00
Begin VB.Form frmChat 
   Caption         =   "OpenAI Chat"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrWatchdog 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   120
      Top             =   5880
   End
   Begin VB.Timer tmrLoop 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   5400
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Default         =   -1  'True
      Height          =   375
      Left            =   7680
      TabIndex        =   2
      Top             =   6120
      Width           =   1215
   End
   Begin VB.TextBox txtInput 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   6120
      Width           =   7455
   End
   Begin VB.TextBox txtChat 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   360
      Width           =   8775
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "Ready"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   8775
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================================
' frmChat — Chat Window UI
' ============================================================================
' Displays conversation history in a read-only text box, accepts user input,
' and drives a ChatEngine through a Timer-based loop.
'
' Usage:
'   Dim frm As New frmChat
'   frm.Show vbModal
'
' Before showing, configure the engine via SetChatEngine().

Private WithEvents mEngine As ChatEngine
Attribute mEngine.VB_VarHelpID = -1

Private mBusy As Boolean

Private Sub Form_Load()
    ' Form is loaded — child controls exist.
    ' Caller must call SetChatEngine() before the form is shown,
    ' or we initialize with defaults here.
    
    If mEngine Is Nothing Then
        ' Default setup: create engine with environment API key
        Dim ai As OpenAI
        Dim reg As ToolRegistry
        
        Set ai = New OpenAI
        ai.ApiKey = Environ$("OPENAI_API_KEY")
        
        Set reg = New ToolRegistry
        RegisterExampleTools reg
        
        Set mEngine = New ChatEngine
        Set mEngine.OpenAI = ai
        Set mEngine.ToolRegistry = reg
        mEngine.SystemPrompt = "You are a helpful assistant. Use tools when needed."
        mEngine.UseAsync = True
    End If
    
    mBusy = False
    AppendToChat "Welcome! Type a message and click Send.", ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Clean up async resources
    If Not mEngine Is Nothing Then
        If Not mEngine.OpenAI Is Nothing Then
            ' Engine owns the OpenAI client — nothing to dispose
        End If
    End If
    
    ' Disable timers
    tmrLoop.Enabled = False
    tmrWatchdog.Enabled = False
End Sub

Public Sub SetChatEngine(ByVal engine As ChatEngine)
    Set mEngine = engine
End Sub

' ============================================================================
' UI Event Handlers
' ============================================================================

Private Sub cmdSend_Click()
    SendMessage
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ' Enter key
        KeyAscii = 0
        SendMessage
    End If
End Sub

Private Sub SendMessage()
    Dim userText As String
    
    If mBusy Then Exit Sub
    
    userText = Trim$(txtInput.Text)
    If LenB(userText) = 0 Then Exit Sub
    
    ' Display user message
    AppendToChat userText, "User"
    
    ' Clear input
    txtInput.Text = ""
    txtInput.Enabled = False
    cmdSend.Enabled = False
    
    ' Start the engine
    mBusy = True
    
    On Error GoTo StartErr
    mEngine.StartConversation userText
    
    ' Enable timers
    tmrLoop.Enabled = True
    tmrWatchdog.Enabled = True
    
    Exit Sub

StartErr:
    AppendToChat "Error: " & Err.Description, "Error"
    mBusy = False
    txtInput.Enabled = True
    cmdSend.Enabled = True
End Sub

' ============================================================================
' Timer Handlers
' ============================================================================

Private Sub tmrLoop_Timer()
    ' Main loop driver — calls ContinueLoop() to advance the state machine
    
    If mEngine Is Nothing Then Exit Sub
    
    On Error Resume Next
    mEngine.ContinueLoop
    On Error GoTo 0
    
    ' If the engine is done, stop the loop
    If mEngine.IsDone Then
        tmrLoop.Enabled = False
        tmrWatchdog.Enabled = False
        mBusy = False
        txtInput.Enabled = True
        cmdSend.Enabled = True
        txtInput.SetFocus
    End If
End Sub

Private Sub tmrWatchdog_Timer()
    ' Watchdog keeps the UI from appearing hung during long API calls.
    ' Just pumps pending messages.
    lblStatus.Refresh
    DoEvents
End Sub

' ============================================================================
' ChatEngine Event Handlers
' ============================================================================

Private Sub mEngine_StatusChanged(ByVal Status As String)
    lblStatus.Caption = Status
    lblStatus.Refresh
    DoEvents
End Sub

Private Sub mEngine_MessageReceived(ByVal Text As String)
    AppendToChat Text, "Assistant"
End Sub

Private Sub mEngine_ToolExecuted(ByVal Name As String, ByVal Args As String, ByVal Result As String)
    Dim display As String
    
    display = "Tool: " & Name & vbCrLf & _
              "  Args: " & Args & vbCrLf & _
              "  Result: " & Result
    
    AppendToChat display, "Tool"
End Sub

Private Sub mEngine_ErrorOccurred(ByVal Msg As String)
    AppendToChat "ERROR: " & Msg, "Error"
    
    ' Stop everything on error
    tmrLoop.Enabled = False
    tmrWatchdog.Enabled = False
    mBusy = False
    txtInput.Enabled = True
    cmdSend.Enabled = True
    lblStatus.Caption = "Error occurred"
End Sub

Private Sub mEngine_Done()
    lblStatus.Caption = "Ready"
End Sub

' ============================================================================
' Helpers
' ============================================================================

Private Sub AppendToChat(ByVal text As String, ByVal role As String)
    Dim prefix As String
    Dim separator As String
    
    ' Build prefix based on role
    Select Case LCase$(role)
        Case "user"
            prefix = vbCrLf & "You: "
        Case "assistant", "assistant text"
            prefix = vbCrLf & "AI: "
        Case "tool"
            prefix = vbCrLf & "[Tool] " & vbCrLf
        Case "error"
            prefix = vbCrLf & "[ERROR] "
        Case Else
            prefix = vbCrLf
    End Select
    
    ' Add separator if there's already content
    If LenB(txtChat.Text) <> 0 Then
        separator = vbCrLf & String$(60, "-") & vbCrLf
    Else
        separator = ""
    End If
    
    txtChat.Text = txtChat.Text & separator & prefix & text
    
    ' Auto-scroll to end
    txtChat.SelStart = Len(txtChat.Text)
End Sub

' ============================================================================
' Example Tool Registration (default setup)
' ============================================================================

Private Sub RegisterExampleTools(ByVal reg As ToolRegistry)
    ' Register built-in example tools so the form works out of the box.
    ' Real applications should call SetChatEngine() with their own tools.
    
    Dim weatherTool As WeatherToolHandler
    Dim coordTool As CoordinatesToolHandler
    
    Set weatherTool = New WeatherToolHandler
    Set coordTool = New CoordinatesToolHandler
    
    reg.Register weatherTool
    reg.Register coordTool
End Sub
