VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MWARNET - CHAT"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7515
   Icon            =   "chat.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "chat.frx":FA8A
   ScaleHeight     =   5115
   ScaleWidth      =   7515
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock wChat 
      Index           =   0
      Left            =   6000
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wListen 
      Left            =   5400
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   615
      Left            =   6000
      TabIndex        =   4
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox txtChat 
      Height          =   2295
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "chat.frx":17302
      Top             =   1200
      Width           =   6495
   End
   Begin VB.TextBox txtSend 
      Height          =   645
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   3600
      Width           =   5415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0"
      Height          =   255
      Left            =   4800
      TabIndex        =   6
      Top             =   740
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CHATROOM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   5
      Top             =   480
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   3360
      Picture         =   "chat.frx":17333
      Top             =   480
      Width           =   690
   End
   Begin VB.Label lblCredits 
      BackStyle       =   0  'Transparent
      Caption         =   "SERVERS [13132 user(s) online]"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   4320
      Width           =   3735
   End
   Begin VB.Label lblIP 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "123.123.123.123"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   4320
      Width           =   2295
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MWARNET 2 - FREEWARE EDITION
'COPYRIGHT(C) 2007 MTechnologi Bali Indonesia
'Programed by A.A.Ngr.Manik Artawan
'e-mail : gungmanik@telkom.net
'---------------------------------------------
'THANK YOu FOR DOWNLOAD THIS SMALL APPLICATION
'---------------------------------------------

Private OnlineCount As Integer
Private Sub cmdSend_Click()
    Dim a As Integer
    Dim Message As String
    Message = "Server says: " & txtSend.Text
    Do Until a = wChat.Count
        If wChat(a).State = sckClosed Then
        Else
            wChat(a).SendData Message
            DoEvents
        End If
        a = a + 1
    Loop
    txtChat.Text = txtChat.Text & Message & vbCrLf
    txtSend.Text = ""
End Sub

Private Sub Form_Load()
    wListen.LocalPort = 1
    wListen.Listen
    wChat(0).LocalPort = 2
    OnlineCount = 0
    cmdSend.Default = True
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    lblCredits.Caption = "SERVER : [" & OnlineCount & " user(s) online]"
    lblIP.Caption = "Your IP = " & wListen.LocalIP
End Sub


Private Sub wChat_Close(Index As Integer)
    wChat(Index).Close
    UpdateUserCount
End Sub

Private Sub wChat_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim a As Integer
    Dim Message As String
    Dim data As String
    wChat(Index).GetData data, vbString
    Message = "Client " & Index & " says: " & data
    Do Until a = wChat.Count
        If wChat(a).State = sckClosed Then
        Else
            wChat(a).SendData Message
            DoEvents
        End If
        a = a + 1
    Loop
    
    txtChat.Text = txtChat.Text & Message & vbCrLf
End Sub


Private Sub wChat_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "An error has ocurred and te connection will therefore be broken." & vbCrLf & vbCrLf & "Error description : " & Description, vbOKOnly + vbCritical, "BartNet Chat 3.0"
    CloseAll
End Sub

Private Sub wListen_ConnectionRequest(ByVal requestID As Long)
    Dim a As Integer
    a = GetNextWinsock
    wChat(a).Close
    wChat(a).Accept requestID
    UpdateUserCount
End Sub

Private Function GetNextWinsock()
    Dim a As Integer
    a = wChat.Count
    Load wChat(a)
    wChat(a).LocalPort = a + 2
    GetNextWinsock = a
End Function

Private Sub UpdateUserCount()
    Dim a As Integer
    Dim Online As Integer
    Do Until a = wChat.Count
        If wChat(a).State = sckClosed Then
        Else
            Online = Online + 1
        End If
        a = a + 1
    Loop
    
    OnlineCount = Online
    lblCredits.Caption = "SERVER : [" & OnlineCount & " user(s) online]"
    a = 0
    Do Until a = wChat.Count
        If wChat(a).State = sckClosed Then
        Else
            wChat(a).SendData "Online" & OnlineCount
            DoEvents
        End If
        a = a + 1
    Loop
End Sub

Private Sub wListen_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "An error has ocurred and te connection will therefore be broken." & vbCrLf & vbCrLf & "Error description : " & Description, vbOKOnly + vbCritical, "BartNet Chat 3.0"
    CloseAll
End Sub

Private Sub CloseAll()
    wListen.Close
    Dim a As Integer
    Do Until a = wChat.Count
        wChat(a).Close
        a = a + 1
    Loop
    OnlineCount = 0
    lblCredits.Caption = "SERVER : [" & OnlineCount & " user(s) online]"
End Sub
