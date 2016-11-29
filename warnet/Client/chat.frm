VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MWARNET - CHAT"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7515
   Icon            =   "chat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "chat.frx":FA8A
   ScaleHeight     =   5130
   ScaleWidth      =   7515
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   375
      Left            =   6000
      TabIndex        =   6
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   615
      Left            =   6000
      TabIndex        =   5
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox txtConnect 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4080
      TabIndex        =   3
      Text            =   "192.168.1.2"
      Top             =   4320
      Width           =   1815
   End
   Begin VB.TextBox txtChat 
      Height          =   2295
      Left            =   600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1200
      Width           =   6375
   End
   Begin VB.TextBox txtSend 
      Height          =   645
      Left            =   600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   3600
      Width           =   5295
   End
   Begin MSWinsockLib.Winsock w1 
      Left            =   2370
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblCredits 
      BackStyle       =   0  'Transparent
      Caption         =   "SERVERS [13132 user(s) online]"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   4440
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   3360
      Picture         =   "chat.frx":17302
      Top             =   480
      Width           =   690
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CHATTING !"
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
      Left            =   4200
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
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

Private Sub cmdConnect_Click()
    If Len(txtConnect.Text) = 0 Then
    Else
        w1.Close
        w1.Connect txtConnect.Text, 1
    End If
End Sub

Private Sub cmdSend_Click()
    If Len(txtSend.Text) = 0 Then
    Else
        w1.SendData txtSend.Text
        txtSend.Text = ""
    End If
End Sub

Private Sub Form_Load()
    Status 2
    cmdSend.Default = True
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
'Try to connect
txtConnect.Text = Form4.sock1.RemoteHostIP
If Len(txtConnect.Text) = 0 Then
    Else
    w1.Close
    w1.Connect txtConnect.Text, 1
    End If
End Sub

Private Sub Status(ByVal WhichStatus As Integer)
    Select Case WhichStatus
        Case 1 'Connected
            txtChat.Enabled = True
            txtSend.Enabled = True
            cmdSend.Enabled = True
            lblCredits.Caption = "SERVERS [" & OnlineCount & " user(s) online]"
            cmdConnect.Enabled = False
            txtConnect.Enabled = False
        Case 2 'Disconnected
            txtChat.Enabled = False
            txtSend.Enabled = False
            cmdSend.Enabled = False
            lblCredits.Caption = "MWARNET 2"
            cmdConnect.Enabled = True
            txtConnect.Enabled = True
            w1.Close
    End Select
End Sub


Private Sub w1_Close()
    Status 2
End Sub

Private Sub w1_Connect()
    Status 1
End Sub

Private Sub w1_DataArrival(ByVal bytesTotal As Long)
    Dim data As String
    w1.GetData data, vbString
    If Mid(data, 1, 6) = "Online" Then
        OnlineCount = Mid(data, 7, Len(data) - 6)
        lblCredits.Caption = "Servers [" & OnlineCount & " user(s) online]"
    Else
        txtChat.Text = txtChat.Text & data & vbCrLf
    End If
End Sub

Private Sub w1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "An error has ocurred and te connection will therefore be broken." & vbCrLf & vbCrLf & "Error description : " & Description, vbOKOnly + vbCritical, "BartNet Chat 3.0"
    w1.Close
End Sub
