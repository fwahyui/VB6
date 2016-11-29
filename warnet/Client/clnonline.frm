VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form4 
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   8100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7890
   Icon            =   "clnonline.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "clnonline.frx":FA8A
   ScaleHeight     =   8100
   ScaleWidth      =   7890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6600
      Top             =   480
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   6600
      Top             =   1440
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   6600
      Top             =   960
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Caption         =   "SETTING"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3855
      Left            =   120
      TabIndex        =   19
      Top             =   4080
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CommandButton Cmsavepasswd 
         Caption         =   "Save New Pswd"
         Height          =   375
         Left            =   2760
         Picture         =   "clnonline.frx":1BFE8
         TabIndex        =   41
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2760
         Locked          =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   45
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton cmDecrypt 
         Caption         =   "Decrypt"
         CausesValidation=   0   'False
         Height          =   375
         Left            =   3600
         TabIndex        =   44
         Top             =   1680
         Width           =   735
      End
      Begin VB.CommandButton cmEncrypt 
         Caption         =   "Encrypt"
         Height          =   375
         Left            =   2760
         TabIndex        =   43
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2760
         Locked          =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   42
         Text            =   "no data input"
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2760
         TabIndex        =   40
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00800000&
         Caption         =   "AUTO OFF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   2880
         TabIndex        =   34
         Top             =   2640
         Width           =   1575
         Begin VB.TextBox Txidle 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            TabIndex        =   35
            Text            =   "3600"
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Idle time/ sec. :"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.CommandButton Settingstop 
         Caption         =   "X"
         Height          =   375
         Left            =   4440
         Picture         =   "clnonline.frx":1E558
         TabIndex        =   30
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton Settingsave 
         Caption         =   "Save I.P./Port"
         Height          =   375
         Left            =   1080
         Picture         =   "clnonline.frx":20AC8
         TabIndex        =   28
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox st4 
         Height          =   285
         Left            =   1080
         TabIndex        =   26
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox st3 
         Height          =   285
         Left            =   1080
         TabIndex        =   24
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox st2 
         Height          =   285
         Left            =   1080
         TabIndex        =   22
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox st1 
         Height          =   285
         Left            =   1080
         TabIndex        =   20
         Top             =   360
         Width           =   1455
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   4
         Height          =   3615
         Left            =   120
         Top             =   120
         Width           =   4575
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "New Password"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   46
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   360
         Left            =   240
         Picture         =   "clnonline.frx":23038
         Top             =   3240
         Width           =   360
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "CLIENTS SETTING"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   720
         TabIndex        =   29
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Picpath :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "User      :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Port       :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "I.P.        :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3120
      ScaleHeight     =   345
      ScaleWidth      =   3225
      TabIndex        =   13
      Top             =   3600
      Width           =   3255
      Begin VB.CommandButton Shrtrestart 
         Caption         =   "Restart"
         Height          =   375
         Left            =   2160
         TabIndex        =   15
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton Shrtshutdown 
         Caption         =   "Shutdown"
         Height          =   375
         Left            =   1080
         TabIndex        =   14
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton Shrtfreenter 
         Caption         =   "Admin"
         Height          =   375
         Left            =   0
         TabIndex        =   37
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   4785
      TabIndex        =   6
      Top             =   0
      Width           =   4815
      Begin VB.CommandButton CmChat 
         Appearance      =   0  'Flat
         Caption         =   "CHAT"
         Height          =   255
         Left            =   4080
         TabIndex        =   17
         Top             =   50
         Width           =   615
      End
      Begin VB.CommandButton CmViewer 
         Caption         =   "STOP"
         Height          =   255
         Left            =   3360
         TabIndex        =   9
         ToolTipText     =   "Stop internet connection"
         Top             =   50
         Width           =   615
      End
      Begin VB.Label TXPayment 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   50
         Width           =   2775
      End
      Begin VB.Image Image4 
         Height          =   240
         Left            =   120
         Picture         =   "clnonline.frx":234F0
         Top             =   45
         Width           =   240
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   3015
      Left            =   120
      ScaleHeight     =   2955
      ScaleWidth      =   6195
      TabIndex        =   0
      Top             =   480
      Width           =   6255
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "clnonline.frx":238EE
         Left            =   2880
         List            =   "clnonline.frx":238FB
         TabIndex        =   39
         Text            =   "Picture-1"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox animasi 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4320
         TabIndex        =   31
         Text            =   "0"
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton Cmsetting 
         Caption         =   "Setting"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2160
         TabIndex        =   18
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton CmAccess 
         Caption         =   "START INTERNET"
         Height          =   735
         Left            =   240
         TabIndex        =   5
         ToolTipText     =   "Click here to start browsing internet"
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CommandButton CmCaptured 
         Caption         =   "Capture"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         TabIndex        =   12
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton CMExit 
         Caption         =   "Exit"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5040
         TabIndex        =   4
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton CmConnect 
         Caption         =   "Connect"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4080
         Picture         =   "clnonline.frx":23920
         TabIndex        =   2
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox Txpass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   4800
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Auto Turnoff"
         Height          =   255
         Left            =   4800
         TabIndex        =   38
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Idle :"
         Height          =   255
         Left            =   4800
         TabIndex        =   33
         Top             =   480
         Width           =   375
      End
      Begin VB.Label lblidle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00"
         Height          =   255
         Left            =   5160
         TabIndex        =   32
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   3000
         TabIndex        =   16
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright (C) 2007 M-Technology Bali Indonesia"
         Height          =   255
         Left            =   2520
         TabIndex        =   11
         Top             =   2640
         Width           =   3495
      End
      Begin VB.Label TxStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "Disconnected"
         Height          =   255
         Left            =   2880
         TabIndex        =   3
         Top             =   480
         Width           =   1095
      End
      Begin VB.Image Image5 
         Height          =   525
         Left            =   2400
         Picture         =   "clnonline.frx":25E90
         Top             =   360
         Width           =   525
      End
      Begin VB.Image Image2 
         Height          =   360
         Left            =   2520
         Picture         =   "clnonline.frx":263CD
         Top             =   1200
         Width           =   360
      End
      Begin VB.Label TXsays 
         BackStyle       =   0  'Transparent
         Caption         =   "TxtSays"
         Height          =   255
         Left            =   3000
         TabIndex        =   7
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Image Image3 
         Height          =   330
         Left            =   3960
         Picture         =   "clnonline.frx":26885
         Top             =   405
         Width           =   270
      End
      Begin VB.Image Image6 
         Height          =   1830
         Left            =   120
         Picture         =   "clnonline.frx":26B1A
         Top             =   120
         Width           =   6000
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Your name :"
         Height          =   255
         Left            =   3720
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
   End
   Begin MSWinsockLib.Winsock sock1 
      Left            =   6600
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
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

Option Explicit
Private Declare Function Taskbar_Show_Hide Lib "InjectApp.dll" (ByVal bShowHide As Boolean) As Integer
Private Declare Function AltTab2_Enable_Disable Lib "InjectApp.dll" (ByVal hwnd As Long, ByVal bEnableDisable As Boolean) As Integer
Private Declare Function TaskSwitching_Enable_Disable Lib "InjectApp.dll" (ByVal bEnableDisable As Boolean) As Integer
Private Declare Function TaskManager_Enable_Disable Lib "InjectApp.dll" (ByVal bEnableDisable As Boolean) As Integer
Private Declare Function CtrlAltDel_Enable_Disable Lib "InjectApp.dll" (ByVal bEnableDisable As Boolean) As Integer
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Dim IPVALUE As String
Dim PORTVALUE As String
Dim USERVALUE As String
Dim Fileku1 As String
Dim Fileku2 As String
Dim Picpath As String
Dim cExitWindows As New clsExitWindows

Private Function SecondsToTime(ByVal dSeconds As Double) As String
    SecondsToTime = Format(DateAdd("s", dSeconds, "00:00:00"), "HH:mm:ss")
End Function

Public Function Capture_Desktop(ByVal Destination$) As Boolean
On Error GoTo errl
DoEvents
Call keybd_event(vbKeySnapshot, 1, 0, 0)
DoEvents
SavePicture Clipboard.GetData(vbCFBitmap), Destination$
Capture_Desktop = True
Exit Function
errl:
MsgBox "Error number: " & Err.Number & ". " & Err.Description
Capture_Desktop = False
End Function

Private Sub CmAccess_Click()
If sock1.State = sckClosed Then
    CmConnect_Click
    Exit Sub
    End If
Picture2.Visible = True
Timer1.Enabled = True
Form4.Width = 4800
Form4.Height = 370
Timer3.Enabled = False: Txidle.Text = 7200
'-----
On Error GoTo t
sock1.SendData USERVALUE
Exit Sub
t:
TXsays.Caption = "Error : " & Err.Description
sock1_Close
End Sub

Private Sub CmCaptured_Click()
On Error GoTo t
Fileku1 = App.Path + "\" & USERVALUE & "desktop.bmp"
Fileku2 = Picpath + USERVALUE & "desktop.jpg"
Capture_Desktop Fileku1
FileCopy Fileku1, Fileku2
Exit Sub
t:
Label2.Caption = "Error:" & Error
End Sub

Private Sub CmChat_Click()
Form1.Show
End Sub

Private Sub CmConnect_Click()
sock1.Close
sock1.RemoteHost = IPVALUE
sock1.RemotePort = PORTVALUE
sock1.Connect
End Sub

Private Sub cmDecrypt_Click()
Text2.Text = decrypted(3, Text1.Text)
If Text2.Text <> "" Then
Text1.Text = Text2.Text
End If
End Sub

Private Sub cmEncrypt_Click()
Text2.Text = encrypted(3, Text1.Text)
If Text2.Text <> "" Then
Text1.Text = Text2.Text
End If
End Sub

Private Sub CMExit_Click()
WINLOCKOPEN
End
End Sub

Private Sub BERHENTI()
Picture2.Visible = False
If Timer1.Enabled = False Then Exit Sub
Timer1.Enabled = False
Timer3.Enabled = True
On Error GoTo t
sock1.SendData USERVALUE & "-STOP"
Exit Sub
t:
TXsays.Caption = "Error : " & Err.Description
sock1_Close
End Sub


Private Sub Cmsavepasswd_Click()
Text1.Text = Text3.Text
cmEncrypt_Click
Dim intFileHandle As Integer
    intFileHandle = FreeFile
    Open App.Path + "\mwarnet.pwd" For Output As #intFileHandle
    Write #intFileHandle, Text1.Text
    Close #intFileHandle
End Sub

Private Sub Cmsetting_Click()
Frame1.Visible = True
Timer3.Enabled = False
End Sub

Private Sub CmViewer_Click()
LAYAR
BERHENTI
End Sub

Private Sub Combo1_Click()
If Combo1.Text = "Picture-1" Then Form4.Picture = LoadPicture("pics01.jpg")
If Combo1.Text = "Picture-2" Then Form4.Picture = LoadPicture("pics02.jpg")
If Combo1.Text = "Picture-3" Then Form4.Picture = LoadPicture("pics03.jpg")
End Sub

Private Sub Form_Load()
If App.PrevInstance Then Unload Me: End
Picture2.Visible = False
LAYAR
CLIENTDATA
TranslucentForm Me, 240
SetWindowOnTop Me, True
TXsays.Caption = USERVALUE & ": " & PORTVALUE & " / " & IPVALUE
CmConnect_Click
WINLOCKCLOSE
PASSVALUES
End Sub

Public Sub PASSVALUES()
On Error GoTo t
Dim intFileHandle As Integer
Dim pssd As String
intFileHandle = FreeFile
Open App.Path + "\mwarnet.pwd" For Input As #1
Input #intFileHandle, pssd
Text1.Text = pssd
cmDecrypt_Click
Close #intFileHandle
Exit Sub
t:
End Sub

Private Sub LAYAR()
With Me
    .Width = Screen.Width
    .Height = Screen.Height
    .Top = 0
    .Left = 0
    End With
With Picture1
    .Move (Screen.Width - .Width) \ 2, (Screen.Height - .Height) \ 2
    End With
With Picture3
    .Move (0), (Screen.Height - .Height)
    End With
End Sub

Private Sub Settingsave_Click()
Dim intFileHandle As Integer
intFileHandle = FreeFile
Open App.Path + "\clientsdata.txt" For Output As #intFileHandle
    Print #intFileHandle, st1.Text
    Print #intFileHandle, st2.Text
    Print #intFileHandle, st3.Text
    Print #intFileHandle, st4.Text
    Close #intFileHandle
WINLOCKOPEN
End
End Sub

Private Sub Settingstop_Click()
Frame1.Visible = False
Timer3.Enabled = True
End Sub

Private Sub Shrtfreenter_Click()
If Txpass.Visible = False Then
    Txpass.Visible = True
    Else
    Txpass.Visible = False
    End If
End Sub

Private Sub Shrtrestart_Click()
cExitWindows.ExitWindows WE_REBOOT
End Sub

Private Sub Shrtshutdown_Click()
cExitWindows.ExitWindows WE_SHUTDOWN
End Sub

Private Sub sock1_Close()
sock1.Close
TXsays.Caption = "Disconnected from server" & vbCrLf
TxStatus.Caption = "Disconnected"
End Sub
 
Private Sub sock1_Connect()
TXsays.Caption = "Connected to " & sock1.RemoteHostIP & vbCrLf
TxStatus.Caption = "Connected"
CmAccess.Enabled = True
sock1.SendData USERVALUE & "-ON"
End Sub

Private Sub sock1_DataArrival(ByVal bytesTotal As Long)
Dim dat As String
sock1.GetData dat, vbString
TXPayment.Caption = dat
If dat = "SHUTDOWN" Then
    cExitWindows.ExitWindows WE_SHUTDOWN
    End If
If dat = "REBOOT" Then
    cExitWindows.ExitWindows WE_REBOOT
    End If
If dat = "LOGOFF" Then
    cExitWindows.ExitWindows WE_LOGOFF
    End If
If dat = "CAPTURED" Then
    CmCaptured_Click
    End If
If dat = "STOP" Then
    CmViewer_Click
    CmAccess.Enabled = True
    End If
If dat = "CHAT" Then
    CmChat_Click
    End If
End Sub

Private Sub sock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
TXsays.Caption = "*** Error : " & Description & vbCrLf
sock1_Close
End Sub

Private Sub CLIENTDATA()
Dim intFileHandle As Integer
Dim strRETP As String
intFileHandle = FreeFile
Open App.Path + "\clientsdata.txt" For Input As #intFileHandle
Line Input #intFileHandle, strRETP: IPVALUE = strRETP: st1.Text = strRETP
Line Input #intFileHandle, strRETP: PORTVALUE = strRETP: st2.Text = strRETP
Line Input #intFileHandle, strRETP: USERVALUE = strRETP: st3.Text = strRETP
Line Input #intFileHandle, strRETP: Picpath = strRETP: st4.Text = strRETP
Close #intFileHandle
End Sub


Private Sub Timer1_Timer()
CmAccess_Click
End Sub

Private Sub WINLOCKOPEN()
    AltTab2_Enable_Disable 0, True
    TaskSwitching_Enable_Disable (True)
    TaskManager_Enable_Disable (True)
    CtrlAltDel_Enable_Disable (True)
End Sub

Private Sub WINLOCKCLOSE()
    AltTab2_Enable_Disable 0, False
    TaskSwitching_Enable_Disable (False)
    TaskManager_Enable_Disable (False)
    CtrlAltDel_Enable_Disable (False)
End Sub

Private Sub Timer2_Timer()
If TxStatus.Caption = "Connected" Then
    animasi.Text = Val(animasi.Text) + 1
    If animasi.Text = 1 Then Image3.Visible = False: Image4.Visible = False
    If animasi.Text = 2 Then Image3.Visible = True: Image4.Visible = True: animasi.Text = 0
    Else
    CmAccess.Enabled = False
    Picture2.Visible = False
    If Form4.Width < 4100 Then
        LAYAR
        End If
    CmConnect_Click
    End If
End Sub

Private Sub Timer3_Timer()
Txidle.Text = Val(Txidle.Text - 1)
lblidle.Caption = SecondsToTime(Txidle.Text)
If Txidle.Text = 0 Then
    Timer3.Enabled = False
    Shrtshutdown_Click
    End If
End Sub

Private Sub Txidle_KeyPress(KeyAscii As Integer)
    Dim Counter As Byte
    Dim IsNumber As Boolean
    IsNumber = False
    For Counter = 0 To 9
        If Chr$(KeyAscii) = Trim(Counter) Then IsNumber = True
    Next
    If Not IsNumber Then Txidle.Text = ""
End Sub

Private Sub Txpass_Change()
If Txpass.Text = Text1.Text Then
    CMExit.Enabled = True
    CmConnect.Enabled = True
    CmCaptured.Enabled = True
    Cmsetting.Enabled = True
    WINLOCKOPEN
    End If
End Sub
