VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "FlDbg9c.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Client"
   ClientHeight    =   8670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   8670
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrConnectDelay 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   4320
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   360
      ScaleHeight     =   255
      ScaleWidth      =   1095
      TabIndex        =   12
      Top             =   840
      Width           =   1095
      Begin VB.Shape Light 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   0
         Shape           =   2  'Oval
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         MouseIcon       =   "frmMain.frx":08CA
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.Timer tmrDelayer 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   3240
      Top             =   0
   End
   Begin VB.Timer tmrConnector 
      Interval        =   3000
      Left            =   2280
      Top             =   0
   End
   Begin VB.PictureBox picAbout 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   4920
      ScaleHeight     =   3105
      ScaleWidth      =   5385
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3120
      Width           =   5415
      Begin VB.CommandButton cmdOK 
         Cancel          =   -1  'True
         Caption         =   "OK"
         Height          =   345
         Left            =   3765
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   2265
         Width           =   1260
      End
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
         Height          =   615
         Left            =   600
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1560
         Width           =   3855
         _cx             =   5080
         _cy             =   5080
         FlashVars       =   ""
         Movie           =   "c:\new cafe\sErVere\Swf\about.swf"
         Src             =   "c:\new cafe\sErVere\Swf\about.swf"
         WMode           =   "Transparent"
         Play            =   -1  'True
         Loop            =   -1  'True
         Quality         =   "High"
         SAlign          =   ""
         Menu            =   -1  'True
         Base            =   ""
         AllowScriptAccess=   ""
         Scale           =   "ExactFit"
         DeviceFont      =   0   'False
         EmbedMovie      =   -1  'True
         BGColor         =   ""
         SWRemote        =   ""
         MovieData       =   ""
         SeamlessTabbing =   -1  'True
         Profile         =   0   'False
         ProfileAddress  =   ""
         ProfilePort     =   0
         AllowNetworking =   "all"
         AllowFullScreen =   "false"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Programmed && Designed by:"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   1995
      End
      Begin VB.Image Image1 
         Height          =   960
         Left            =   120
         Picture         =   "frmMain.frx":0A1C
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label lblVersion 
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Version 1.2"
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Left            =   1440
         TabIndex        =   7
         Top             =   720
         Width           =   3885
      End
      Begin VB.Label lblTitle 
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "MASIFA Empire's Internet System          CLIENT APP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   600
         Left            =   1320
         TabIndex        =   6
         Top             =   120
         Width           =   4365
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "fitrianwahyuilahi45@gmail.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1440
         TabIndex        =   5
         Top             =   2760
         Width           =   3135
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2760
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   -100
      TabIndex        =   2
      Top             =   -100
      Width           =   15
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Command5"
         Default         =   -1  'True
         Height          =   495
         Left            =   0
         TabIndex        =   1
         Top             =   480
         Width           =   75
      End
   End
   Begin MSWinsockLib.Winsock Sucket 
      Left            =   1800
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   3982
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Shock 
      Height          =   7695
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   11535
      _cx             =   20346
      _cy             =   13573
      FlashVars       =   ""
      Movie           =   "E:\Beta\warnet\Internet Cafe System\sErVere\iron.swf"
      Src             =   "E:\Beta\warnet\Internet Cafe System\sErVere\iron.swf"
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "AutoHigh"
      SAlign          =   ""
      Menu            =   0   'False
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ExactFit"
      DeviceFont      =   0   'False
      EmbedMovie      =   -1  'True
      BGColor         =   "000000"
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin CLIENTE.SkinForm SkinnedForm 
      Height          =   855
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1508
      MinimizeBtn     =   0   'False
      Caption         =   "MASIFA Empire's Internet System Screen Lock"
      CaptionTop      =   250
   End
   Begin VB.PictureBox VolumeMaster 
      Height          =   480
      Left            =   3720
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   14
      Top             =   0
      Width           =   1200
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Setting"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      MouseIcon       =   "frmMain.frx":12E6
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  picAbout.Visible = False
  cmdOK.Default = False
  Command5.Default = True
End Sub

Private Sub Command5_Click()
Dim a As String
  a = DbRst!Pwd
  If Text1.Text = a Then
    frmConfig.Show 1
    Text1.Text = ""
  Else
    Text1.Text = ""
    Text1.SetFocus
  End If
End Sub

Private Sub Command5_GotFocus()
  Text1.SetFocus
End Sub

Private Sub Form_Activate()
  picAbout.Visible = False
  If Locked = True Then
  '  BlockCtrl_Alt_Del True
  Else
  '  BlockCtrl_Alt_Del False
  End If
End Sub

Private Sub Form_Load()
  Sucket.Close
  SkinnedForm.LoadSkin Me
  Set Shells = CreateObject("Shell.Application")

  'VolumeMaster.Volume = 100
  'VolumeMaster.Mute = False
  RemoveProgramFromList
  'BlockCtrl_Alt_Del True
  Sucket.Connect Trim(DbRst!SvrIP)
End Sub

Private Sub Form_Resize()
  SkinnedForm.LoadSkin Me
  Shock.Width = Me.Width - 500
  Shock.Height = Me.Height - 800
End Sub

Private Sub Label1_Click()
  picAbout.Visible = True
  Command5.Default = False
  cmdOK.Default = True
  Call Command5_Click
End Sub

Private Sub Label4_Click()
    Call Command5_Click
End Sub

Private Sub Shock_GotFocus()
  Text1.SetFocus
End Sub

Private Sub Sucket_Close()
  Sucket.Close
  If Mark = 1 Then
    frmPopUp.cmdLogout.Enabled = False
    frmTrayMenu.mnuLogOut.Enabled = False
    Mark = 0
  End If
  Sucket.Connect Trim(DbRst!SvrIP)
  Light.BackColor = vbRed
End Sub

Private Sub Sucket_Connect()
  Light.BackColor = vbGreen
  Sucket.SendData "ID" & DbRst!CompID
End Sub

Private Sub Sucket_DataArrival(ByVal bytesTotal As Long)
Dim X As String
  Sucket.GetData X
  X = Check_Data(X)
  
  'Language Headers
  'OP - UnlockComp w/Open acnt: addt'l data (start time)
  'LT - UnlockComp w/Ltd acnt:  addt'l data (3 chars of amt, start time)
  'ET -  ET1-Lock/ ET2-UnlockComp/ ET-Elapse Time
  'T1 - open account data
  'T2 - limited account data
  'SD - shutdown
  'SZ - Synchronize the flashplayer
  'MT - MT1-Mute Volume/ MT2-Unmute Volume
  'VL - Set Volume
  'ER - (Error) ER1-ID# isalready used/ ER2-Server is full
Alex:
  Select Case Left(X, 2)
  Case "OP" ' set open account
    'DATA FORMAT
    'OP + (Name-30Chr) + (ServiceType-8Chr) + (Rate-2Chr) + (LoginTime)
    BlockCtrl_Alt_Del False
    TypeTrans = True
    Mark = 1
    frmMain.Hide
    frmPopUp.Show
    Unload frmConfig
    frmPopUp.lblName.Caption = Trim(Mid(X, 3, 30))
    frmPopUp.lblType.Caption = Mid(X, 33, 8)
    frmPopUp.lblRate.Caption = "P " & FormatNumber(Trim(Mid(X, 41, 2)), 2) & " / hour"
    frmPopUp.lblLogIn.Caption = Mid(X, 43, Len(X) - 43)
    frmPopUp.cmdLogout.Enabled = True
    frmTrayMenu.mnuLogOut.Enabled = True
  Case "LT" 'set limited account
    'DATA FORMAT
    'LT + (Name-30Chr) + (ServiceType-8Chr) + (Amt-3Chr) + (Rate-2Chr) + (LogInTime)
    BlockCtrl_Alt_Del False
    TypeTrans = False
    Mark = 1
    frmMain.Hide
    frmPopUp.Show
    Unload frmConfig
    frmPopUp.lblName2.Caption = Trim(Mid(X, 3, 30))
    frmPopUp.lblType2.Caption = Mid(X, 33, 8)
    frmPopUp.lblAmt2.Caption = FormatNumber(Val(Mid(X, 41, 3)), 2)
    frmPopUp.lblRate2.Caption = "P " & FormatNumber(Val(Mid(X, 44, 2)), 2) & " / hour"
    frmPopUp.lblLogIn2.Caption = Mid(X, 46, Len(X) - 46)
    frmPopUp.lblService.Caption = Formatter(Val(frmPopUp.lblAmt2.Caption) * (60 / Val(Mid(X, 44, 2))))
    frmPopUp.cmdLogout.Enabled = True
    frmTrayMenu.mnuLogOut.Enabled = True
  Case "ET"
    If Len(X) = 3 Then 'Elapse Time
      BlockCtrl_Alt_Del True
      Shells.MinimizeAll
      Unload frmPopUp
      frmMain.WindowState = vbMaximized
      frmMain.Show
      tmrDelayer.Enabled = True
    ElseIf X = "ET1@" Then 'Lock
      Locked = True
      BlockCtrl_Alt_Del True
      Unload frmPopUp
      Unload frmConfig
      frmMain.Show
    Else 'Unlock
      Locked = False
      BlockCtrl_Alt_Del False
      Unload frmConfig
      frmMain.Hide
    End If
  Case "T1" ' time data for open acnt
    frmPopUp.lblDuration.Caption = Formatter(Val(Trim(Mid(X, 3, 4))))
    frmPopUp.lblAmt.Caption = Mid(X, 7, Len(X) - 7) & ".00"
    frmPopUp.Caption = frmPopUp.lblDuration.Caption
    frmPopUp.Tray.ToolTip = frmPopUp.Caption & "(P" & Trim(Mid(X, 7, Len(X) - 7)) & ")"
    
  Case "T2" ' time data for limited acnt
    frmPopUp.lblDuration2.Caption = Formatter(Val(Trim(Mid(X, 3, 4))))
    frmPopUp.Caption = frmPopUp.lblDuration2.Caption
    frmPopUp.Tray.ToolTip = frmPopUp.Caption & "(P" & Trim(Mid(X, 7, 3)) & ")"
  Case "SD" 'Shutdown
    ExitWindowsEx 1, 0
    End
  Case "SZ" 'Synchronize
    Shock.Rewind
  Case "MT" 'Mute
    If X = "MT1@" Then
      'VolumeMaster.Mute = True 'Mute
    Else
      'VolumeMaster.Mute = False 'Unmute
    End If
  Case "VL" 'Volume
    'VolumeMaster.Volume = Val(Mid(X, 3, Len(X) - 3))
  Case "ER" 'Error
    If X = "ER1@" Then 'ER1 - ID# is already used
      tmrConnector.Enabled = False
      Sucket.Close
      Light.BackColor = vbRed
      MsgBox "ID# " & DbRst!CompID & " is currently used!" & vbCrLf & "Please choose another ID#", vbExclamation
      frmConfig.Show 1
    ElseIf X = "ER2@" Then 'ER2 - server is full- cant connect
      MsgBox "Server is full can't connect!" & vbCrLf & "Terminating Program...", vbExclamation
      Sucket.Close
      End
    Else 'ER3 - ID# is out of range!
      tmrConnector.Enabled = False
      Sucket.Close
      Light.BackColor = vbRed
      MsgBox "ID# is out of range!" & vbCrLf & "Server only allows ID # 1 to " & Mid(X, 4, Len(X) - 4), vbExclamation
      frmConfig.Show 1
    End If
  End Select
  
  If ExtraCommand <> "" Then
    X = ExtraCommand
    ExtraCommand = ""
    GoTo Alex
  End If
End Sub

Private Sub Sucket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  Sucket.Close
  Sucket.Connect Trim(DbRst!SvrIP)
End Sub

Private Function Formatter(b As Integer) As String
  If (b Mod 60) < 30 Then
    Formatter = Str(FormatNumber(b / 60, 0)) & " hrs " & (b Mod 60) & " mins"
  Else
    Formatter = Str(FormatNumber(b / 60, 0) - 1) & " hrs " & (b Mod 60) & " mins"
  End If
End Function

Private Sub Timer1_Timer()
On Error Resume Next
  Text1.SetFocus
  Me.WindowState = vbMaximized
  Shock.Width = Me.Width - 465
  Shock.Height = Me.Height - 975
End Sub

Private Sub tmrConnector_Timer()
  Select Case Sucket.State
  Case 8, 4, 3, 0
    Sucket.Close
    Light.BackColor = vbRed
    Sucket.Connect Trim(DbRst!SvrIP)
  End Select
End Sub

Private Sub tmrDelayer_Timer()
  frmMain.Hide
  frmMain.WindowState = vbMaximized
  frmMain.Show
  tmrDelayer.Enabled = False
End Sub
