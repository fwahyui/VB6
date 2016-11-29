VERSION 5.00
Begin VB.Form frmConfig 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3495
   ClientLeft      =   3930
   ClientTop       =   2850
   ClientWidth     =   4800
   ControlBox      =   0   'False
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   4800
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbID 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "frmConfig.frx":000C
      Left            =   1800
      List            =   "frmConfig.frx":00F1
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdLock 
      BackColor       =   &H00BFA839&
      Caption         =   "Lock"
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton cmdUnlock 
      BackColor       =   &H00BFA839&
      Caption         =   "Unlock"
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton cmdShut 
      BackColor       =   &H00BFA839&
      Caption         =   "Shutdown Win"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00BFA839&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3360
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00BFA839&
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtPass 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox txtIP 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "A D M I N"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   2415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000006&
      X1              =   120
      X2              =   4560
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Controls:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Admin Pass:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Server IP Address:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Computer ID#:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
On Error Resume Next
  
  If Locked = False Then
    frmMain.Sucket.SendData "LK" 'sends lock signal to server
    frmMain.Show
    Locked = True
  End If
  Unload Me
End Sub

Private Sub cmdLock_Click()
On Error Resume Next
  frmMain.Sucket.SendData "LK" 'sends lock signal to server
  Locked = True
  BlockCtrl_Alt_Del True
  cmdLock.Enabled = False
  cmdUnlock.Enabled = True
  Unload Me
  frmMain.Show
  frmConfig.Show 1
End Sub

Private Sub cmdSave_Click()
  Select Case True
  Case Trim(cmbID.Text) = ""
    cmbID.SetFocus
    Exit Sub
  Case Trim(txtIP.Text) = ""
    txtIP.SetFocus
    Exit Sub
  Case Trim(txtPass.Text) = ""
    txtPass.SetFocus
    Exit Sub
  End Select

  DbRst!CompID = Trim(cmbID.Text)
  DbRst!SvrIP = Trim(txtIP.Text)
  DbRst!Pwd = Trim(txtPass.Text)
  DbRst.Update
  Unload Me
  
  If Cur_ID <> DbRst!CompID Then
    Cur_ID = DbRst!CompID
    frmMain.Sucket.Close
    frmMain.Sucket.Connect Trim(DbRst!SvrIP)
  End If
  frmMain.tmrConnector.Enabled = True
  frmMain.Show
End Sub

Private Sub cmdShut_Click()
  ExitWindowsEx 1, 0
End Sub

Private Sub cmdUnlock_Click()
On Error Resume Next
  frmMain.Sucket.SendData "UL" 'sends unlock signal to server
  Locked = False
  BlockCtrl_Alt_Del False
  cmdLock.Enabled = True
  cmdUnlock.Enabled = False
  Unload Me
  frmMain.Hide
  frmConfig.Show
End Sub

Private Sub Form_Load()
  If Locked = True Then
    cmdLock.Enabled = False
  Else
    cmdUnlock.Enabled = False
  End If
  If DbRst!SvrIP = "143" Then cmdCancel.Visible = False
On Error Resume Next
  txtPass.Text = DbRst!Pwd
  cmbID.Text = DbRst!CompID
  txtIP.Text = DbRst!SvrIP
End Sub

