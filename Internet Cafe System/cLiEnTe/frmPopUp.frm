VERSION 5.00
Begin VB.Form frmPopUp 
   BorderStyle     =   0  'None
   ClientHeight    =   4140
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   4785
   Begin CLIENTE.TrayArea Tray 
      Left            =   960
      Top             =   0
      _ExtentX        =   900
      _ExtentY        =   900
      Icon            =   "frmPopUp.frx":0000
      ToolTip         =   "PC Empire Time Log"
   End
   Begin VB.Frame frOpen 
      BackColor       =   &H00CECECE&
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   3495
      Begin CLIENTE.SkinButton cmdLogout 
         Height          =   360
         Left            =   720
         TabIndex        =   33
         Top             =   2760
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   635
         Caption         =   "Log Out"
      End
      Begin VB.Label lblAmt 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
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
         Left            =   1200
         TabIndex        =   14
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label lblDuration 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   13
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label lblLogIn 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   12
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Log-In Time:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount:       P"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Duration:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblCompNum 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   7
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label7 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Service Type:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblType 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   3
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label lblRate 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   2
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Rate:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Open Account"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   0
         TabIndex        =   11
         Top             =   120
         Width           =   3495
      End
   End
   Begin VB.Frame frLimited 
      BackColor       =   &H00CECECE&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   3255
      Left            =   600
      TabIndex        =   15
      Top             =   720
      Visible         =   0   'False
      Width           =   3495
      Begin VB.Label lblDuration2 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
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
         Left            =   1200
         TabIndex        =   31
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label lblLogIn2 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   30
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label11 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Log-In Time:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label12 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount:       P"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label13 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Duration:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label14 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Service Time:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label lblService 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   24
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label lblCompNum2 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   23
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblName2 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   21
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label lblAmt2 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   20
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label Label9 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Service Type:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblType2 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   18
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label lblRate2 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   17
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label17 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Rate:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Limited Account"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   0
         TabIndex        =   29
         Top             =   120
         Width           =   3495
      End
   End
   Begin CLIENTE.SkinForm SkinnedForm 
      Height          =   735
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
      Caption         =   "Account Summary"
      CaptionTop      =   250
   End
End
Attribute VB_Name = "frmPopUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub cmdLogout_Click()
  If MsgBox("Are you sure to Log-Out!", vbOKCancel, "Confirm") = vbOK Then
    frmMain.Sucket.SendData "LO"
    BlockCtrl_Alt_Del True
    frmMain.Show
    Unload frmPopUp
  End If
End Sub

Private Sub Form_Load()
  SkinnedForm.LoadSkin Me
  Me.lblCompNum.Caption = DbRst!CompID
  Me.lblCompNum2.Caption = DbRst!CompID
  If TypeTrans = True Then 'open
    Me.frOpen.Visible = True
  Else 'limited
    Me.frLimited.Visible = True
    frmTrayMenu.mnuLogOut.Enabled = False
  End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  cmdLogout.Refresh
End Sub

Private Sub Form_Resize()
  If Me.WindowState = 1 Then
    frmPopUp.Hide
    Tray.Visible = True
  End If
End Sub

Private Sub frOpen_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  cmdLogout.Refresh
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  cmdLogout.Refresh
End Sub

Private Sub lblAmt_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  cmdLogout.Refresh
End Sub

Private Sub SkinnedForm_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  cmdLogout.Refresh
End Sub

Private Sub Tray_DblClick()
  Tray.Visible = False
  frmPopUp.WindowState = 0
  frmPopUp.Show
End Sub

Private Sub Tray_MouseUp(Button As Integer)
  If Button = 2 Then
    Select Case TypeTrans
    Case True 'Open
      frmTrayMenu.mnuLogOut.Visible = True
    Case False 'Limited
      frmTrayMenu.mnuLogOut.Visible = False
    End Select
    frmTrayMenu.mnuMax.Visible = True
    PopupMenu frmTrayMenu.mnuTray
  End If
End Sub

