VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "User's Login"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPass 
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1170
      MaxLength       =   20
      PasswordChar    =   "="
      TabIndex        =   1
      Top             =   2280
      Width           =   3105
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   2970
      Width           =   1395
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Login"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   2970
      Width           =   1395
   End
   Begin MSDataListLib.DataCombo dcUser 
      Height          =   315
      Left            =   1170
      TabIndex        =   0
      Top             =   1740
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      ForeColor       =   &H00004040&
      Height          =   195
      Left            =   150
      TabIndex        =   5
      Top             =   1785
      Width           =   840
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      ForeColor       =   &H00004040&
      Height          =   195
      Left            =   150
      TabIndex        =   4
      Top             =   2310
      Width           =   750
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   0
      Picture         =   "frmLogin.frx":0000
      Top             =   0
      Width           =   4830
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    modMain.CloseMe = True
    Unload Me
End Sub

Private Sub cmdLogin_Click()
    'Verify
    If dcUser.Text = "" Then dcUser.SetFocus: Exit Sub
    If txtPass.Text = "" Then txtPass.SetFocus: Exit Sub
    Dim strPass As String
    strPass = getValueAt("SELECT PK,Password FROM Users WHERE PK=" & dcUser.BoundText, "Password")
    strPass = Enc.DecryptString(strPass)
    'Very short code of login system
    If LCase(txtPass.Text) = LCase(strPass) Then
        With CurrUser
            .USER_NAME = dcUser.Text
            .USER_PK = dcUser.BoundText
            .USER_ISADMIN = CBool(changeYNValue(getValueAt("SELECT PK,Admin FROM Users WHERE PK=" & dcUser.BoundText, "Admin")))
        End With
        Unload Me
    Else
        MsgBox "Invalid password.Please try again!", vbExclamation
        txtPass.SetFocus
    End If
    strPass = vbNullString
End Sub

Private Sub Form_Load()
    bind_dc "SELECT * FROM Users", "UserID", dcUser, "PK"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub txtPass_Change()
    txtPass.SelStart = Len(txtPass.Text)
End Sub

Private Sub txtPass_GotFocus()
    HLText txtPass
End Sub

