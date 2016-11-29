VERSION 5.00
Begin VB.Form frmPass 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1695
   ClientLeft      =   3765
   ClientTop       =   3015
   ClientWidth     =   3825
   ControlBox      =   0   'False
   Icon            =   "frmPass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   3825
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   350
      Left            =   2520
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   350
      Left            =   1200
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtPass 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "Ï"
      TabIndex        =   0
      Top             =   540
      Width           =   2175
   End
   Begin VB.Image Image2 
      Height          =   765
      Left            =   240
      Picture         =   "frmPass.frx":000C
      Top             =   840
      Width           =   825
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00000000&
      Caption         =   " ADMIN"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   3855
   End
   Begin VB.Label P 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "frmPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PassNum As Integer

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOk_Click()
Dim i As Byte
Dim tmpPwd As Integer
  
  txtPass.Text = Trim(txtPass.Text)
  For i = 1 To Len(txtPass.Text)
    tmpPwd = tmpPwd + (Asc(Mid(txtPass.Text, i, 1)))
  Next i
  If SetPass = True Then 'set password
    Open App.Path & "\Tmp.txt" For Output As #1
      Print #1, tmpPwd
    Close #1
    Unload Me
  Else 'get password
    If PassNum = tmpPwd Then
      Unload Me
      If TrayPass = False Then
        Select Case Config
        Case True
          frmOption.Show 1
        Case False
          frmClientControl.Show 1
        End Select
      Else 'Tray password
        TrayPass = False
        frmMain.WindowState = vbMaximized
        frmMain.Visible = True
      End If
    Else
      Unload Me
    End If
  End If
End Sub

Private Sub Form_Load()
Dim i As Integer
  Me.Icon = frmMain.Icon
  If SetPass = True Then lblTitle.Caption = " SET PASSWORD"
  If Dir(App.Path & "\Tmp.txt", vbNormal) <> "" Then
    Open App.Path & "\Tmp.txt" For Input As #1
      PassNum = Val(Input(FileLen(App.Path & "\Tmp.txt"), #1))
    Close #1
  Else
    PassNum = 0
  End If
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  MoveForm Me
End Sub
