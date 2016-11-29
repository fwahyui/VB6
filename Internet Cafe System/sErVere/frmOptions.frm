VERSION 5.00
Begin VB.Form frmOption 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5475
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4800
   ControlBox      =   0   'False
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   4800
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox cmbGRR 
      Height          =   285
      Left            =   3000
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox cmbIR 
      Height          =   285
      Left            =   3000
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.ComboBox cmbMinAmt 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmOptions.frx":000C
      Left            =   3000
      List            =   "frmOptions.frx":0031
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdSetPass 
      Caption         =   "Set Password"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4920
      Width           =   1335
   End
   Begin VB.ComboBox cmbNumComps 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmOptions.frx":005C
      Left            =   3000
      List            =   "frmOptions.frx":0141
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "C&ancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   4920
      Width           =   1215
   End
   Begin VB.ComboBox cmbGRR2 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmOptions.frx":0268
      Left            =   3000
      List            =   "frmOptions.frx":02AB
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1680
      Width           =   975
   End
   Begin VB.ComboBox cmbIR2 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmOptions.frx":0303
      Left            =   3000
      List            =   "frmOptions.frx":0346
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Minimum Amount:"
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
      Left            =   720
      TabIndex        =   12
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label lblType 
      BackColor       =   &H00000000&
      Caption         =   " OPTIONS"
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
      TabIndex        =   11
      Top             =   0
      Width           =   4815
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Minimum Amount:"
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
      Left            =   720
      TabIndex        =   10
      Top             =   2085
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "No. of Client Comp(s):"
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
      Left            =   720
      TabIndex        =   9
      Top             =   1005
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Games/Rental Rate:"
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
      Left            =   720
      TabIndex        =   8
      Top             =   1725
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Internet Rate:"
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
      Left            =   720
      TabIndex        =   7
      Top             =   1365
      Width           =   1575
   End
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdSave_Click()
Dim i As Integer
  If cmbIR.Text = "" Or cmbGRR.Text = "" Then Exit Sub
  Rst.MoveFirst
  Rst!Month = Val(cmbIR.Text)
  Rst!Day = Val(cmbGRR.Text)
  Rst!CompNum = Val(cmbNumComps.Text)
  Rst!Year = Val(cmbMinAmt.Text)
  Rst.Update
  NumberComps = Val(cmbNumComps.Text)
  frmMain.Fill_ListView
  Main
  Unload Me
End Sub

Private Sub cmdSetPass_Click()
  SetPass = True
  frmPass.Show 1
End Sub

Private Sub Form_Load()
  Me.Icon = frmMain.Icon
  Rst.MoveFirst
  cmbIR.Text = Rst!Month
  cmbGRR.Text = Rst!Day
  cmbNumComps.Text = Rst!CompNum
  cmbMinAmt.Text = Rst!Year
End Sub

Private Sub lblType_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  MoveForm Me
End Sub
