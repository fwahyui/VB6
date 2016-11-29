VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   3465
      Left            =   30
      Picture         =   "frmSplash.frx":0000
      Top             =   210
      Width           =   4635
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This programmed is licensed to:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   5010
      TabIndex        =   3
      Top             =   15
      Width           =   2880
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   4425
      Left            =   15
      Top             =   0
      Width           =   7365
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ABC"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   615
      Left            =   5640
      TabIndex        =   2
      Top             =   255
      Width           =   1515
   End
   Begin VB.Image Image4 
      Height          =   150
      Left            =   6270
      Picture         =   "frmSplash.frx":10F1
      Top             =   3915
      Width           =   1005
   End
   Begin VB.Image Image3 
      Height          =   270
      Left            =   2970
      Picture         =   "frmSplash.frx":192B
      Top             =   3660
      Width           =   4245
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COPYRIGHT © 3JNet 2008   |  Developed by: Jomar I. Pabuaya"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   2190
      TabIndex        =   1
      Top             =   4125
      Width           =   4590
   End
   Begin VB.Label lblStat 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6420
      TabIndex        =   0
      Top             =   3105
      Width           =   840
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      Top             =   3645
      Width           =   7695
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   3810
      Picture         =   "frmSplash.frx":471B
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   16110
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'API for Top Most form
Private Declare Function SetWindowPos Lib "user32" _
    (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Private Const HWND_NOTOPMOST = -2

Dim isOn As Boolean

'Public Function ShowSplash()
'
'    'show form
'    SetWindowPos Me.hWnd, HWND_TOPMOST, _
'    0, 0, 0, 0, FLAGS
'    Me.Show
'
'    DoEvents
'    DoEvents
'    DoEvents
'
'    'continue loading...
'    Call modMain.Main_AfterSD
'
'End Function


Public Function ShowForm()
    
    lblStat.Caption = ""
    'show form
    Me.Show
End Function

Public Function UnloadSplash()
    Unload Me
End Function

Private Sub Form_Deactivate()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub


