VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   Picture         =   "frmidealpassword.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5760
      Top             =   5880
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   7440
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   5400
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   315
      Left            =   7440
      TabIndex        =   0
      Top             =   4560
      Width           =   2415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Contact For Registered Version."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1440
      TabIndex        =   6
      Top             =   600
      Width           =   7680
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   7320
      Shape           =   4  'Rounded Rectangle
      Top             =   5295
      Width           =   2655
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   7320
      Shape           =   4  'Rounded Rectangle
      Top             =   4440
      Width           =   2655
   End
   Begin VB.Shape Shape5 
      DrawMode        =   9  'Not Mask Pen
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   7580
      Shape           =   4  'Rounded Rectangle
      Top             =   7480
      Width           =   1545
   End
   Begin VB.Shape Shape4 
      DrawMode        =   9  'Not Mask Pen
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   8800
      Shape           =   4  'Rounded Rectangle
      Top             =   7100
      Width           =   1110
   End
   Begin VB.Shape Shape3 
      DrawMode        =   9  'Not Mask Pen
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   7580
      Shape           =   4  'Rounded Rectangle
      Top             =   7080
      Width           =   1110
   End
   Begin VB.Shape Shape2 
      DrawMode        =   9  'Not Mask Pen
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   9240
      Shape           =   4  'Rounded Rectangle
      Top             =   7450
      Width           =   315
   End
   Begin VB.Shape Shape1 
      DrawMode        =   6  'Mask Pen Not
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   9620
      Shape           =   4  'Rounded Rectangle
      Top             =   7485
      Width           =   315
   End
   Begin VB.Image imgexit 
      Height          =   315
      Left            =   9620
      MouseIcon       =   "frmidealpassword.frx":143B5
      MousePointer    =   99  'Custom
      Top             =   7485
      Width           =   315
   End
   Begin VB.Image imghint 
      Height          =   315
      Left            =   9240
      MouseIcon       =   "frmidealpassword.frx":146BF
      MousePointer    =   99  'Custom
      Top             =   7450
      Width           =   315
   End
   Begin VB.Image imgenter 
      Height          =   285
      Left            =   7580
      MouseIcon       =   "frmidealpassword.frx":149C9
      MousePointer    =   99  'Custom
      Top             =   7110
      Width           =   1080
   End
   Begin VB.Image imgclear 
      Height          =   285
      Left            =   8800
      MouseIcon       =   "frmidealpassword.frx":14CD3
      MousePointer    =   99  'Custom
      Top             =   7100
      Width           =   1080
   End
   Begin VB.Image imgchangepassword 
      Height          =   285
      Left            =   7580
      MouseIcon       =   "frmidealpassword.frx":14FDD
      MousePointer    =   99  'Custom
      Top             =   7480
      Width           =   1560
   End
   Begin VB.Label lbldate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12/11/02"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   9000
      TabIndex        =   5
      Top             =   5880
      Width           =   630
   End
   Begin VB.Label lbltime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12:07:10 am"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   7560
      TabIndex        =   4
      Top             =   5880
      Width           =   930
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   8160
      TabIndex        =   3
      Top             =   5040
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Secrete Code:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   8040
      TabIndex        =   1
      Top             =   4200
      Width           =   1170
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub imgenter_Click()
frmpatient.Show
Unload Me
End Sub

Private Sub imgexit_Click()
End
End Sub

Private Sub Timer1_Timer()
lbltime.Caption = Time
lbldate.Caption = Date

End Sub
