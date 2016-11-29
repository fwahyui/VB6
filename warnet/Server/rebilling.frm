VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "MWARNET - REBILLING"
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7500
   LinkTopic       =   "Form8"
   Picture         =   "rebilling.frx":0000
   ScaleHeight     =   5100
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2520
      TabIndex        =   17
      Top             =   1800
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5520
      Top             =   360
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1440
      TabIndex        =   15
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2520
      TabIndex        =   13
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox MT07 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4800
      TabIndex        =   10
      Text            =   "Ready"
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox MT06 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "0"
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox MT05 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "0"
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox MT04 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "0"
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox MT03 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "0"
      Top             =   3600
      Width           =   495
   End
   Begin VB.TextBox MT02 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox MT01 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "0"
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PROCESS"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      Height          =   375
      Left            =   5640
      TabIndex        =   2
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox MT09 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox MT11 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Time Out : ( 00:00:00 )"
      Height          =   255
      Left            =   600
      TabIndex        =   18
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Time       :"
      Height          =   255
      Left            =   600
      TabIndex        =   16
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Time in    : ( 00:00:00 )"
      Height          =   255
      Left            =   600
      TabIndex        =   14
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "  Counter    Duration   Step      Cost     Discount    Total       Type      Timein     Index"
      Height          =   255
      Left            =   720
      TabIndex        =   12
      Top             =   3240
      Width           =   6015
   End
   Begin VB.Shape Shape2 
      Height          =   975
      Left            =   600
      Top             =   3120
      Width           =   6255
   End
   Begin VB.Label Label6 
      BackColor       =   &H000000FF&
      Caption         =   "DESTINATION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   2880
      Width           =   1335
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Timer1_Timer()
Text2.Text = Time
End Sub
