VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MWARNET - CLIENT MOVED"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7500
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "moveclient.frx":0000
   ScaleHeight     =   5100
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox MT11 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox MT09 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5520
      TabIndex        =   26
      Top             =   3720
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5400
      Top             =   480
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      Height          =   375
      Left            =   4560
      TabIndex        =   25
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PROCESS"
      Height          =   375
      Left            =   5760
      TabIndex        =   24
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox MT01 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "0"
      Top             =   3720
      Width           =   735
   End
   Begin VB.TextBox MT02 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   3720
      Width           =   735
   End
   Begin VB.TextBox MT03 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "0"
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox MT04 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "0"
      Top             =   3720
      Width           =   735
   End
   Begin VB.TextBox MT05 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "0"
      Top             =   3720
      Width           =   735
   End
   Begin VB.TextBox MT06 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "0"
      Top             =   3720
      Width           =   735
   End
   Begin VB.TextBox MT07 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4800
      TabIndex        =   15
      Text            =   "Ready"
      Top             =   3720
      Width           =   735
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "moveclient.frx":7878
      Left            =   5880
      List            =   "moveclient.frx":788B
      TabIndex        =   13
      Text            =   "PC01"
      Top             =   1200
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "moveclient.frx":78AD
      Left            =   2640
      List            =   "moveclient.frx":78C0
      TabIndex        =   10
      Text            =   "PC01"
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox ATX01 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "0"
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox ATX02 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox ATX03 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "0"
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox ATX04 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "0"
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox ATX05 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "0"
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox ATX06 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "0"
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox ATX07 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4800
      TabIndex        =   2
      Text            =   "Ready"
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox ATX09 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox ATX11 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   2280
      Width           =   495
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
      TabIndex        =   23
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      Height          =   975
      Left            =   600
      Top             =   3120
      Width           =   6255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "  Counter    Duration   Step      Cost     Discount    Total       Type      Timein     Index"
      Height          =   255
      Left            =   720
      TabIndex        =   22
      Top             =   3360
      Width           =   6015
   End
   Begin VB.Shape Shape1 
      Height          =   975
      Left            =   600
      Top             =   1800
      Width           =   6255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "TO CLINET NUMBER :"
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
      Left            =   3840
      TabIndex        =   12
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MOVE THIS CLINET :"
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
      Left            =   720
      TabIndex        =   11
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "  Counter    Duration   Step      Cost     Discount    Total       Type      Timein     Index"
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   1920
      Width           =   6015
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000FF&
      Caption         =   "SOURCE"
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
      TabIndex        =   14
      Top             =   1560
      Width           =   1335
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If ATX02.Text = "" Then Exit Sub
MT01.Text = ATX01.Text
MT02.Text = ATX02.Text
MT03.Text = ATX03.Text
MT04.Text = ATX04.Text
MT05.Text = ATX05.Text
MT06.Text = ATX06.Text
MT07.Text = ATX07.Text
MT09.Text = ATX09.Text
MT11.Text = ATX11.Text
transfersdata
End Sub

Private Sub transfersdata()
'---- SOURCE
If Combo1.Text = "PC01" Then Form1.ATX07.Text = "MOVE"
If Combo1.Text = "PC02" Then Form1.BTX07.Text = "MOVE"
If Combo1.Text = "PC03" Then Form1.CTX07.Text = "MOVE"
If Combo1.Text = "PC04" Then Form1.DTX07.Text = "MOVE"
If Combo1.Text = "PC05" Then Form1.ETX07.Text = "MOVE"
'---- DESTINATION
If Combo2.Text = "PC01" Then
    Form1.ATX01.Text = ATX01.Text
    Form1.ATX02.Text = ATX02.Text
    Form1.ATX03.Text = ATX03.Text
    Form1.ATX04.Text = ATX04.Text
    Form1.ATX05.Text = ATX05.Text
    Form1.ATX06.Text = ATX06.Text
    Form1.ATX07.Text = ATX07.Text
    Form1.ATX09.Text = ATX09.Text
    Form1.ATX11.Text = ATX11.Text
    End If
If Combo2.Text = "PC02" Then
    Form1.BTX01.Text = ATX01.Text
    Form1.BTX02.Text = ATX02.Text
    Form1.BTX03.Text = ATX03.Text
    Form1.BTX04.Text = ATX04.Text
    Form1.BTX05.Text = ATX05.Text
    Form1.BTX06.Text = ATX06.Text
    Form1.BTX07.Text = ATX07.Text
    Form1.BTX09.Text = ATX09.Text
    Form1.BTX11.Text = ATX11.Text
    End If
If Combo2.Text = "PC03" Then
    Form1.CTX01.Text = ATX01.Text
    Form1.CTX02.Text = ATX02.Text
    Form1.CTX03.Text = ATX03.Text
    Form1.CTX04.Text = ATX04.Text
    Form1.CTX05.Text = ATX05.Text
    Form1.CTX06.Text = ATX06.Text
    Form1.CTX07.Text = ATX07.Text
    Form1.CTX09.Text = ATX09.Text
    Form1.CTX11.Text = ATX11.Text
    End If
If Combo2.Text = "PC04" Then
    Form1.DTX01.Text = ATX01.Text
    Form1.DTX02.Text = ATX02.Text
    Form1.DTX03.Text = ATX03.Text
    Form1.DTX04.Text = ATX04.Text
    Form1.DTX05.Text = ATX05.Text
    Form1.DTX06.Text = ATX06.Text
    Form1.DTX07.Text = ATX07.Text
    Form1.DTX09.Text = ATX09.Text
    Form1.DTX11.Text = ATX11.Text
    End If
If Combo2.Text = "PC05" Then
    Form1.ETX01.Text = ATX01.Text
    Form1.ETX02.Text = ATX02.Text
    Form1.ETX03.Text = ATX03.Text
    Form1.ETX04.Text = ATX04.Text
    Form1.ETX05.Text = ATX05.Text
    Form1.ETX06.Text = ATX06.Text
    Form1.ETX07.Text = ATX07.Text
    Form1.ETX09.Text = ATX09.Text
    Form1.ETX11.Text = ATX11.Text
    End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
If Combo1.Text = "PC01" Then SOURCE01
If Combo1.Text = "PC02" Then SOURCE02
If Combo1.Text = "PC03" Then SOURCE03
If Combo1.Text = "PC04" Then SOURCE04
If Combo1.Text = "PC05" Then SOURCE05
End Sub

Private Sub SOURCE05()
'---PC05
ATX01.Text = Form1.ETX01.Text
ATX02.Text = Form1.ETX02.Text
ATX03.Text = Form1.ETX03.Text
ATX04.Text = Form1.ETX04.Text
ATX05.Text = Form1.ETX05.Text
ATX06.Text = Form1.ETX06.Text
ATX07.Text = Form1.ETX07.Text
ATX09.Text = Form1.ETX09.Text
ATX11.Text = Form1.ETX11.Text
End Sub

Private Sub SOURCE04()
'---PC04
ATX01.Text = Form1.DTX01.Text
ATX02.Text = Form1.DTX02.Text
ATX03.Text = Form1.DTX03.Text
ATX04.Text = Form1.DTX04.Text
ATX05.Text = Form1.DTX05.Text
ATX06.Text = Form1.DTX06.Text
ATX07.Text = Form1.DTX07.Text
ATX09.Text = Form1.DTX09.Text
ATX11.Text = Form1.DTX11.Text
End Sub

Private Sub SOURCE03()
'---PC03
ATX01.Text = Form1.CTX01.Text
ATX02.Text = Form1.CTX02.Text
ATX03.Text = Form1.CTX03.Text
ATX04.Text = Form1.CTX04.Text
ATX05.Text = Form1.CTX05.Text
ATX06.Text = Form1.CTX06.Text
ATX07.Text = Form1.CTX07.Text
ATX09.Text = Form1.CTX09.Text
ATX11.Text = Form1.CTX11.Text
End Sub
Private Sub SOURCE02()
'---PC02
ATX01.Text = Form1.BTX01.Text
ATX02.Text = Form1.BTX02.Text
ATX03.Text = Form1.BTX03.Text
ATX04.Text = Form1.BTX04.Text
ATX05.Text = Form1.BTX05.Text
ATX06.Text = Form1.BTX06.Text
ATX07.Text = Form1.BTX07.Text
ATX09.Text = Form1.BTX09.Text
ATX11.Text = Form1.BTX11.Text
End Sub

Private Sub SOURCE01()
'---PC01
ATX01.Text = Form1.ATX01.Text
ATX02.Text = Form1.ATX02.Text
ATX03.Text = Form1.ATX03.Text
ATX04.Text = Form1.ATX04.Text
ATX05.Text = Form1.ATX05.Text
ATX06.Text = Form1.ATX06.Text
ATX07.Text = Form1.ATX07.Text
ATX09.Text = Form1.ATX09.Text
ATX11.Text = Form1.ATX11.Text
End Sub
