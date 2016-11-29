VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MWARNET - ONLINE TIMER"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4200
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "tmeronline.frx":0000
   ScaleHeight     =   4470
   ScaleWidth      =   4200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "OK"
      Height          =   615
      Left            =   2400
      TabIndex        =   28
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   26
      Text            =   "0"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   25
      Text            =   "0"
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   ">"
      Height          =   255
      Left            =   1800
      TabIndex        =   24
      Top             =   3720
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      Height          =   255
      Left            =   1560
      TabIndex        =   23
      Top             =   3720
      Width           =   255
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "tmeronline.frx":6D74
      Left            =   1080
      List            =   "tmeronline.frx":6D87
      TabIndex        =   21
      Top             =   3360
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3600
      Top             =   240
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   1440
      Width           =   735
   End
   Begin VB.CheckBox Check5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "On"
      Height          =   255
      Left            =   3000
      TabIndex        =   15
      Top             =   2880
      Width           =   615
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "On"
      Height          =   255
      Left            =   3000
      TabIndex        =   14
      Top             =   2520
      Width           =   615
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "On"
      Height          =   255
      Left            =   3000
      TabIndex        =   13
      Top             =   2160
      Width           =   615
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "On"
      Height          =   255
      Left            =   3000
      TabIndex        =   12
      Top             =   1800
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "On"
      Height          =   255
      Left            =   3000
      TabIndex        =   11
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Hours"
      Height          =   255
      Left            =   360
      TabIndex        =   27
      Top             =   3720
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Turn Off"
      Height          =   255
      Left            =   360
      TabIndex        =   22
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CLIENTS AUTOMATIC TURNOFF"
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
      Left            =   480
      TabIndex        =   10
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label ELBL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label DLBL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label CLBL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label BLBL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label ALBL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label ATX10 
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   255
      Left            =   660
      TabIndex        =   4
      Top             =   1440
      Width           =   255
   End
   Begin VB.Image Image4 
      Height          =   240
      Left            =   480
      Picture         =   "tmeronline.frx":6DA9
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label BTX10 
      BackStyle       =   0  'Transparent
      Caption         =   "02"
      Height          =   255
      Left            =   660
      TabIndex        =   3
      Top             =   1800
      Width           =   255
   End
   Begin VB.Image Image5 
      Height          =   240
      Left            =   480
      Picture         =   "tmeronline.frx":6FC1
      Top             =   1800
      Width           =   135
   End
   Begin VB.Label CTX10 
      BackStyle       =   0  'Transparent
      Caption         =   "03"
      Height          =   255
      Left            =   660
      TabIndex        =   2
      Top             =   2160
      Width           =   255
   End
   Begin VB.Image Image6 
      Height          =   240
      Left            =   480
      Picture         =   "tmeronline.frx":71D9
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label DTX10 
      BackStyle       =   0  'Transparent
      Caption         =   "04"
      Height          =   255
      Left            =   660
      TabIndex        =   1
      Top             =   2520
      Width           =   255
   End
   Begin VB.Image Image8 
      Height          =   240
      Left            =   480
      Picture         =   "tmeronline.frx":73F1
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label ETX10 
      BackStyle       =   0  'Transparent
      Caption         =   "05"
      Height          =   255
      Left            =   660
      TabIndex        =   0
      Top             =   2880
      Width           =   255
   End
   Begin VB.Image Image9 
      Height          =   240
      Left            =   480
      Picture         =   "tmeronline.frx":7609
      Top             =   2880
      Width           =   135
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MWARNET 2 - FREEWARE EDITION
'COPYRIGHT(C) 2007 MTechnologi Bali Indonesia
'Programed by A.A.Ngr.Manik Artawan
'e-mail : gungmanik@telkom.net
'---------------------------------------------
'THANK YOu FOR DOWNLOAD THIS SMALL APPLICATION
'---------------------------------------------


Private Function SecondsToTime(ByVal dSeconds As Double) As String
    SecondsToTime = Format(DateAdd("s", dSeconds, "00:00:00"), "HH:mm:ss")
End Function

Private Sub Command1_Click()
Text6.Text = Val(Text6.Text) - 1
If Text6.Text < 1 Then Text6.Text = 0
Text7.Text = Val(Text6.Text) * 60 * 60
End Sub

Private Sub Command2_Click()
Text6.Text = Val(Text6.Text) + 1
Text7.Text = Val(Text6.Text) * 60 * 60
End Sub

Private Sub Command3_Click()
If Combo1.Text = "PC01" Then Text1.Text = Text7.Text
If Combo1.Text = "PC02" Then Text2.Text = Text7.Text
If Combo1.Text = "PC03" Then Text3.Text = Text7.Text
If Combo1.Text = "PC04" Then Text4.Text = Text7.Text
If Combo1.Text = "PC05" Then Text5.Text = Text7.Text
End Sub

Private Sub Form_Load()

End Sub

Private Sub Timer1_Timer()
If Check1.Value = Checked Then
    If Text1.Text = "" Then
        Else
        Form1.ATX07.Text = "ONLINE"
        Text1.Text = Val(Text1.Text - 1)
        ALBL.Caption = SecondsToTime(Text1.Text)
        If Text1.Text = 0 Then
            Form1.Oprmt6.Value = True
            Form1.Combo1.Text = "PC01"
            Form1.CmTurnoff.Value = True
            Text1.Text = ""
            End If
        End If
    End If
If Check2.Value = Checked Then
    If Text2.Text = "" Then
        Else
        Form1.BTX07.Text = "ONLINE"
        Text2.Text = Val(Text2.Text - 1)
        BLBL.Caption = SecondsToTime(Text2.Text)
        If Text2.Text = 0 Then
            Form1.Oprmt6.Value = True
            Form1.Combo1.Text = "PC02"
            Form1.CmTurnoff.Value = True
            Text2.Text = ""
            End If
        End If
    End If
If Check3.Value = Checked Then
    If Text3.Text = "" Then
        Else
        Form1.CTX07.Text = "ONLINE"
        Text3.Text = Val(Text3.Text - 1)
        CLBL.Caption = SecondsToTime(Text3.Text)
        If Text3.Text = 0 Then
            Form1.Oprmt6.Value = True
            Form1.Combo1.Text = "PC03"
            Form1.CmTurnoff.Value = True
            Text3.Text = ""
            End If
        End If
    End If
If Check4.Value = Checked Then
    If Text4.Text = "" Then
        Else
        Form1.DTX07.Text = "ONLINE"
        Text4.Text = Val(Text4.Text - 1)
        DLBL.Caption = SecondsToTime(Text4.Text)
        If Text4.Text = 0 Then
            Form1.Oprmt6.Value = True
            Form1.Combo1.Text = "PC04"
            Form1.CmTurnoff.Value = True
            Text4.Text = ""
            End If
        End If
    End If
If Check5.Value = Checked Then
    If Text5.Text = "" Then
        Else
        Form1.ETX07.Text = "ONLINE"
        Text5.Text = Val(Text5.Text - 1)
        ELBL.Caption = SecondsToTime(Text5.Text)
        If Text5.Text = 0 Then
            Form1.Oprmt6.Value = True
            Form1.Combo1.Text = "PC05"
            Form1.CmTurnoff.Value = True
            Text5.Text = ""
            End If
        End If
    End If
End Sub
