VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MWARNET - PICVIEWER"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9015
   Icon            =   "picviewer.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "picviewer.frx":FA8A
   ScaleHeight     =   6120
   ScaleWidth      =   9015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmfullscreen 
      Caption         =   "Fullscreen"
      Height          =   375
      Left            =   5280
      TabIndex        =   7
      Top             =   5160
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   7920
      Top             =   1200
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000009&
      Caption         =   "Slide show"
      Height          =   255
      Left            =   6480
      TabIndex        =   6
      Top             =   5400
      Width           =   1215
   End
   Begin VB.FileListBox File1 
      Height          =   1845
      Left            =   6480
      TabIndex        =   2
      Top             =   3480
      Width           =   1935
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   6480
      TabIndex        =   1
      Top             =   1680
      Width           =   1935
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6480
      TabIndex        =   0
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Click the filename to view"
      Height          =   255
      Left            =   6480
      TabIndex        =   5
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0"
      Height          =   255
      Left            =   6000
      TabIndex        =   4
      Top             =   735
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PICS VIEWER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   3
      Top             =   480
      Width           =   2055
   End
   Begin VB.Image Image2 
      Height          =   750
      Left            =   4200
      Picture         =   "picviewer.frx":18512
      Top             =   465
      Width           =   750
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   4215
      Left            =   600
      Top             =   1440
      Width           =   5775
   End
   Begin VB.Image Image1 
      Height          =   4215
      Left            =   600
      Picture         =   "picviewer.frx":18F86
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   5775
   End
End
Attribute VB_Name = "Form3"
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


Dim i As Integer
Dim picViewXFileSpec As String
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Sub Check1_Click()
If Timer1.Enabled = False Then
    Timer1.Enabled = True
    Dir1.Enabled = False
    Drive1.Enabled = False
    Else
    Timer1.Enabled = False
    Dir1.Enabled = True
    Drive1.Enabled = True
    i = 0
    End If
End Sub

Private Sub Cmfullscreen_Click()
    On Error Resume Next
    Dim mfilespec As String
    mfilespec = Dir1.Path & "\" & File1.FileName
    mresult = ShellExecute(Me.hwnd, "Open", mfilespec, &H0&, &H0&, SW_RESTORE)
End Sub

Private Sub Dir1_Change()
File1.FileName = "*.jpg"
File1.FileName = Dir1.Path
Blankpic
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1.Path = Drive1.Drive
File1.FileName = "*.jpg"
File1.FileName = Dir1.Path
End Sub

Private Sub File1_Click()
On Error Resume Next
Image1.Picture = LoadPicture(Dir1.Path & "\" & File1.FileName)
End Sub

Private Sub Blankpic()
On Error Resume Next
If File1.ListCount = 0 Then Image1.Picture = LoadPicture("logo.jpg")
End Sub

Private Sub Form_Load()
File1.FileName = "*.jpg"
File1.FileName = Dir1.Path
End Sub

Private Sub Timer1_Timer()
On Error GoTo t
c = File1.ListCount - 1
i = i + 1
If i = 0 Then
    File1.Selected(0) = True
    Else
        If i > c Then i = 0
        File1.Selected(i) = True
    End If
Exit Sub
t:
End Sub
