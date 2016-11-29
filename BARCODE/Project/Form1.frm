VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8955
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   8955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Save As Bitmap"
      Height          =   390
      Left            =   2280
      TabIndex        =   9
      Top             =   5385
      Width           =   1995
   End
   Begin MSComDlg.CommonDialog SaveFile 
      Left            =   6480
      Top             =   5895
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   6750
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   2640
      List            =   "Form1.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   360
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "PREVIEW"
      Height          =   3975
      Left            =   105
      TabIndex        =   3
      Top             =   1305
      Width           =   8730
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   1935
         Left            =   120
         ScaleHeight     =   129
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   253
         TabIndex        =   4
         Top             =   360
         Width           =   3795
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Text            =   "750103131130"
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   4455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Copy To Clipboard"
      Height          =   375
      Left            =   135
      TabIndex        =   0
      Top             =   5385
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "TYPE"
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   7
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "TEXT"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cl As New arisBarcode
Private Sub Command1_Click()
    'Restore original font that might replaced by PDF417
    Picture1.FontName = Me.FontName
    Picture1.FontSize = Me.FontSize
    
    Select Case Combo1.ListIndex
    Case 0
         cl.Code128 Picture1, 0.5, Text1, True
    Case 1
         cl.Code39 Picture1, 0.5, Text1, True, True
    Case 2
         cl.EAN13 Picture1, 1, Text1, True
    Case 3
         cl.EAN8 Picture1, 0.5, Text1, True
    Case 4
         cl.PDF417 Picture1, Text1
    End Select
End Sub

Private Sub Command2_Click()
    Clipboard.Clear
    Clipboard.SetData Picture1.Image, 2
    
End Sub


Private Sub Command4_Click()
    
    On Error GoTo Err_Trap
    SaveFile.Filter = "Pictures (*.jpg;*.bmp;*.ico)|*.jpg;*.bmp;*.ico|All Files (*.*)|*.*"
    SaveFile.ShowSave
    If SaveFile.CancelError Or SaveFile.FileName = "" Then Exit Sub
    
    SavePicture Picture1.Picture, SaveFile.FileName
    
    Exit Sub

Err_Trap:
    MsgBox Err.Description

End Sub
