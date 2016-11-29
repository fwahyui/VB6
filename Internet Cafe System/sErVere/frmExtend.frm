VERSION 5.00
Begin VB.Form frmExtend 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2535
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4245
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   4245
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbAmt 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "frmExtend.frx":0000
      Left            =   1800
      List            =   "frmExtend.frx":0040
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1200
      Width           =   1935
   End
   Begin VB.ComboBox cmbCompNum 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "C&ancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Appearance      =   0  'Flat
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount:     P"
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
      Left            =   480
      TabIndex        =   6
      Top             =   1275
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Computer #:"
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
      Left            =   480
      TabIndex        =   1
      Top             =   810
      Width           =   1215
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00000000&
      Caption         =   " ACCOUNT EXTENSION"
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
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmExtend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOk_Click()
Dim i As Byte
Dim j As Byte
  'check the fields
  If cmbCompNum.Text = "" Or cmbAmt.Text = "" Then
    MsgBox "Pls complete the fields!", vbExclamation
    Exit Sub
  End If
  
  For i = 1 To frmMain.Client.UBound - 1
    If frmMain.Client(i).ComputerNumber = Val(cmbCompNum.Text) Then
      frmMain.Client(i).Amount_Limited = frmMain.Client(i).Amount_Limited + Val(cmbAmt.Text)
      frmMain.Client(i).Status = v_LoggedConnected
      frmMain.Client(i).Exceeded = False
      frmMain.Client(i).InitialStart = Now
      Select Case frmMain.Client(i).Service
      Case v_Internet 'internet
        frmMain.Client(i).SendData "LT" & AddSpace(frmMain.Client(i).CustomerName, 30) & "Internet" & AddSpace(CStr(frmMain.Client(i).Amount_Limited), 3) & AddSpace(Str(INTERNET_RATE), 2) & frmMain.Client(i).StartLog & "@"
      Case v_Rental_Games 'Games/Rental
        frmMain.Client(i).SendData "LT" & AddSpace(frmMain.Client(i).CustomerName, 30) & "Gms/Rntl" & AddSpace(CStr(frmMain.Client(i).Amount_Limited), 3) & AddSpace(Str(RENTAL_RATE), 2) & frmMain.Client(i).StartLog & "@"
      End Select
      
      frmMain.lvMain.ListItems(frmMain.Client(i).ComputerNumber).SmallIcon = 3
      frmMain.lvMain.ListItems(frmMain.Client(i).ComputerNumber).ListSubItems(7).Text = FormatNumber(frmMain.Client(i).Amount_Limited, 2)
      
      Select Case frmMain.Client(i).Service
      Case v_Internet 'Internet
        frmMain.Client(i).EndLog = DateAdd("n", (frmMain.Client(i).Amount_Limited - frmMain.Client(i).PreviousAmount) * (60 / INTERNET_RATE), frmMain.Client(i).InitialStart)
      Case v_Rental_Games 'Rental
        frmMain.Client(i).EndLog = DateAdd("n", (frmMain.Client(i).Amount_Limited - frmMain.Client(i).PreviousAmount) * (60 / RENTAL_RATE), frmMain.Client(i).InitialStart)
      End Select
      frmMain.lvMain.ListItems(frmMain.Client(i).ComputerNumber).ListSubItems(5).Text = frmMain.Client(i).EndLog
            
      'Delete Entry in Listbox
      For j = 1 To frmMain.lstLoggedOut.ListCount
        If Left(frmMain.lstLoggedOut.List(j - 1), Len("Comp #" & frmMain.Client(i).ComputerNumber)) = "Comp #" & frmMain.Client(i).ComputerNumber Then
          frmMain.lstLoggedOut.RemoveItem j - 1
          Exit For
        End If
      Next j
      
      frmMain.Check_LstLogged 'Stop Blink Effect and Change Color
      
      Exit For
    End If
  Next i
    
  Unload Me
  MsgBox "Account Extension Completed!"
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim i As Byte
  
  For i = 1 To NumberComps
    If frmMain.Client(i).Status = v_Connected And frmMain.Client(i).Exceeded = True And frmMain.Client(i).Account = v_Limited Then
      cmbCompNum.AddItem frmMain.Client(i).ComputerNumber
    End If
  Next i
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  MoveForm Me
End Sub
