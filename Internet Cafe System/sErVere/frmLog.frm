VERSION 5.00
Begin VB.Form frmLog 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3195
   ClientLeft      =   3390
   ClientTop       =   2640
   ClientWidth     =   4695
   ControlBox      =   0   'False
   Icon            =   "frmLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4695
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   3255
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4695
      Begin VB.TextBox txtName 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   1
         Top             =   1080
         Width           =   2175
      End
      Begin VB.CommandButton cmdOk 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   1920
         TabIndex        =   4
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "C&ancel"
         Height          =   375
         Left            =   3240
         TabIndex        =   5
         Top             =   2640
         Width           =   1215
      End
      Begin VB.ComboBox cmbCompNum 
         BackColor       =   &H00E0E0E0&
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
         TabIndex        =   0
         Top             =   600
         Width           =   2175
      End
      Begin VB.ComboBox cmbAmt 
         BackColor       =   &H00E0E0E0&
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
         ItemData        =   "frmLog.frx":000C
         Left            =   1800
         List            =   "frmLog.frx":004C
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2040
         Width           =   2175
      End
      Begin VB.ComboBox cmbType 
         BackColor       =   &H00E0E0E0&
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
         ItemData        =   "frmLog.frx":00A0
         Left            =   1800
         List            =   "frmLog.frx":00AA
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "(NAME):"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   1155
         Width           =   1335
      End
      Begin VB.Label lblType 
         BackColor       =   &H00000000&
         Caption         =   " OPEN ACCOUNT"
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
         TabIndex        =   10
         Top             =   0
         Width           =   4695
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
         TabIndex        =   9
         Top             =   680
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
         TabIndex        =   8
         Top             =   2100
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Service Type:"
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
         TabIndex        =   7
         Top             =   1600
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private numC As Byte

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOk_Click()
Dim i As Byte
Dim strType As String
Dim Rte As String
Dim temp As Date
  
  Select Case True
  Case cmbCompNum.Text = ""
    MsgBox "Invalid Computer No.! "
    cmbCompNum.SetFocus
    Exit Sub
  Case Me.cmbType.Text = ""
    MsgBox "Invalid Service Type!"
    Me.cmbType.SetFocus
    Exit Sub
  Case Me.cmbAmt.Text = ""
    If TypeTrans = False Then
      MsgBox "Invalid Amount!"
      Me.cmbAmt.SetFocus
      Exit Sub
    End If
  End Select
  
  For i = 1 To NumberComps
    If frmMain.Client(i).ComputerNumber = Val(cmbCompNum.Text) Then
      temp = Now
              
      frmMain.Client(i).Status = v_LoggedConnected
      frmMain.Client(i).StartLog = temp
      frmMain.Client(i).InitialStart = temp
      frmMain.Client(i).CustomerName = Trim(txtName)
      
      Select Case cmbType.Text
      Case "Internet"
        frmMain.Client(i).Service = v_Internet
        strType = "Internet"
        Rte = Trim(Str(INTERNET_RATE))
      Case "Games/Rental"
        frmMain.Client(i).Service = v_Rental_Games
        strType = "Gms/Rntl"
        Rte = Trim(Str(RENTAL_RATE))
      End Select
      
      If TypeTrans = True Then
        'OPEN
        frmMain.Client(i).Account = v_Open
        frmMain.Client(i).SendData "OP" & AddSpace(frmMain.Client(i).CustomerName, 30) & strType & AddSpace(Rte, 2) & temp & "@"
        frmMain.lvMain.ListItems(frmMain.Client(i).ComputerNumber).ListSubItems(2).Text = "OPEN"
      Else
        'LIMITED
        frmMain.Client(i).Account = v_Limited
        frmMain.Client(i).Amount_Limited = Val(cmbAmt.Text)
        frmMain.Client(i).SendData "LT" & AddSpace(Trim(txtName.Text), 30) & strType & AddSpace(Trim(cmbAmt.Text), 3) & AddSpace(Rte, 2) & temp & "@"
        frmMain.lvMain.ListItems(frmMain.Client(i).ComputerNumber).ListSubItems(2).Text = "LIMITED"
        frmMain.lvMain.ListItems(frmMain.Client(i).ComputerNumber).ListSubItems(7).Text = FormatNumber(frmMain.Client(i).Amount_Limited, 2)
        Select Case frmMain.Client(i).Service
        Case v_Internet
          frmMain.Client(i).EndLog = DateAdd("n", frmMain.Client(i).Amount_Limited * (60 / INTERNET_RATE), frmMain.Client(i).InitialStart)
        Case v_Rental_Games
          frmMain.Client(i).EndLog = DateAdd("n", frmMain.Client(i).Amount_Limited * (60 / RENTAL_RATE), frmMain.Client(i).InitialStart)
        End Select
        frmMain.lvMain.ListItems(frmMain.Client(i).ComputerNumber).ListSubItems(5).Text = FormatDateTime(frmMain.Client(i).EndLog, vbLongTime)
      End If
             
      frmMain.lvMain.ListItems(frmMain.Client(i).ComputerNumber).ListSubItems(3).Text = strType
      frmMain.lvMain.ListItems(frmMain.Client(i).ComputerNumber).SmallIcon = 3
      frmMain.lvMain.ListItems(frmMain.Client(i).ComputerNumber).ListSubItems(1).Text = txtName.Text
      frmMain.lvMain.ListItems(frmMain.Client(i).ComputerNumber).ListSubItems(4).Text = FormatDateTime(temp, vbLongTime)
      frmMain.Client(i).Enabled = True
      Exit For
    End If
  Next i

  Unload Me
End Sub

Private Sub Form_Activate()
  If numC = 0 Then Unload Me
End Sub

Private Sub Form_Load()
Dim i As Byte
  
  Me.Icon = frmMain.Icon
  numC = 0
  For i = 1 To frmMain.Client.UBound - 1
    If frmMain.Client(i).Status = v_Connected And frmMain.Client(i).Exceeded = False Then
      cmbCompNum.AddItem frmMain.Client(i).ComputerNumber
      numC = numC + 1
    End If
  Next i
  
  If numC = 0 Then
    MsgBox "Ambigious to Log!" & vbCrLf & "No client connection available."
    Exit Sub
  End If
  
  If TypeTrans = True Then
    lblType.Caption = " OPEN ACCOUNT"
    lblType.ForeColor = vbGreen
    Label2.Visible = False
    cmbAmt.Visible = False
  Else
    lblType.Caption = " LIMITED ACCOUNT"
    lblType.ForeColor = vbYellow
  End If
  
  ExplodeForm Me, 500
End Sub

Private Sub Form_Unload(Cancel As Integer)
  ImplodeForm Me, 2, 500, 1
End Sub

Private Sub lblType_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  MoveForm Me
End Sub
