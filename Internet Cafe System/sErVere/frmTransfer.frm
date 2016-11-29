VERSION 5.00
Begin VB.Form frmTransfer 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2985
   ClientLeft      =   3390
   ClientTop       =   2640
   ClientWidth     =   5175
   ControlBox      =   0   'False
   Icon            =   "frmTransfer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "X"
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdTransfer 
      Caption         =   "Transfer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "To Comp#:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   3240
      TabIndex        =   1
      Top             =   600
      Width           =   1695
      Begin VB.ListBox lstTo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1785
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "From Comp#:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1695
      Begin VB.ListBox lstFrom 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1785
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Label lblType 
      BackColor       =   &H00000000&
      Caption         =   " TRANSFER ACCOUNT"
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
      TabIndex        =   5
      Top             =   0
      Width           =   5175
   End
End
Attribute VB_Name = "frmTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdTransfer_Click()
Dim i As Byte
  If lstFrom.Tag = "" Or lstTo.Tag = "" Then
    MsgBox "Please Select Computer!", vbCritical, "Transfer Invalid!"
    Exit Sub
  End If

Dim tmpNum As Byte
Dim strType
Dim Rte As String
  'Transfer Process
  
  Screen.MousePointer = vbHourglass
  
  For i = 1 To frmMain.Client.UBound - 1
    If frmMain.Client(i).ComputerNumber = Val(lstFrom.Tag) Then
    
      frmMain.Client(i).Enabled = False 'Stop the computation
      'Store Data on temporary
      tmpService = frmMain.Client(i).Service
      tmpPreviousElapse = frmMain.Client(i).PreviousElapse
      tmpCurrentElapse = frmMain.Client(i).CurrentElapse
      tmpStartLog = frmMain.Client(i).StartLog
      tmpEndLog = frmMain.Client(i).EndLog
      tmpName = frmMain.Client(i).CustomerName
      tmpPreviousAmount = frmMain.Client(i).PreviousAmount
      tmpCurrentAmount = frmMain.Client(i).CurrentAmount
      tmpAmount_Limited = frmMain.Client(i).Amount_Limited
      tmpInitialStart = frmMain.Client(i).InitialStart
      tmpConnected = frmMain.Client(i).Status
      tmpAccount = frmMain.Client(i).Account
      tmpExceeded = frmMain.Client(i).Exceeded
      
      frmMain.Client(i).Reset_Var
      frmMain.Client(i).Status = v_Connected
      frmMain.Client(i).SendData "ET1" & "@"
    End If
  Next i
  
  For i = 1 To frmMain.Client.UBound - 1
    If frmMain.Client(i).ComputerNumber = Val(lstTo.Tag) Then
      
      frmMain.Client(i).Service = tmpService
      frmMain.Client(i).PreviousElapse = tmpPreviousElapse
      frmMain.Client(i).CurrentElapse = tmpCurrentElapse
      frmMain.Client(i).StartLog = tmpStartLog
      frmMain.Client(i).EndLog = tmpEndLog
      frmMain.Client(i).CustomerName = tmpName
      frmMain.Client(i).PreviousAmount = tmpPreviousAmount
      frmMain.Client(i).CurrentAmount = tmpCurrentAmount
      frmMain.Client(i).Amount_Limited = tmpAmount_Limited
      frmMain.Client(i).InitialStart = tmpInitialStart
      frmMain.Client(i).Status = tmpConnected
      frmMain.Client(i).Account = tmpAccount
      frmMain.Client(i).Exceeded = tmpExceeded
    
      Select Case tmpService
      Case v_Internet
        strType = "Internet"
        Rte = Trim(Str(INTERNET_RATE))
      Case v_Rental_Games
        strType = "Gms/Rntl"
        Rte = Trim(Str(RENTAL_RATE))
      End Select
      
      If tmpAccount = v_Open Then 'Open
        frmMain.Client(i).SendData "OP" & AddSpace(tmpName, 30) & strType & AddSpace(Rte, 2) & Trim(tmpStartLog) & "@"
      Else 'Limited
        frmMain.Client(i).SendData "LT" & AddSpace(tmpName, 30) & strType & AddSpace(Trim(Str(tmpAmount_Limited)), 3) & AddSpace(Rte, 2) & Trim(tmpStartLog) & "@"
      End If
      frmMain.Client(i).Enabled = True
    End If
  Next i
    
  'Clear Backup Data
  Remove_BackUp_Data CByte(lstFrom.Tag)
    
  'Clear Data Grid
  Remove_Grid_Data CByte(lstFrom.Tag)
  
  'Show Data in DataGrid
  tmpNum = Val(lstTo.Tag)
  
  frmMain.lvMain.ListItems(tmpNum).ListSubItems(1).Text = tmpName
  frmMain.lvMain.ListItems(tmpNum).ListSubItems(4).Text = FormatDateTime(tmpStartLog, vbLongTime)
  If tmpAccount = False Then 'Limited
    frmMain.lvMain.ListItems(tmpNum).ListSubItems(5).Text = FormatDateTime(tmpEndLog, vbLongTime)
    frmMain.lvMain.ListItems(tmpNum).ListSubItems(2).Text = "LIMITED"
    frmMain.lvMain.ListItems(tmpNum).ListSubItems(7).Text = tmpAmount_Limited
  Else 'Open
    frmMain.lvMain.ListItems(tmpNum).ListSubItems(2).Text = "OPEN"
    frmMain.lvMain.ListItems(tmpNum).ListSubItems(7).Text = tmpPreviousAmount + tmpCurrentAmount
  End If
  
  If tmpService = v_Internet Then 'Internet
    frmMain.lvMain.ListItems(tmpNum).ListSubItems(3).Text = "Internet"
  Else 'Games/Rental
    frmMain.lvMain.ListItems(tmpNum).ListSubItems(3).Text = "Gms/Rntl"
  End If
  
  'Change Light Color
  frmMain.lvMain.ListItems(Val(lstFrom.Tag)).SmallIcon = 2
  frmMain.lvMain.ListItems(Val(lstTo.Tag)).SmallIcon = 3
  
  Screen.MousePointer = vbNormal
  MsgBox "Transfer Complete"
  Unload Me
End Sub

Private Sub Form_Load()
Dim i As Byte

  Me.Icon = frmMain.Icon
  For i = 1 To frmMain.Client.UBound - 1
    'Add Data "From Computer"
    If frmMain.Client(i).Status = v_LoggedConnected Then
      lstFrom.AddItem "Comp " & frmMain.Client(i).ComputerNumber
    End If
    'Add Data "To Computer"
    If frmMain.Client(i).Status = v_Connected Then
      lstTo.AddItem "Comp " & frmMain.Client(i).ComputerNumber
    End If
  Next i
End Sub

Private Sub lblType_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  MoveForm Me
End Sub

Private Sub lstFrom_Click()
Dim i As Byte

  For i = 0 To lstFrom.ListCount - 1
    If lstFrom.Selected(i) = True Then
      lstFrom.Tag = Mid(lstFrom.List(i), 6, 2)
    End If
  Next i
End Sub

Private Sub lstTo_Click()
Dim i As Byte

  For i = 0 To lstTo.ListCount - 1
    If lstTo.Selected(i) = True Then
      lstTo.Tag = Mid(lstTo.List(i), 6, 2)
    End If
  Next i
End Sub
