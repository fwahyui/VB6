VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl ComputerClient 
   ClientHeight    =   2100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1050
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   2100
   ScaleWidth      =   1050
   ToolboxBitmap   =   "ctlClient.ctx":0000
   Begin MSWinsockLib.Winsock Socket 
      Left            =   0
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   3982
   End
   Begin VB.Timer tmrTimer 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   0
      Top             =   840
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   450
      Left            =   0
      Picture         =   "ctlClient.ctx":0312
      Stretch         =   -1  'True
      Top             =   0
      Width           =   420
   End
End
Attribute VB_Name = "ComputerClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum Comp_Status
  v_NotConnected
  v_Connected
  v_LoggedConnected
  v_UnlockedConnected
  v_PausedLogged
End Enum

Public Enum TypeAccount
  v_Open
  v_Limited
End Enum

Public Enum TypeService
  v_Internet
  v_Rental_Games
End Enum

Public Service As TypeService
Public PreviousElapse As Long  'in Minutes
Public CurrentElapse As Long
Public StartLog As String
Public EndLog As String
Public CustomerName As String
Public PreviousAmount As Integer 'Previous Amount
Public CurrentAmount As Integer 'Initial amount
Public Amount_Limited As Integer 'Amount for limited account only
Public InitialStart As String 'Initial Start Time Log
Public ComputerNumber As Integer
Public Status As Comp_Status
Public Account As TypeAccount
Public Exceeded As Boolean

Private Sub Socket_Close()
On Error Resume Next
  'Socket.Close
  UserControl.Enabled = False
  Status = v_NotConnected
  frmMain.lvMain.ListItems(ComputerNumber).SmallIcon = 1
  ComputerNumber = 0
  Reset_Var
End Sub

Private Sub Socket_Connect()
  SynchronizeFlash
End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
Dim X As String
Dim Num As Integer
Dim i As Byte
Socket.GetData X
  
  'Header Commands
  'ID - computer number
  'LO - log out
  'UL - Unlock Signal
  'LK - Lock signal
  
  Select Case Left(X, 2)
  Case "ID"
    Num = Val(Mid(X, 3, Len(X) - 2))
    'check if ID# is in Server allow range
    If Num > NumberComps Then
      Socket.SendData "ER3" & NumberComps & "@"
      Exit Sub
    End If
    'check if ID# is used
    Dim iTmp As Byte
    For i = 1 To frmMain.Client.UBound - 1
      If frmMain.Client(i).ComputerNumber = Num Then
        Socket.SendData "ER1" & "@"
        Exit Sub
      End If
      If frmMain.Client(i).Status = v_Connected Then iTmp = iTmp + 1
    Next i
    'check for maximum connection
    If iTmp > NumberComps Then
      Socket.SendData "ER2" & "@"
      Exit Sub
    End If
    
    ComputerNumber = Num
    Status = v_Connected
    frmMain.lvMain.ListItems(Num).SmallIcon = 2
    Socket.SendData "SC@" 'Success Connection
  Case "LO"
    Exceeded = True
    Status = v_Connected
    EndLog = Now
    frmMain.lvMain.ListItems(ComputerNumber).ListSubItems(5).Text = FormatDateTime(EndLog, vbLongTime)
    If PreviousAmount + CurrentAmount < MIN_AMT Then
      PreviousAmount = MIN_AMT
      CurrentAmount = 0
    End If
    frmMain.lstLoggedOut.AddItem "Comp #" & Trim(Str(ComputerNumber)) & "  -----  P " & FormatNumber(PreviousAmount + CurrentAmount, 2)
    frmMain.lvMain.ListItems(ComputerNumber).SmallIcon = 2
    frmMain.tmrLogBlinker.Enabled = True
  Case "UL"
    Status = v_UnlockedConnected
    If StartLog = "" And EndLog = "" Then
      StartLog = Now
      frmMain.lvMain.ListItems(ComputerNumber).SmallIcon = 3
      
      'Fill Entry in DataGrid
      For i = 1 To 7
        frmMain.lvMain.ListItems(ComputerNumber).ListSubItems(i).Text = "Unlocked"
      Next i
    End If
  Case "LK"
    Status = v_Connected
    If EndLog = "" And StartLog <> "" Then
      EndLog = Now
      frmMain.lvMain.ListItems(ComputerNumber).SmallIcon = 2
      
      'Save in ClientMonitor's Dbase
      Mon_Rst.AddNew
      Mon_Rst!Month = Format(StartLog, "m")
      Mon_Rst!Day = Format(StartLog, "d")
      Mon_Rst!Year = Format(StartLog, "yyyy")
      Mon_Rst!CN = ComputerNumber
      Mon_Rst!UnlockTime = FormatDateTime(StartLog, vbLongTime)
      Mon_Rst!LockTime = FormatDateTime(EndLog, vbLongTime)
      Mon_Rst!Duration = DateDiff("n", StartLog, EndLog)
      Mon_Rst.Update
      
      'Delete BackUp Entry
      SvrDbRst.MoveFirst
      SvrDbRst.Find "ComNum LIKE " & ComputerNumber, 1, adSearchForward
      SvrDbRst!StartLog = Null
      
      SvrDbRst!Amt = 0
      SvrDbRst!Account = 0
      SvrDbRst!Service = 0
      SvrDbRst!Unlock = False
      SvrDbRst.Update
      
      'Clear Time Var
      StartLog = ""
      EndLog = ""
      
      'Clear Grid's data
      Remove_Grid_Data CByte(ComputerNumber)
    End If
  End Select
End Sub

Private Sub tmrTimer_Timer()
On Error Resume Next
  Compute
  
  'display elapse, amount, and update client
  If Status = v_LoggedConnected And Exceeded = False Then
    frmMain.lvMain.ListItems(ComputerNumber).ListSubItems(6).Text = Formatter(PreviousElapse + CurrentElapse)
    
    Dim tmpN As Integer
    Select Case Service
    Case v_Internet
      tmpN = INTERNET_RATE
    Case v_Rental_Games
      tmpN = RENTAL_RATE
    End Select
    
    Select Case Account
    Case v_Open 'OPEN
      Socket.SendData "T1" & AddSpace(CStr(PreviousElapse + CurrentElapse), 4) & CStr(PreviousAmount + CurrentAmount) & "@"
      frmMain.lvMain.ListItems(ComputerNumber).ListSubItems(7).Text = FormatNumber(PreviousAmount + CurrentAmount, 2)
    Case v_Limited 'LIMITED
      tmpN = ((PreviousElapse + CurrentElapse) * tmpN) / 60
      Socket.SendData "T2" & AddSpace(CStr(PreviousElapse + CurrentElapse), 4) & AddSpace(Trim(Val(tmpN)), 3) & "@"
    End Select
  End If
    
  Save_Data
  
  'check if exceeded
  If Account = v_Limited And Status = v_LoggedConnected Then
    Dim tmpMustElapse As Long
    Select Case Service
    Case v_Internet
      tmpMustElapse = Amount_Limited / (INTERNET_RATE / 60)
    Case v_Rental_Games
      tmpMustElapse = Amount_Limited / (RENTAL_RATE / 60)
    End Select
    If (PreviousElapse + CurrentElapse) >= tmpMustElapse Then
      UserControl.Enabled = False
      Socket.SendData "ET" & "@"
      Status = v_Connected
      Exceeded = True
      EndLog = Now
      frmMain.lvMain.ListItems(ComputerNumber).ListSubItems(5).Text = FormatDateTime(EndLog, vbLongTime)
      frmMain.lstLoggedOut.AddItem "Comp #" & Trim(Str(ComputerNumber)) & "  -----  P " & FormatNumber(PreviousAmount + CurrentAmount, 2)
      frmMain.lvMain.ListItems(ComputerNumber).SmallIcon = 2
      frmMain.tmrLogBlinker.Enabled = True
    End If
  End If
End Sub

Private Sub UserControl_Initialize()
  Reset_Var
End Sub

Private Sub UserControl_Resize()
  UserControl.Width = Image1.Width
  UserControl.Height = Image1.Height
End Sub

Public Property Let Enabled(ByVal tmpValue As Boolean)
  tmrTimer.Enabled = tmpValue
  PropertyChanged "Enabled"
  
  If tmpValue = False Then
    PreviousAmount = PreviousAmount + CurrentAmount
    PreviousElapse = PreviousElapse + CurrentElapse
    CurrentAmount = 0
    CurrentElapse = 0
  End If
End Property

Public Property Get Enabled() As Boolean
  Enabled = tmrTimer.Enabled
End Property

Public Sub Accept(requestID As Long)
  Socket.Accept requestID
End Sub

Public Sub Cloze()
  Socket.Close
  UserControl.Enabled = False
End Sub

Public Sub SendData(Datos As String)
  Socket.SendData Datos
End Sub

Private Sub Compute()
  'compute initial elapse
  CurrentElapse = DateDiff("n", InitialStart, Now)
  'compute initial Amount and Total elapse
  Select Case Service
  Case v_Internet
    CurrentAmount = INTERNET_RATE / 60 * CurrentElapse
  Case v_Rental_Games
    CurrentAmount = RENTAL_RATE / 60 * CurrentElapse
  End Select
End Sub

Public Sub Reset_Var()
  PreviousElapse = 0
  CurrentElapse = 0
  StartLog = ""
  EndLog = ""
  CustomerName = ""
  PreviousAmount = 0
  CurrentAmount = 0
  Amount_Limited = 0
  InitialStart = ""
  Exceeded = False
End Sub

Private Sub UserControl_Terminate()
  Socket.Close
  tmrTimer.Enabled = False
End Sub

Private Sub Save_Data()
  SvrDbRst.MoveFirst
  SvrDbRst.Find "ComNum LIKE " & ComputerNumber, 1, adSearchForward
  With SvrDbRst
    !Name = Format(CustomerName, "")
    !StartLog = StartLog
    !Account = Account
    If Status = v_UnlockedConnected Then !Unlock = True
    !AmtLimited = Amount_Limited
    !Amt = PreviousAmount + CurrentAmount
    !Elapse = PreviousElapse + CurrentElapse
    !Service = Service
    .Update
  End With
End Sub

Public Property Get Sock_State() As Byte
  Sock_State = Socket.State
End Property

