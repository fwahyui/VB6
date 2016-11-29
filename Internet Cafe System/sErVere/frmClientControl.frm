VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmClientControl 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5040
   ClientLeft      =   2640
   ClientTop       =   1140
   ClientWidth     =   6735
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmClientControl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   6735
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Account Info:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   3855
      Begin VB.Timer tmrDisp 
         Interval        =   1000
         Left            =   3360
         Top             =   240
      End
      Begin VB.Label lblAmt 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "n/a"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   27
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label lblDuration 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "n/a"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   26
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label lblLogOut 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "n/a"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   25
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label lblAccount 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "n/a"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   24
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label lblService 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "n/a"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   23
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label lblLogIn 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "n/a"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   22
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "n/a"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   21
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Duration:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Log-Out Time:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Log-In Time:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Service Type:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Account Type:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Icon Legend:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   4320
      TabIndex        =   5
      Top             =   2640
      Width           =   2175
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "All PC Selected"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   13
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Image Image4 
         Height          =   255
         Left            =   240
         Picture         =   "frmClientControl.frx":000C
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Account-Logged PC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   12
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Image Image2 
         Height          =   255
         Left            =   240
         Picture         =   "frmClientControl.frx":0CD6
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Unlocked PC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   8
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Locked PC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
      Begin VB.Image Image3 
         Height          =   255
         Left            =   240
         Picture         =   "frmClientControl.frx":19A0
         Stretch         =   -1  'True
         Top             =   840
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   240
         Picture         =   "frmClientControl.frx":266A
         Stretch         =   -1  'True
         Top             =   480
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "REMOTE CONTROL:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   6255
      Begin VB.CommandButton cmdUnlock 
         Caption         =   "Unlock"
         Enabled         =   0   'False
         Height          =   855
         Left            =   3600
         Picture         =   "frmClientControl.frx":3334
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmdShut 
         Caption         =   "Shutdown"
         Enabled         =   0   'False
         Height          =   855
         Left            =   4800
         Picture         =   "frmClientControl.frx":3FFE
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmdLock 
         Caption         =   "Lock"
         Enabled         =   0   'False
         Height          =   855
         Left            =   2400
         Picture         =   "frmClientControl.frx":4CC8
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   480
         Width           =   975
      End
      Begin MSComctlLib.ImageList ImglIcons 
         Left            =   120
         Top             =   1200
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClientControl.frx":5992
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClientControl.frx":666E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClientControl.frx":734A
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClientControl.frx":8026
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageCombo ImcComp 
         Height          =   570
         Left            =   480
         TabIndex        =   4
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1005
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         Text            =   "##"
         ImageList       =   "ImglIcons"
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Computer #:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   28
         Top             =   480
         Width           =   930
      End
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblType 
      BackColor       =   &H00000000&
      Caption         =   " ADVANCE CLIENT CONTROL"
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
      TabIndex        =   2
      Top             =   0
      Width           =   6735
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "tmp"
      Visible         =   0   'False
      Begin VB.Menu mnuLock 
         Caption         =   "Lock"
      End
      Begin VB.Menu mnuUnlock 
         Caption         =   "Unlock"
      End
      Begin VB.Menu mnuSpace2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSpace 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShut 
         Caption         =   "Shutdown"
      End
   End
End
Attribute VB_Name = "frmClientControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Ctr As Byte

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdLock_Click()
Dim idx As Byte
Dim i, j, k As Byte
  
  If ImcComp.Text = "All" Then
    k = ImcComp.ComboItems.Count - 1 'lock all computers
  Else
    k = 1 'lock selected computer
  End If


  For j = 1 To k
    If ImcComp.Text = "All" Then
      idx = CByte(ImcComp.ComboItems(j).Tag)
    Else
      idx = CByte(ImcComp.SelectedItem.Tag)
    End If
        
    If frmMain.Client(idx).Status = v_UnlockedConnected Then
      frmMain.Client(idx).SendData "ET1" & "@"
      frmMain.Client(idx).Status = v_Connected
      frmMain.Client(idx).EndLog = Now
      frmMain.lvMain.ListItems(frmMain.Client(idx).ComputerNumber).SmallIcon = 2
      ImcComp.SelectedItem.Image = 1
      
      'Save in ClientMonitor's Dbase
      Mon_Rst.AddNew
      Mon_Rst!Month = Format(frmMain.Client(idx).StartLog, "m")
      Mon_Rst!Day = Format(frmMain.Client(idx).StartLog, "d")
      Mon_Rst!Year = Format(frmMain.Client(idx).StartLog, "yyyy")
      Mon_Rst!CN = frmMain.Client(idx).ComputerNumber
      Mon_Rst!UnlockTime = FormatDateTime(frmMain.Client(idx).StartLog, vbLongTime)
      Mon_Rst!LockTime = FormatDateTime(frmMain.Client(idx).EndLog, vbLongTime)
      Mon_Rst!Duration = DateDiff("n", frmMain.Client(idx).StartLog, frmMain.Client(idx).EndLog)
      Mon_Rst.Update
      
      'Delete BackUp Entry
      SvrDbRst.MoveFirst
      SvrDbRst.Find "ComNum LIKE " & frmMain.Client(idx).ComputerNumber, 1, adSearchForward
      SvrDbRst!StartLog = Null
      SvrDbRst!Amt = 0
      SvrDbRst!Account = 0
      SvrDbRst!Service = 0
      SvrDbRst!Unlock = False
      SvrDbRst.Update
      
      'Clear Entry in DataGrid
      Remove_Grid_Data CByte(frmMain.Client(idx).ComputerNumber)
    End If
  Next j
  
  If ImcComp.Text = "All" Then
    Form_Load
  Else
    ImcComp_Click
  End If
End Sub

Private Sub cmdRefresh_Click()
  Form_Load
End Sub

Private Sub cmdShut_Click()
Dim j, k As Byte
Dim idx As Byte

  
  If ImcComp.Text = "All" Then
    k = ImcComp.ComboItems.Count - 1 'Shutdown all computers
  Else
    k = 1 'Shutdown selected computer
  End If

  For j = 1 To k
    If ImcComp.Text = "All" Then
      idx = CByte(ImcComp.ComboItems(j).Tag)
    Else
      idx = CByte(ImcComp.SelectedItem.Tag)
    End If
        
    'Filters not to shutdown the currently log account
    If frmMain.Client(idx).Status <> v_LoggedConnected Then
      frmMain.Client(idx).SendData "SD" & "@"
      frmMain.lvMain.ListItems(frmMain.Client(idx).ComputerNumber).SmallIcon = 1
      frmMain.Client(idx).Status = v_NotConnected
      frmMain.Client(idx).ComputerNumber = 0
    End If
  Next j
  Form_Load
  cmdShut.Enabled = False
  cmdLock.Enabled = False
  cmdUnlock.Enabled = False
End Sub

Private Sub cmdUnlock_Click()
Dim i, j, k As Byte
Dim idx As Byte


  If ImcComp.Text = "All" Then
    k = ImcComp.ComboItems.Count - 1 'unlock all computers
  Else
    k = 1 'unlock selected computer
  End If

  For j = 1 To k
    If ImcComp.Text = "All" Then
      idx = CByte(ImcComp.ComboItems(j).Tag)
    Else
      idx = CByte(ImcComp.SelectedItem.Tag)
    End If
    
    'Filters to unlock only the locked computer
    If frmMain.Client(idx).Status = v_Connected Then
      frmMain.Client(idx).SendData "ET2" & "@"
      frmMain.Client(idx).Status = v_UnlockedConnected
      frmMain.Client(idx).StartLog = Now
      frmMain.lvMain.ListItems(frmMain.Client(idx).ComputerNumber).SmallIcon = 3
      ImcComp.SelectedItem.Image = 2
      
      'Fill Entry in DataGrid
      For i = 1 To 7
        frmMain.lvMain.ListItems(frmMain.Client(idx).ComputerNumber).ListSubItems(i) = "Unlocked"
      Next i
    End If
  Next j
    
  If ImcComp.Text = "All" Then
    Form_Load
  Else
    ImcComp_Click
  End If
End Sub

Private Sub Form_Activate()
  If Ctr = 1 Then
    MsgBox "No current connections available!", vbExclamation
    Unload Me
  End If
End Sub

Private Sub Form_Load()
Dim i As Byte
Dim j As Byte
  Me.Icon = frmMain.Icon
  ImcComp.ComboItems.Clear
  ImcComp.Text = "##"
  
  Ctr = 1
  For i = 1 To NumberComps
    j = 0
    Do Until j = frmMain.Client.UBound - 1
      j = j + 1
      If frmMain.Client(j).ComputerNumber = i And frmMain.Client(j).Status <> v_NotConnected Then
        Select Case frmMain.Client(j).Status
        Case v_Connected 'Locked
          ImcComp.ComboItems.Add , , i, 1
        Case v_LoggedConnected 'logged
          ImcComp.ComboItems.Add , , i, 3
        Case v_UnlockedConnected 'Unlocked
          ImcComp.ComboItems.Add , , i, 2
        End Select
        ImcComp.ComboItems(Ctr).Tag = j
        Ctr = Ctr + 1
        Exit Do
      End If
    Loop
  Next i
  
  If Ctr <> 1 Then ImcComp.ComboItems.Add , , "All", 4
    
End Sub

Private Sub ImcComp_Click()
Dim idx As Byte
  If ImcComp.Text = "All" Then
    cmd_Enable True, True, True
    Exit Sub
  End If
  
  idx = Val(ImcComp.SelectedItem.Tag)
  Select Case frmMain.Client(idx).Status
  Case v_Connected 'Locked
    cmd_Enable False, True, True
  Case v_LoggedConnected 'Account Logged
      cmd_Enable False, False, False
  Case v_UnlockedConnected 'Unlocked
    cmd_Enable True, False, True
  End Select
End Sub

Private Sub cmd_Enable(cLock As Boolean, cUnlock As Boolean, cShut As Boolean)
  cmdLock.Enabled = cLock
  cmdUnlock.Enabled = cUnlock
  cmdShut.Enabled = cShut
End Sub


Private Sub lblType_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  MoveForm Me
End Sub

Private Sub tmrDisp_Timer()
Dim idx As Byte
On Error GoTo AleX

  idx = ImcComp.SelectedItem.Tag
  If frmMain.Client(idx).Status = v_LoggedConnected Then
    lblName.Caption = frmMain.Client(idx).CustomerName
    If frmMain.Client(idx).Account = v_Open Then 'Open
      lblAccount.Caption = "OPEN"
    Else 'limited
      lblAccount.Caption = "LIMITED"
    End If
    If frmMain.Client(idx).Service = v_Internet Then
      lblService.Caption = "Internet"
    Else
      lblService.Caption = "Games/Rental"
    End If
    lblLogIn.Caption = frmMain.Client(idx).StartLog
    lblLogOut.Caption = frmMain.Client(idx).EndLog
    lblDuration.Caption = Formatter(Val(frmMain.Client(idx).PreviousElapse + frmMain.Client(idx).CurrentElapse))
    lblAmt.Caption = FormatNumber(frmMain.Client(idx).PreviousAmount + frmMain.Client(idx).CurrentAmount, 2)
  Else
AleX:
    lblName.Caption = "n/a"
    lblAccount.Caption = "n/a"
    lblService.Caption = "n/a"
    lblLogIn.Caption = "n/a"
    lblLogOut.Caption = "n/a"
    lblDuration.Caption = "n/a"
    lblAmt.Caption = "n/a"
  End If
End Sub
