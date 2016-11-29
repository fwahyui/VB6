VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "PC Empire Internet Cafe System  - Server 1.3"
   ClientHeight    =   9480
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   11400
   FillStyle       =   0  'Solid
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9480
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin SERVERE.TrayArea Tray 
      Left            =   5760
      Top             =   2040
      _ExtentX        =   900
      _ExtentY        =   900
      ToolTip         =   "Internet Cafe System"
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "PC Status Legend:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   9480
      TabIndex        =   15
      Top             =   960
      Width           =   2295
      Begin VB.Image Image4 
         Height          =   480
         Left            =   240
         Picture         =   "frmMain.frx":08CA
         Stretch         =   -1  'True
         Top             =   720
         Width           =   480
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   240
         Picture         =   "frmMain.frx":1794
         Stretch         =   -1  'True
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   240
         Picture         =   "frmMain.frx":265E
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "frmMain.frx":3528
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   480
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Paused Acnt PC"
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
         Left            =   840
         TabIndex        =   19
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "UnLocked PC"
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
         Left            =   840
         TabIndex        =   18
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label label12 
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
         Left            =   840
         TabIndex        =   17
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Disconnected PC"
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
         Left            =   840
         TabIndex        =   16
         Top             =   1320
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "Service Rates:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   6000
      TabIndex        =   10
      Top             =   960
      Width           =   3375
      Begin VB.Label lblGmsRntl 
         BackStyle       =   0  'Transparent
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
         Left            =   1680
         TabIndex        =   14
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblInternet 
         BackStyle       =   0  'Transparent
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
         Left            =   1680
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Games/Rental:   Rp"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Internet:              Rp"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSComctlLib.Toolbar tBar 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   1482
      ButtonWidth     =   1323
      ButtonHeight    =   1429
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgl"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open"
            Key             =   "Open"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Limited"
            Key             =   "Limited"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Transfer"
            Key             =   "Transfer"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Extend"
            Key             =   "Extend"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            Key             =   "Help"
            ImageIndex      =   7
         EndProperty
      EndProperty
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   4080
         ScaleHeight     =   735
         ScaleWidth      =   7695
         TabIndex        =   5
         Top             =   0
         Width           =   7695
         Begin VB.CommandButton Command1 
            Caption         =   "RESET"
            Height          =   255
            Left            =   4440
            TabIndex        =   9
            Top             =   240
            Width           =   855
         End
         Begin VB.Timer tmrDateTime 
            Interval        =   250
            Left            =   840
            Top             =   360
         End
         Begin VB.CheckBox chkMute 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Mute"
            Height          =   255
            Left            =   5880
            TabIndex        =   6
            Top             =   480
            Width           =   700
         End
         Begin MSComctlLib.Slider sldVolMaster 
            Height          =   615
            Left            =   6480
            TabIndex        =   7
            Top             =   0
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   1085
            _Version        =   393216
            Max             =   100
            SelStart        =   100
            TickStyle       =   2
            TickFrequency   =   10
            Value           =   100
         End
         Begin VB.Image imgSound 
            Height          =   360
            Left            =   6000
            Stretch         =   -1  'True
            Top             =   0
            Width           =   375
         End
         Begin VB.Label lblDate 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "September 3, 1982"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   8
            Top             =   480
            Width           =   4095
         End
      End
   End
   Begin MSComctlLib.ImageList imgl 
      Left            =   4800
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":43F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4CCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":55AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":58C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":61A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":64BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":67DA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrLogBlinker 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5280
      Top             =   1560
   End
   Begin VB.ListBox lstLoggedOut 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1710
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   4335
   End
   Begin VB.Timer tmrResume 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   4800
      Top             =   2040
   End
   Begin VB.Timer tmrCloser 
      Interval        =   1000
      Left            =   5280
      Top             =   2040
   End
   Begin MSWinsockLib.Winsock Listener 
      Left            =   4800
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   3982
   End
   Begin MSComctlLib.ListView lvMain 
      Height          =   4335
      Left            =   240
      TabIndex        =   3
      Top             =   3600
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   7646
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imglCol"
      ForeColor       =   -2147483640
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Comp#"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Account Type"
         Object.Width           =   2364
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Service Type"
         Object.Width           =   2364
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Log In Time"
         Object.Width           =   2364
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Log Out Time"
         Object.Width           =   2364
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Duration"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "AMOUNT  "
         Object.Width           =   2011
      EndProperty
   End
   Begin MSComctlLib.ImageList imglCol 
      Left            =   5400
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6AF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":79D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":88AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":978A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A666
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B542
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin SERVERE.ComputerClient Client 
      Index           =   1
      Left            =   5760
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   794
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "  REMOTE STATUS:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   3360
      Width           =   11535
   End
   Begin VB.Label lblLog 
      BackStyle       =   0  'Transparent
      Caption         =   "  LOGGED OUT:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   4335
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   240
      Top             =   1080
      Width           =   4335
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New Account"
         Begin VB.Menu mnuOpenA 
            Caption         =   "&Open Account"
            Shortcut        =   {F5}
         End
         Begin VB.Menu mnuLimitedA 
            Caption         =   "&Limited Account"
            Shortcut        =   {F8}
         End
      End
      Begin VB.Menu mnudash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTransfer 
         Caption         =   "Trans&fer Account"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuExtend 
         Caption         =   "&Extend Account"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnudash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuClientLogs 
         Caption         =   "Account &Logs"
      End
      Begin VB.Menu mnuLock_Unlock 
         Caption         =   "Lock/Unlock Logs"
      End
   End
   Begin VB.Menu mnuAdv 
      Caption         =   "&Tools"
      Begin VB.Menu mnuOptions 
         Caption         =   "Config / &Options"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuClientControl 
         Caption         =   "Advance &Client Control"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuKwh 
         Caption         =   "Kwh Consumption"
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnuDaily 
         Caption         =   "Daily Report"
      End
      Begin VB.Menu mnuMonthly 
         Caption         =   "Monthly Report"
      End
   End
   Begin VB.Menu mnuGrid 
      Caption         =   "GridMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuComp 
         Caption         =   "Computer"
         Begin VB.Menu mnuPause 
            Caption         =   "Pause"
         End
         Begin VB.Menu mnuContinue 
            Caption         =   "Continue"
         End
         Begin VB.Menu mnuExtension 
            Caption         =   "Extend"
         End
         Begin VB.Menu mnudash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuLogIn 
            Caption         =   "Log In"
            Begin VB.Menu mnuOpen 
               Caption         =   "Open"
            End
            Begin VB.Menu mnuLimited 
               Caption         =   "Limited"
            End
         End
         Begin VB.Menu mnuLogOut 
            Caption         =   "Log Out"
         End
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Mark As Boolean

Private Sub chkMute_Click()

On Error Resume Next
  Select Case chkMute.Value
  Case vbChecked
    imgSound.Picture = imgl.ListImages(6).Picture
    Mute_UnMuteAll "MT1@"
  Case vbUnchecked
    imgSound.Picture = imgl.ListImages(5).Picture
    Mute_UnMuteAll "MT2@"
  End Select
End Sub

Private Sub Mute_UnMuteAll(Data As String)
Dim i As Byte
  For i = 1 To Client.UBound - 1
      Client(i).SendData Data
  Next i
End Sub

Public Sub Fill_ListView()
Dim i As Integer
Dim Item As Object

  lvMain.ListItems.Clear
  For i = 1 To NumberComps
    Set Item = lvMain.ListItems.Add(, , " " & i)
    
    Item.ListSubItems.Add , , "" 'Name
    Item.ListSubItems.Add , , "" 'Account Type
    Item.ListSubItems.Add , , "" 'Service Type
    Item.ListSubItems.Add , , "" 'Log In Time
    Item.ListSubItems.Add , , "" 'Log Out Time
    Item.ListSubItems.Add , , "" 'Duration
    Item.ListSubItems.Add , , "" 'Amount
    
    lvMain.ListItems(i).SmallIcon = 1
  Next i
  
  For i = 1 To NumberComps
    lvMain.ListItems(i).ListSubItems(7).Bold = True
  Next i
End Sub

Private Sub Command1_Click()
Dim A As Double
  A = Shell(App.Path & "\" & App.EXEName)
  End
End Sub

Private Sub Form_Load()
  'DoEvents
  'SkinForm1.LoadSkin Me
  Tray.Icon = imglCol.ListImages(5).Picture
  Me.Caption = "PC Empire Internet Cafe System  - Server " & App.Major & "." & App.Minor
  lblInternet.Caption = FormatNumber(INTERNET_RATE, 2)
  lblGmsRntl.Caption = FormatNumber(RENTAL_RATE, 2)
  If chkMute.Value = vbChecked Then
    imgSound.Picture = imgl.ListImages(6).Picture
  Else
    imgSound.Picture = imgl.ListImages(5).Picture
  End If
  Fill_ListView
  Me.Show
  Listener.Listen
End Sub

Private Sub Form_Resize()
  If Me.WindowState = vbMinimized Then
    frmMain.Visible = False
    Tray.Visible = True
  Else
    Tray.Visible = False
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If MsgBox("Are you sure you want to exit?", vbCritical + vbYesNo, "Confirm!") = vbNo Then
    Cancel = 1
  Else
    End
  End If
End Sub

Private Sub Listener_ConnectionRequest(ByVal requestID As Long)
Dim i As Byte
  For i = 1 To Client.UBound
    If Client(i).Sock_State = sckClosed Then
      Client(i).Accept requestID
      Exit For
    End If
  Next i
  
  If Client(Client.UBound).Sock_State <> 0 Then
    DoEvents
    'Loading this object might have a halt on slower computers
    'PROBLEM: bypassing/delaying any incoming Data from windows socket
    'so does this DoEvents helps the problem
    Load Client(Client.UBound + 1)
  End If
End Sub

Private Sub lstLoggedOut_DblClick()
Dim i As Byte
Dim j As Byte
Dim k As Byte
Dim tmp As Byte
Dim A As Object
  
  For i = 0 To lstLoggedOut.ListCount - 1
    If lstLoggedOut.Selected(i) = True Then
      tmp = Val(Mid(lstLoggedOut.List(i), 7, 2))
      lstLoggedOut.RemoveItem i
    End If
  Next i
      
  If tmpStartLog <> "" Then 'DISCONNECTED CLIENT
    Rst.AddNew
    Rst!Month = Format(tmpStartLog, "m")
    Rst!Day = Format(tmpStartLog, "d")
    Rst!Year = Format(tmpStartLog, "yyyy")
    Rst!Name = tmpName
    Rst!CompNum = tmpComputerNumber
    Rst!StartLog = FormatDateTime(tmpStartLog, vbLongTime)
    Rst!EndLog = FormatDateTime(Now, vbLongTime)
    Rst!Elapse = tmpPreviousElapse
    Rst!Service = tmpService
    If tmpPreviousAmount < MIN_AMT Then
      Rst!Amt = MIN_AMT
    Else
      Rst!Amt = tmpPreviousAmount
    End If
    Rst.Update
    
    'Clear BackUp Data
    Remove_BackUp_Data CByte(tmpComputerNumber)
    
    'Clears Temp Variables
    tmpName = ""
    tmpPreviousAmount = 0
    tmpAmount_Limited = 0
    tmpPreviousElapse = 0
    tmpEndLog = ""
    tmpService = 0
    tmpStartLog = ""
    
    'Clear Grid's Data
    Remove_Grid_Data CByte(tmpComputerNumber)

    Check_LstLogged 'Stop Blink Effect and Change Color
    
  Else 'CONNECTED CLIENT
  
    lvMain.ListItems(tmp).Tag = ""
          
    For j = 1 To NumberComps
      If Client(j).ComputerNumber = tmp Then
        Rst.AddNew
        Rst!Month = Format(Client(j).StartLog, "m")
        Rst!Day = Format(Client(j).StartLog, "d")
        Rst!Year = Format(Client(j).StartLog, "yyyy")
        Rst!Name = Client(j).CustomerName
        Rst!CompNum = Client(j).ComputerNumber
        Rst!StartLog = FormatDateTime(Client(j).StartLog, vbLongTime)
        Rst!EndLog = FormatDateTime(Client(j).EndLog, vbLongTime)
        Rst!Elapse = Client(j).PreviousElapse + Client(j).CurrentElapse
        Rst!Service = Client(j).Service
        If Client(j).PreviousAmount + Client(j).CurrentAmount < MIN_AMT Then
          Rst!Amt = MIN_AMT
        Else
          Rst!Amt = Client(j).PreviousAmount + Client(j).CurrentAmount
        End If
        Rst.Update
        
        'Clears BackUp Data
        Remove_BackUp_Data CByte(Client(j).ComputerNumber)

        'clears Variable's data
        Client(j).Reset_Var
                  
        'Clears Grid's Data
        Remove_Grid_Data CByte(Client(j).ComputerNumber)
                
        Check_LstLogged 'stop blink effect and change color
        Exit For
      End If
    Next j
  End If
    
End Sub

Private Sub lvMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Byte
Dim idx As Byte
  
  Select Case Button
  Case 2
    idx = Val(Trim(lvMain.SelectedItem))
  
    Mark = False
    mnuComp.Caption = "Comp " & idx
    For i = 1 To Client.UBound
      If Client(i).ComputerNumber = idx Then
        Mark = True
        Select Case Client(i).Status
        Case v_Connected
          If Client(i).Exceeded = False Then
            Show_Grid_Menu False, True, False, False, False, False
            'mnuLogIn.Enabled = True
            'mnudash.Visible = False
            'mnuExtension.Visible = False
            'mnuContinue.Visible = False
          Else
            If Client(i).Account = v_Limited Then 'Limited
              Show_Grid_Menu False, False, True, True, False, False
              'mnuExtension.Visible = True
              'mnuContinue.Visible = False
            Else 'Open
              Show_Grid_Menu False, False, True, False, True, False
              'mnuExtension.Visible = False
              'mnuContinue.Visible = True
            End If
            'mnudash.Visible = True
            'mnuLogIn.Enabled = False
          End If
          'mnuLogOut.Enabled = False
          'mnuPause.Visible = False
        Case v_LoggedConnected
          Show_Grid_Menu True, False, True, False, False, True
          'mnuLogIn.Enabled = False
          'mnuLogOut.Enabled = True
          'mnudash.Visible = True
          'mnuExtension.Visible = False
          'mnuContinue.Visible = False
          'mnuPause.Visible = True
        Case v_PausedLogged
          Show_Grid_Menu True, False, True, False, True, False
          'mnuLogIn.Enabled = False
          'mnuLogOut.Enabled = True
          'mnudash.Visible = True
          'mnuExtension.Visible = False
          'mnuContinue.Visible = True
          'mnuPause.Visible = False
        Case v_UnlockedConnected
          Exit Sub
        End Select
        Exit For
      End If
    Next i
    
    If Mark = False Then
      If lvMain.ListItems(idx).ListSubItems(4).Text <> "" And lvMain.ListItems(idx).Tag = "" Then
        Show_Grid_Menu True, False, False, False, False, False
        'mnuLogOut.Enabled = True
        'mnuLogIn.Enabled = False
        'mnudash.Visible = False
        'mnuExtension.Visible = False
        'mnuContinue.Visible = False
        'mnuPause.Visible = False
        'PopupMenu mnuGrid
      End If
    End If
  End Select
End Sub

Private Sub Show_Grid_Menu(LogOut As Boolean, LogIn As Boolean, Dash As Boolean, _
                          Extension As Boolean, Continue As Boolean, Pause As Boolean)
  mnuLogOut.Enabled = LogOut
  mnuLogIn.Enabled = LogIn
  mnudash.Visible = Dash
  mnuExtension.Visible = Extension
  mnuContinue.Visible = Continue
  mnuPause.Visible = Pause
  PopupMenu mnuGrid
End Sub

Private Sub mnuAbout_Click()
  frmAbout.Show 1
End Sub

Private Sub mnuClientControl_Click()
  SetPass = False
  Config = False
  frmPass.Show 1
End Sub

Private Sub mnuContinue_Click()
Dim Ctr As Byte
Dim Num As Byte
Dim i As Byte
  
  Num = Val(lvMain.SelectedItem)
    
  For i = 1 To Client.UBound - 1
    If Num = Client(i).ComputerNumber Then
      Select Case Client(i).Status
      Case v_Connected
        lvMain.ListItems(Num).ListSubItems(5).Text = ""
        Client(i).Exceeded = False
      Case v_PausedLogged
        Client(i).Enabled = True
        Client(i).InitialStart = Now
        Client(i).Status = v_LoggedConnected
        lvMain.ListItems(Client(i).ComputerNumber).SmallIcon = 3
        Exit Sub
      End Select
      Exit For
    End If
  Next i
  
  'Delete Entry in Listbox
  For i = 1 To lstLoggedOut.ListCount
    If Left(lstLoggedOut.List(i - 1), Len("Comp #" & Num)) = "Comp #" & Num Then
      lstLoggedOut.RemoveItem i - 1
    End If
  Next i
  
  'Stop Blink Effect and change color
  Check_LstLogged
End Sub

Private Sub mnuDaily_Click()
  frmDailyReport.Show 1
End Sub

Private Sub mnuClientLogs_Click()
  frmViewLogs.Show 1
End Sub

Private Sub mnuEx_daily_Click()
  Export = True
  frmDailyReport.Show 1
End Sub

Private Sub mnuEx_monthly_Click()
  Export = True
  frmMonthlyReport.Show 1
End Sub

Private Sub mnuExit_Click()
  End
End Sub

Private Sub mnuExtend_Click()
  frmExtend.Show 1
End Sub

Private Sub mnuExtension_Click()
  frmExtend.cmbCompNum.Text = Trim(lvMain.SelectedItem)
  frmExtend.Show 1
End Sub

Private Sub mnuKwh_Click()
On Error GoTo Irn
  KwhMon_Rst.MoveFirst
  If KwhMon_Rst.EOF = True Then
Irn: Dim tmp As String
    tmp = InputBox("Input the Initial KwhMeter reading!" & vbCrLf & vbCrLf & "Warning: Be sure that it's accurate and numbers only", "PUT-IN")
    If Val(tmp) = 0 Then
      MsgBox "Invalid Data!"
      Exit Sub
    Else
      KwhMon_Rst.AddNew
      KwhMon_Rst!KwhRead = Val(tmp)
      KwhMon_Rst.Update
    End If
  End If
  frmKwh.Show 1
End Sub

Private Sub mnuLimitedA_Click()
  TypeTrans = False
  frmLog.Show 1
End Sub

Private Sub mnuLock_Unlock_Click()
  frmLock_Unlock.Show 1
End Sub

Private Sub mnuLogOut_Click()
Dim i As Byte
  
On Error Resume Next
  lvMain.ListItems(CInt(lvMain.SelectedItem)).Tag = "1"
  
  If Mark = False Then
    tmpComputerNumber = Val(lvMain.SelectedItem)
    SvrDbRst.MoveFirst
    SvrDbRst.Find "ComNum LIKE " & tmpComputerNumber, 1, adSearchForward
    tmpStartLog = SvrDbRst!StartLog
    tmpEndLog = Now
    tmpPreviousAmount = SvrDbRst!Amt
    If lvMain.ListItems(tmpComputerNumber).ListSubItems(3).Text = "Internet" Then 'Internet
      tmpPreviousElapse = tmpPreviousAmount / (INTERNET_RATE / 60)
      tmpService = v_Internet
    Else 'Games/Rental
      tmpPreviousElapse = tmpPreviousAmount / (RENTAL_RATE / 60)
      tmpService = v_Rental_Games
    End If
    
    If tmpPreviousAmount < MIN_AMT Then
      lvMain.ListItems(tmpComputerNumber).ListSubItems(7).Text = FormatNumber(MIN_AMT, 2) 'amount
      lstLoggedOut.AddItem "Comp #" & Trim(Str(tmpComputerNumber)) & "  -----  P " & FormatNumber(MIN_AMT, 2)
    Else
      lvMain.ListItems(tmpComputerNumber).ListSubItems(7).Text = FormatNumber(tmpPreviousAmount, 2) 'amount
      lstLoggedOut.AddItem "Comp #" & Trim(Str(tmpComputerNumber)) & "  -----  P " & FormatNumber(tmpPreviousAmount, 2)
    End If
    
    lvMain.ListItems(tmpComputerNumber).ListSubItems(5).Text = FormatDateTime(tmpEndLog, vbLongTime) 'log OUT
    lvMain.ListItems(tmpComputerNumber).ListSubItems(6).Text = Formatter(tmpPreviousElapse)  'duration
  End If
  
  For i = 1 To Client.UBound - 1
    If Client(i).ComputerNumber = Val(lvMain.SelectedItem) Then
      Client(i).Enabled = False
      Client(i).SendData "ET1" & "@"
      Client(i).Status = v_Connected
      Client(i).Exceeded = True
      Client(i).EndLog = DateAdd("n", Client(i).PreviousElapse + Client(i).CurrentElapse, Client(i).StartLog)
      lvMain.ListItems(Client(i).ComputerNumber).ListSubItems(5).Text = FormatDateTime(Client(i).EndLog, vbLongTime)
      
      lvMain.ListItems(Client(i).ComputerNumber).SmallIcon = 2
      tmrLogBlinker.Enabled = True
      If Client(i).PreviousAmount + Client(i).CurrentAmount < MIN_AMT Then
        lvMain.ListItems(Client(i).ComputerNumber).ListSubItems(7).Text = FormatNumber(MIN_AMT, 2)
        lstLoggedOut.AddItem "Comp #" & Trim(Str(Client(i).ComputerNumber)) & "  -----  P " & FormatNumber(MIN_AMT, 2)
      Else
        lvMain.ListItems(Client(i).ComputerNumber).ListSubItems(7).Text = FormatNumber(Client(i).PreviousAmount + Client(i).CurrentAmount, 2)
        lstLoggedOut.AddItem "Comp #" & Trim(Str(Client(i).ComputerNumber)) & "  -----  P " & FormatNumber(Client(i).PreviousAmount + Client(i).CurrentAmount, 2)
      End If
      
      Exit For
    End If
  Next i
  
End Sub

Private Sub mnuMonthly_Click()
  frmMonthlyReport.Show 1
End Sub

Private Sub mnuOpen_Click()
  TypeTrans = True
  frmLog.cmbCompNum.Text = Trim(lvMain.SelectedItem)
  frmLog.Show 1
End Sub

Private Sub mnuOpenA_Click()
  TypeTrans = True
  frmLog.Show 1
End Sub

Private Sub mnuOptions_Click()
  SetPass = False
  Config = True
  frmPass.Show 1
End Sub

Private Sub mnuLimited_Click()
  TypeTrans = False
  frmLog.cmbCompNum.Text = Trim(lvMain.SelectedItem)
  frmLog.Show 1
End Sub

Private Sub mnuPause_Click()
Dim Num As Byte
Dim i As Byte
  Num = Val(lvMain.SelectedItem)
    
  For i = 1 To Client.UBound - 1
    If Num = Client(i).ComputerNumber Then
      lvMain.ListItems(Num).SmallIcon = 4
      Client(i).Enabled = False
      Client(i).Status = v_PausedLogged
      Exit For
    End If
  Next i
  
End Sub

Private Sub mnuTransfer_Click()
  frmTransfer.Show 1
End Sub

Private Sub sldVolMaster_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Byte
On Error Resume Next
  
  For i = 1 To Client.UBound - 1
    Client(i).SendData "VL" & sldVolMaster.Value & "@"
  Next i
End Sub

Private Sub tBar_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
    Case "Open"
      mnuOpenA_Click
    Case "Limited"
      mnuLimitedA_Click
    Case "Transfer"
      mnuTransfer_Click
    Case "Extend"
      mnuExtend_Click
    Case "Help"
      frmAbout.Show 1
  End Select
End Sub

Private Sub tmrCloser_Timer()
Dim i As Byte
  On Error Resume Next
  
  For i = 1 To Client.UBound - 1
    If Client(i).Sock_State = 8 Or Client(i).Sock_State = 4 _
    Or Client(i).Sock_State = 3 Then
      Client(i).Cloze
      lvMain.ListItems(Client(i).ComputerNumber).SmallIcon = 1
      Client(i).ComputerNumber = 0
      Client(i).Status = v_NotConnected
      Exit For
    End If
  Next i
End Sub

Private Sub tmrDateTime_Timer()
  lblDate.Caption = Format(Now, "dddddd") & " " & Time
End Sub

Private Sub tmrLogBlinker_Timer()
Static A As Boolean
  
  lblLog.ForeColor = vbRed
  If A = True Then
    lblLog.Visible = True
    Tray.Icon = imglCol.ListImages(5).Picture
    A = False
  Else
    lblLog.Visible = False
    Tray.Icon = imglCol.ListImages(6).Picture
    A = True
  End If
End Sub

Private Sub tmrResume_Timer()
On Error Resume Next
Dim i As Byte
   
  tmrResume.Enabled = False
  For i = 1 To Client.UBound - 1
    If Client(i).Status = v_Connected And Client(i).Exceeded = False Then
      With SvrDbRst
        .MoveFirst
        .Find "ComNum LIKE " & Client(i).ComputerNumber, 1, adSearchForward
        If !StartLog <> "" Then
          Open_Comp i
          tmrResume.Enabled = True
          Exit Sub
        End If
      End With
    End If
  Next i
  tmrResume.Enabled = True
End Sub

Public Sub Open_Comp(tmp As Byte)
Dim i As Byte
Dim strType As String
Dim Rte As String
  With SvrDbRst
    If !Unlock = True Then
      'Add in CleintMonitor's Dbase
      Mon_Rst.AddNew
      Mon_Rst!Month = Format(!StartLog, "m")
      Mon_Rst!Day = Format(!StartLog, "d")
      Mon_Rst!Year = Format(!StartLog, "yyyy")
      Mon_Rst!CN = !ComNum
      Mon_Rst!UnlockTime = FormatDateTime(!StartLog, vbLongTime)
      Mon_Rst!LockTime = FormatDateTime(Now, vbLongTime)
      Mon_Rst!Duration = DateDiff("n", !StartLog, Now)
      Mon_Rst.Update
      
      'Delete BackUp Entry
      !StartLog = Null
      !Amt = 0
      !Account = 0
      !Service = 0
      !Unlock = False
      SvrDbRst.Update
      
      'Clear Entry in DataGrid
      Remove_Grid_Data CByte(Client(tmp).ComputerNumber)
      
      If Client(tmp).Status = v_Connected Then
        Client(tmp).SendData "ET1" & "@"
      End If
    Else 'REsume Transaction
      Client(tmp).CustomerName = !Name
      Client(tmp).StartLog = !StartLog
      Client(tmp).Service = !Service
      Client(tmp).Account = !Account
      Client(tmp).Status = v_LoggedConnected
      Client(tmp).PreviousAmount = !Amt
      Client(tmp).InitialStart = Now
      Client(tmp).Amount_Limited = !AmtLimited
      Client(tmp).PreviousElapse = !Elapse
      Select Case Client(tmp).Service
      Case v_Internet
        strType = "Internet"
        Rte = Trim(Str(INTERNET_RATE))
      Case v_Rental_Games
        strType = "Gms/Rntl"
        Rte = Trim(Str(RENTAL_RATE))
      End Select
      lvMain.ListItems(Client(tmp).ComputerNumber).SmallIcon = 3
      If Client(tmp).Account = v_Open Then 'open
        Client(tmp).SendData "OP" & AddSpace(Client(tmp).CustomerName, 30) & strType & AddSpace(Rte, 2) & Trim(Client(tmp).StartLog) & "@"
      Else 'limited
        Client(tmp).EndLog = DateAdd("n", (Client(tmp).Amount_Limited - Client(tmp).PreviousAmount) * (60 / Val(Rte)), Client(tmp).InitialStart)
        lvMain.ListItems(Client(tmp).ComputerNumber).ListSubItems(5).Text = FormatDateTime(Client(tmp).EndLog, vbLongTime)
        Client(tmp).SendData "LT" & AddSpace(Client(tmp).CustomerName, 30) & strType & AddSpace(Trim(Str(Client(tmp).Amount_Limited)), 3) & AddSpace(Rte, 2) & Trim(Client(tmp).StartLog) & "@"
      End If
      Client(tmp).Enabled = True
    End If
  End With
End Sub

Private Sub Tray_MouseUp(Button As Integer)
  TrayPass = True
  frmPass.Show
End Sub

Public Sub Check_LstLogged()
  If lstLoggedOut.ListCount = 0 Then
    lblLog.ForeColor = &HE0E0E0
    tmrLogBlinker.Enabled = False
    Tray.Icon = imglCol.ListImages(5).Picture
    lblLog.Visible = True
  End If
End Sub
