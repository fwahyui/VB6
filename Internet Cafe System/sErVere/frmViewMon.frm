VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmLock_Unlock 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6645
   ClientLeft      =   1515
   ClientTop       =   765
   ClientWidth     =   7830
   ControlBox      =   0   'False
   Icon            =   "frmViewMon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   7830
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "X"
      Height          =   375
      Left            =   7460
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   375
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "View Inclusively by Date:"
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
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6255
      Begin VB.CommandButton cmdView 
         Caption         =   "V&iew"
         Height          =   315
         Left            =   4800
         TabIndex        =   8
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox cmbDay 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmViewMon.frx":000C
         Left            =   2160
         List            =   "frmViewMon.frx":0070
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.ComboBox cmbMonth 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmViewMon.frx":00EB
         Left            =   240
         List            =   "frmViewMon.frx":0116
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   1695
      End
      Begin VB.ComboBox cmbYear 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmViewMon.frx":0180
         Left            =   3360
         List            =   "frmViewMon.frx":0196
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Year:"
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
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Day:"
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
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Month:"
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
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
   End
   Begin MSComctlLib.ListView ListView 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   8493
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Rec#"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Comp #"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Unlock Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Lock Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Duration"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblType 
      BackColor       =   &H00000000&
      Caption         =   " LOCK/UNLOCK LOGS"
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
      TabIndex        =   9
      Top             =   0
      Width           =   10695
   End
End
Attribute VB_Name = "frmLock_Unlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ltmView As ListItem
Private Ctr As Long


Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdView_Click()

  Select Case cmbMonth.Text
  Case "--": cmbMonth.Tag = "0"
  Case "January": cmbMonth.Tag = "1"
  Case "February": cmbMonth.Tag = "2"
  Case "March": cmbMonth.Tag = "3"
  Case "April": cmbMonth.Tag = "4"
  Case "May": cmbMonth.Tag = "5"
  Case "June": cmbMonth.Tag = "6"
  Case "July": cmbMonth.Tag = "7"
  Case "August": cmbMonth.Tag = "8"
  Case "September": cmbMonth.Tag = "9"
  Case "October": cmbMonth.Tag = "10"
  Case "November": cmbMonth.Tag = "11"
  Case "December": cmbMonth.Tag = "12"
  End Select
  
  ListView.ListItems.Clear
  
  Mon_Rst.MoveFirst
  
  Ctr = 1
  Do Until Mon_Rst.EOF = True
    Select Case True
    Case cmbYear.Text = "--" And cmbMonth.Text = "--" And cmbDay.Text = "--"
      Form_Load
    Case cmbYear.Text = "--" And cmbMonth.Text = "--"
      Mon_Rst.Find "Day LIKE " & cmbDay.Text, 1, adSearchForward
      
    Case cmbYear.Text = "--" And cmbDay.Text = "--"
      Mon_Rst.Find "Month LIKE " & cmbMonth.Tag, 1, adSearchForward
      
    Case cmbMonth.Text = "--" And cmbDay.Text = "--"
      Mon_Rst.Find "Year LIKE " & cmbYear.Text, 1, adSearchForward
    
    Case cmbYear.Text = "--"
      Mon_Rst.Find "Month LIKE " & cmbMonth.Tag, 1, adSearchForward
      If Mon_Rst.EOF = True Then Exit Do
      If Mon_Rst!Day <> Val(cmbDay.Text) Then GoTo Slap
    
    Case cmbMonth.Text = "--"
      Mon_Rst.Find "Year LIKE " & cmbYear.Text, 1, adSearchForward
      If Mon_Rst.EOF = True Then Exit Do
      If Mon_Rst!Day <> Val(cmbDay.Text) Then GoTo Slap
    
    Case cmbDay.Text = "--"
      Mon_Rst.Find "Year LIKE " & cmbYear.Text, 1, adSearchForward
      If Mon_Rst.EOF = True Then Exit Do
      If Mon_Rst!Month <> Val(cmbMonth.Tag) Then GoTo Slap
    
    Case Else
      Mon_Rst.Find "Year LIKE " & cmbYear.Text, 1, adSearchForward
      If Mon_Rst.EOF = True Then Exit Do
      If Mon_Rst!Month <> Val(cmbMonth.Tag) Or Mon_Rst!Day <> Val(cmbDay.Text) Then GoTo Slap
    End Select
    
    If Mon_Rst.EOF = False Then
      Set ltmView = ListView.ListItems.Add(, , Ctr)
      ltmView.ListSubItems.Add , , Format(Mon_Rst!Month & "/" & Mon_Rst!Day & "/" & Mon_Rst!Year, "ddddd") 'Date
      ltmView.ListSubItems.Add , , Mon_Rst!CN 'Computer #
      ltmView.ListSubItems.Add , , Mon_Rst!UnlockTime 'Unlock Time
      ltmView.ListSubItems.Add , , Mon_Rst!LockTime 'Lock Time
      ltmView.ListSubItems.Add , , Formatter(Mon_Rst!Duration) 'Duration
      Ctr = Ctr + 1
    End If
Slap:
  Loop
End Sub

Private Sub Form_Load()
On Error GoTo Error
  Me.Icon = frmMain.Icon
  
  cmbMonth.Text = "--"
  cmbDay.Text = "--"
  cmbYear.Text = "--"
    
  Mon_Rst.MoveFirst
  Mon_Rst.MoveNext
       
  Ctr = 1
  Do Until Mon_Rst.EOF = True
    Set ltmView = ListView.ListItems.Add(, , Ctr)
    ltmView.ListSubItems.Add , , Format(Mon_Rst!Month & "/" & Mon_Rst!Day & "/" & Mon_Rst!Year, "ddddd") 'Date
    ltmView.ListSubItems.Add , , Mon_Rst!CN 'Computer #
    ltmView.ListSubItems.Add , , Mon_Rst!UnlockTime 'Unlock Time
    ltmView.ListSubItems.Add , , Mon_Rst!LockTime 'Lock Time
    ltmView.ListSubItems.Add , , Formatter(Mon_Rst!Duration) 'Duration
    Mon_Rst.MoveNext
    Ctr = Ctr + 1
  Loop

Error:
End Sub

Private Sub lblType_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  MoveForm Me
End Sub
