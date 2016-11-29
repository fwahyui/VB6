VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmViewLogs 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8070
   ClientLeft      =   630
   ClientTop       =   315
   ClientWidth     =   10695
   ControlBox      =   0   'False
   Icon            =   "frmViewLogs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   10695
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Totals:"
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
      Left            =   7680
      TabIndex        =   11
      Top             =   480
      Width           =   2895
      Begin VB.Label lblGross 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   480
         TabIndex        =   12
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "X"
      Height          =   375
      Left            =   10320
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   375
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
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
         ItemData        =   "frmViewLogs.frx":000C
         Left            =   2160
         List            =   "frmViewLogs.frx":0070
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.ComboBox cmbMonth 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmViewLogs.frx":00EB
         Left            =   240
         List            =   "frmViewLogs.frx":0116
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   1695
      End
      Begin VB.ComboBox cmbYear 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmViewLogs.frx":0180
         Left            =   3360
         List            =   "frmViewLogs.frx":0196
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
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   11033
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   9
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
         Text            =   "Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Comp #"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Service Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Log In Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Log Out Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Duration"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Amt (P)"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblType 
      BackColor       =   &H00000000&
      Caption         =   " ACCOUNT LOGS"
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
Attribute VB_Name = "frmViewLogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ltmView As ListItem
Private Ctr As Long
Private Total As Long


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
  
  Rst.MoveFirst
  Total = 0
  Ctr = 1
  Do Until Rst.EOF = True
    Select Case True
    Case cmbYear.Text = "--" And cmbMonth.Text = "--" And cmbDay.Text = "--"
      Form_Load
    Case cmbYear.Text = "--" And cmbMonth.Text = "--"
      Rst.Find "Day LIKE " & cmbDay.Text, 1, adSearchForward
      
    Case cmbYear.Text = "--" And cmbDay.Text = "--"
      Rst.Find "Month LIKE " & cmbMonth.Tag, 1, adSearchForward
      
    Case cmbMonth.Text = "--" And cmbDay.Text = "--"
      Rst.Find "Year LIKE " & cmbYear.Text, 1, adSearchForward
    
    Case cmbYear.Text = "--"
      Rst.Find "Month LIKE " & cmbMonth.Tag, 1, adSearchForward
      If Rst.EOF = True Then Exit Do
      If Rst!Day <> Val(cmbDay.Text) Then GoTo Slap
    
    Case cmbMonth.Text = "--"
      Rst.Find "Year LIKE " & cmbYear.Text, 1, adSearchForward
      If Rst.EOF = True Then Exit Do
      If Rst!Day <> Val(cmbDay.Text) Then GoTo Slap
    
    Case cmbDay.Text = "--"
      Rst.Find "Year LIKE " & cmbYear.Text, 1, adSearchForward
      If Rst.EOF = True Then Exit Do
      If Rst!Month <> Val(cmbMonth.Tag) Then GoTo Slap
    
    Case Else
      Rst.Find "Year LIKE " & cmbYear.Text, 1, adSearchForward
      If Rst.EOF = True Then Exit Do
      If Rst!Month <> Val(cmbMonth.Tag) Or Rst!Day <> Val(cmbDay.Text) Then GoTo Slap
    End Select
    
    If Rst.EOF = False Then
      Set ltmView = ListView.ListItems.Add(, , Ctr)
      ltmView.ListSubItems.Add , , Format(Rst!Month & "/" & Rst!Day & "/" & Rst!Year, "ddddd") 'Date
      ltmView.ListSubItems.Add , , Format(Rst!Name, "")
      ltmView.ListSubItems.Add , , Rst!CompNum 'Computer #
      If Rst!Service = v_Internet Then 'Service Type
        ltmView.ListSubItems.Add , , "Internet" 'Internet
      Else
        ltmView.ListSubItems.Add , , "Gms/Rntl" 'Games/Rental
      End If
      ltmView.ListSubItems.Add , , Rst!StartLog 'Log In Time
      ltmView.ListSubItems.Add , , Rst!EndLog 'Log Out Time
      ltmView.ListSubItems.Add , , Formatter(Rst!Elapse) 'Duration
      ltmView.ListSubItems.Add , , FormatNumber(Rst!Amt, 2) 'Amount
      
      Total = Total + Rst!Amt
      Ctr = Ctr + 1
    End If
Slap:
  Loop
  lblGross.Caption = FormatNumber(Total, 2)
End Sub

Private Sub Form_Load()
  Me.Icon = frmMain.Icon
   
  cmbMonth.Text = "--"
  cmbDay.Text = "--"
  cmbYear.Text = "--"
  
  Rst.MoveFirst
  Rst.MoveNext
       
  Total = 0
  Ctr = 1
  Do Until Rst.EOF = True
    Set ltmView = ListView.ListItems.Add(, , Ctr)
    ltmView.ListSubItems.Add , , Format(Rst!Month & "/" & Rst!Day & "/" & Rst!Year, "ddddd") 'Date
    ltmView.ListSubItems.Add , , Format(Rst!Name, "") 'Name
    ltmView.ListSubItems.Add , , Rst!CompNum 'Computer #
    If Rst!Service = v_Internet Then 'Service Type
      ltmView.ListSubItems.Add , , "Internet" 'Internet
    Else
      ltmView.ListSubItems.Add , , "Gms/Rntl" 'Games/Rental
    End If
    ltmView.ListSubItems.Add , , Rst!StartLog 'Log In Time
    ltmView.ListSubItems.Add , , Rst!EndLog 'Log Out Time
    ltmView.ListSubItems.Add , , Formatter(Rst!Elapse) 'Duration
    ltmView.ListSubItems.Add , , FormatNumber(Rst!Amt, 2) 'Amount
       
    Total = Total + Rst!Amt
    Ctr = Ctr + 1
    Rst.MoveNext
  Loop
  lblGross.Caption = FormatNumber(Total, 2)
End Sub

Private Sub lblType_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  MoveForm Me
End Sub
