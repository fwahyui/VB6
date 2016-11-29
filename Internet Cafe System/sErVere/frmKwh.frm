VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmKwh 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6300
   ClientLeft      =   3390
   ClientTop       =   2640
   ClientWidth     =   6375
   ControlBox      =   0   'False
   Icon            =   "frmKwh.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   5760
      Width           =   1335
   End
   Begin VB.TextBox txtKwh 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   0
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "X"
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   375
   End
   Begin MSComctlLib.ListView ListView 
      Height          =   4575
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   8070
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "#"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   4234
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "FROM"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "TO"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Difference"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Warning: The Input (Meter Reading Data) must be accurate                       or else the operation will function abnormally"
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
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   5160
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Present Meter Reading:"
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
      Left            =   120
      TabIndex        =   3
      Top             =   5760
      Width           =   2055
   End
   Begin VB.Label lblType 
      BackColor       =   &H00000000&
      Caption         =   " KWH CONSUMPTION"
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
      TabIndex        =   1
      Top             =   0
      Width           =   6375
   End
End
Attribute VB_Name = "frmKwh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Ctr As Long
Private ltmView As ListItem

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdUpdate_Click()
  If MsgBox("Are you sure that the (meter reading data) is accurate", vbYesNo + vbExclamation, "Confirm") = vbYes Then
    If Val(txtKwh.Text) = 0 Then
      MsgBox "Invalid Data"
      Exit Sub
    End If
    KwhMon_Rst.AddNew
    KwhMon_Rst!Day = Format(Now, "d")
    KwhMon_Rst!Month = Format(Now, "m")
    KwhMon_Rst!Year = Format(Now, "yyyy")
    KwhMon_Rst!KwhRead = txtKwh.Text
    KwhMon_Rst.Update
    txtKwh.Text = ""
    cmdUpdate.Enabled = False
    MsgBox "The Data was Added Successfully!"
    Unload Me
    frmKwh.Show 1
  End If
End Sub

Private Sub Form_Load()
Dim from As Long
  KwhMon_Rst.MoveFirst
  from = KwhMon_Rst!KwhRead
  KwhMon_Rst.MoveNext
  Ctr = 1
  Do Until KwhMon_Rst.EOF = True
    Set ltmView = ListView.ListItems.Add(, , Ctr)
    ltmView.ListSubItems.Add , , Format(KwhMon_Rst!Month & "/" & KwhMon_Rst!Day & "/" & KwhMon_Rst!Year, "dddddd") 'Date
    ltmView.ListSubItems.Add , , from '<From> reading
    ltmView.ListSubItems.Add , , KwhMon_Rst!KwhRead '<To> reading
    ltmView.ListSubItems.Add , , KwhMon_Rst!KwhRead - from 'Difference
    from = KwhMon_Rst!KwhRead
    KwhMon_Rst.MoveNext
    Ctr = Ctr + 1
  Loop
End Sub

Private Sub lblType_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  MoveForm Me
End Sub
