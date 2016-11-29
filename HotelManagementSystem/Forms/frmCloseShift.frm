VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCloseShift 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cashier Report"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   5040
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdClose 
      Caption         =   "Close"
      Height          =   405
      Left            =   3480
      TabIndex        =   2
      Top             =   4740
      Width           =   1335
   End
   Begin VB.CommandButton CmdCloseShift 
      Caption         =   "Close Shift"
      Height          =   405
      Left            =   2130
      TabIndex        =   1
      Top             =   4740
      Width           =   1335
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   4290
      Left            =   210
      TabIndex        =   0
      Top             =   300
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   7567
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Type"
         Object.Width           =   5362
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Amount"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   4020
      Top             =   120
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
            Picture         =   "frmCloseShift.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCloseShift.frx":0A12
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCloseShift.frx":1424
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCloseShift.frx":1E36
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCloseShift.frx":2848
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCloseShift.frx":325A
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCloseShift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New Recordset

Private Sub CmdClose_Click()
    Unload Me
End Sub

Private Sub CmdCloseShift_Click()
    mdiMain.UnloadChilds
    
    CurrUser.USER_NAME = ""
    CurrUser.USER_PK = 0
    
    frmLogin.Show vbModal: If CloseMe = True Then Unload Me: Exit Sub: Exit Sub
End Sub

Private Sub Form_Load()
    'Set the graphics needed
    'Set the graphics for the controls
    With mdiMain
        'For listview
        Set lvList.SmallIcons = .i16x16
        Set lvList.Icons = .i16x16
    End With

    RefreshRecords
End Sub
    
Private Sub RefreshRecords()
    Me.Enabled = False
    
    If RS.State = adStateOpen Then RS.Close
    
    RS.Open "SELECT PaymentType, TotalAmount, PaymentTypeID FROM qry_rpt_Close_Shift WHERE [Date] = #" & Date & "# AND [CheckOutBy] = " & CurrUser.USER_PK & " ORDER BY PaymentTypeID ASC", CN, adOpenStatic, adLockOptimistic
    
    FillListView lvList, RS, 2, 2, False, True, "PaymentTypeID"
    
    Me.Enabled = True
End Sub
