VERSION 5.00
Object = "{9EDDC69F-10E8-4DE7-9648-EC8A45F005C0}#1.0#0"; "b8Controls4.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmPayment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Amount Paid Details"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9000
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   9000
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtValidTill 
      Height          =   315
      Left            =   4950
      TabIndex        =   7
      Top             =   2490
      Width           =   1245
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   6750
      TabIndex        =   10
      Top             =   5460
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.TextBox txtAuthCode 
      Height          =   315
      Left            =   3660
      TabIndex        =   6
      Top             =   2490
      Width           =   1245
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   375
      Left            =   7260
      TabIndex        =   9
      Top             =   2460
      Width           =   795
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   7740
      TabIndex        =   11
      Top             =   5460
      Width           =   1005
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   6330
      TabIndex        =   8
      Top             =   2460
      Width           =   825
   End
   Begin VB.TextBox txtAmountPaid 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4410
      TabIndex        =   3
      Text            =   "0.00"
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox txtBalance 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1290
      Width           =   1815
   End
   Begin VB.TextBox txtCardNumber 
      Height          =   315
      Left            =   2160
      TabIndex        =   5
      Top             =   2490
      Width           =   1455
   End
   Begin VB.TextBox txtGuestName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton btnRemove 
      Height          =   275
      Left            =   300
      Picture         =   "frmPayment.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Remove"
      Top             =   3030
      Visible         =   0   'False
      Width           =   275
   End
   Begin VB.PictureBox bgHeader 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   581
      TabIndex        =   12
      Top             =   0
      Width           =   8715
      Begin VB.Image Image1 
         Height          =   480
         Left            =   30
         Picture         =   "frmPayment.frx":01B2
         Top             =   30
         Width           =   480
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payments"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00926747&
         Height          =   345
         Left            =   600
         TabIndex        =   14
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fill all fields or fields with '*' then click 'Save' button to update."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00926747&
         Height          =   180
         Left            =   600
         TabIndex        =   13
         Top             =   360
         Width           =   3900
      End
   End
   Begin b8Controls4.b8Line b8Line5 
      Height          =   30
      Left            =   180
      TabIndex        =   19
      Top             =   5340
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   53
      BorderColor1    =   15592425
      BorderColor2    =   16777215
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   4410
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   840
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   661
      _Version        =   393216
      Format          =   27721729
      CurrentDate     =   39539
   End
   Begin b8Controls4.b8GradLine b8GradLine2 
      Height          =   240
      Left            =   180
      TabIndex        =   20
      Top             =   1830
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   423
      Color1          =   14737632
      Color2          =   16119285
      Caption         =   "Payment Method"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
   End
   Begin MSDataListLib.DataCombo dcPaymentType 
      Height          =   315
      Left            =   270
      TabIndex        =   4
      Top             =   2490
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   "Cash"
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   2220
      Left            =   240
      TabIndex        =   21
      Top             =   2940
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   3916
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   275
      ForeColorFixed  =   -2147483640
      BackColorSel    =   1091552
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483633
      GridColorUnpopulated=   -2147483633
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.PictureBox bgMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E3F9FB&
      Height          =   5415
      Left            =   0
      ScaleHeight     =   361
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   599
      TabIndex        =   15
      Top             =   570
      Width           =   8985
      Begin b8Controls4.b8Line b8Line2 
         Height          =   30
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   53
         BorderColor1    =   15592425
         BorderColor2    =   16777215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Guest Name"
         Height          =   195
         Left            =   90
         TabIndex        =   29
         Top             =   300
         Width           =   885
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Type"
         Height          =   195
         Left            =   270
         TabIndex        =   28
         Top             =   1680
         Width           =   1020
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Card Number"
         Height          =   195
         Left            =   2160
         TabIndex        =   27
         Top             =   1650
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Balance"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   810
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Paid"
         Height          =   195
         Left            =   3450
         TabIndex        =   25
         Top             =   780
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Auth Code"
         Height          =   195
         Left            =   3660
         TabIndex        =   24
         Top             =   1650
         Width           =   750
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Left            =   3990
         TabIndex        =   23
         Top             =   270
         Width           =   345
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valid Till"
         Height          =   195
         Left            =   4950
         TabIndex        =   22
         Top             =   1650
         Width           =   585
      End
      Begin VB.Label lblRM 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   9450
         TabIndex        =   17
         Top             =   3030
         Width           =   45
      End
   End
End
Attribute VB_Name = "frmPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FolioNumber      As String
Public GuestName        As String
Public Balance          As Currency
Public RefForm          As Form 'Calling form

Dim AmountPaid          As Currency
Dim cIRowCount          As Integer

Private Sub btnRemove_Click()
    'Remove selected load product
    With Grid
        'Update the record count
        cIRowCount = cIRowCount - 1
        
        If .Rows = 2 Then Grid.Rows = Grid.Rows + 1
        .RemoveItem (.RowSel)
    End With

    btnRemove.Visible = False
    
    Balance = Balance + toNumber(Grid.TextMatrix(Grid.RowSel, 8))
    txtBalance.Text = toMoney(Balance)
    
    Grid_Click
End Sub

Private Sub cmdAdd_Click()
    If Trim(txtAmountPaid.Text) = "" Or Trim(txtAmountPaid.Text) = "0.00" Then Exit Sub

'    If toNumber(txtAmountPaid) > Balance Then
'        MsgBox "Amount paid exceed balance. Please enter correct amount.", vbInformation
'        txtAmountPaid.SetFocus
'
'        Exit Sub
'    End If

    Dim CurrRow As Integer
    Dim intUnitID As Integer

    'Add to grid
    With Grid

        'Perform if the record is not exist
        If .Rows = 2 And .TextMatrix(1, 2) = "" Then
            .TextMatrix(1, 2) = FolioNumber
            .TextMatrix(1, 3) = dtpDate.Value
            .TextMatrix(1, 4) = dcPaymentType.Text
            .TextMatrix(1, 5) = txtCardNumber.Text
            .TextMatrix(1, 6) = txtAuthCode.Text
            .TextMatrix(1, 7) = txtValidTill.Text
            .TextMatrix(1, 8) = toMoney(txtAmountPaid.Text)
        Else
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 2) = FolioNumber
            .TextMatrix(.Rows - 1, 3) = dtpDate.Value
            .TextMatrix(.Rows - 1, 4) = dcPaymentType.Text
            .TextMatrix(.Rows - 1, 5) = txtCardNumber.Text
            .TextMatrix(.Rows - 1, 6) = txtAuthCode.Text
            .TextMatrix(.Rows - 1, 7) = txtValidTill.Text
            .TextMatrix(.Rows - 1, 8) = toMoney(txtAmountPaid.Text)

            .Row = .Rows - 1
        End If
        'Increase the record count
        cIRowCount = cIRowCount + 1
        
        txtBalance.Text = toMoney(Balance - toNumber(txtAmountPaid.Text))
    
        Balance = Balance - toNumber(txtAmountPaid.Text)
        
        'Highlight the current row's column
        .ColSel = 8
        'Display a remove button
        Call Grid_Click
        
        Call ResetFields
    End With
End Sub

Private Sub CmdClose_Click()
    Dim c As Integer
    
    With Grid
        'Save the details of the records
        For c = 1 To cIRowCount
            .Row = c
            
            AmountPaid = AmountPaid + .TextMatrix(c, 8)
        Next c
    End With

    'Clear variables
    c = 0
    RefForm.AmountPaid = AmountPaid
    
    Me.Hide
End Sub

Public Sub cmdSave_Click()
    Dim rsPayments As New Recordset

    rsPayments.CursorLocation = adUseClient
    rsPayments.Open "SELECT * FROM Payments WHERE FolioNumber='" & FolioNumber & "'", CN, adOpenStatic, adLockOptimistic
    
    DeleteItems
    
    Dim c As Integer
    
    With Grid
        'Save the details of the records
        For c = 1 To cIRowCount
            .Row = c
            
            If .TextMatrix(c, 1) = "" Then
                rsPayments.AddNew
                
                rsPayments![FolioNumber] = FolioNumber
            Else
                rsPayments.Filter = "PaymentID = " & toNumber(.TextMatrix(c, 1))
            
                If rsPayments.RecordCount = 0 Then
                    rsPayments.AddNew
                    
                    rsPayments![FolioNumber] = FolioNumber
                End If
            End If

            rsPayments![Date] = .TextMatrix(c, 3)
            rsPayments![PaymentTypeID] = getValueAt("SELECT * FROM [Payment Type] WHERE PaymentType = '" & .TextMatrix(c, 4) & "'", "PaymentTypeID")
            rsPayments![CardNumber] = .TextMatrix(c, 5)
            rsPayments![AuthorityNumber] = .TextMatrix(c, 6)
            rsPayments![ValidTill] = .TextMatrix(c, 7)
            rsPayments![Amount] = .TextMatrix(c, 8)

            rsPayments.Update

            AmountPaid = AmountPaid + .TextMatrix(c, 8)
        Next c
    End With

    'Clear variables
    c = 0
    Set rsPayments = Nothing
    
    Unload frmPayment
End Sub

Private Sub cmdUpdate_Click()
    With Grid
        .TextMatrix(.RowSel, 3) = dtpDate.Value
        .TextMatrix(.RowSel, 4) = dcPaymentType.Text
        .TextMatrix(.RowSel, 5) = txtCardNumber.Text
        .TextMatrix(.RowSel, 6) = txtAuthCode.Text
        .TextMatrix(.RowSel, 7) = txtValidTill.Text
        .TextMatrix(.RowSel, 8) = toMoney(txtAmountPaid.Text)
    End With
End Sub

Private Sub dcPaymentType_Click(Area As Integer)
    If dcPaymentType.Text = "Cash Refund" Then
        If txtAmountPaid.Text > 0 Then txtAmountPaid.Text = toMoney(toNumber(-txtAmountPaid.Text))
    End If
End Sub

Private Sub Form_Activate()
    AmountPaid = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    Call InitGrid
    
    bind_dc "SELECT * FROM [Payment Type] ORDER BY PaymentType", "PaymentType", dcPaymentType, "paymentTypeID", False
    
    dcPaymentType.Text = "Cash"
    dtpDate.Value = Date
    
    txtGuestName.Text = GuestName
    
    DisplayForEditing
End Sub

'Procedure used to initialize the grid
Private Sub InitGrid()
    cIRowCount = 0
    With Grid
        .Clear
        .ClearStructure
        .Rows = 2
        .FixedRows = 1
        .FixedCols = 1
        .Cols = 9
        .ColSel = 8
        'Initialize the column size
        .ColWidth(0) = 315
        .ColWidth(1) = 0
        .ColWidth(2) = 0
        .ColWidth(3) = 800
        .ColWidth(4) = 1200
        .ColWidth(5) = 1200
        .ColWidth(6) = 1200
        .ColWidth(7) = 1200
        .ColWidth(8) = 1200

        'Initialize the column name
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Payment ID"
        .TextMatrix(0, 2) = "Folio Number"
        .TextMatrix(0, 3) = "Date"
        .TextMatrix(0, 4) = "Payment Type"
        .TextMatrix(0, 5) = "Card Number"
        .TextMatrix(0, 6) = "Auth Code"
        .TextMatrix(0, 7) = "Valid Till"
        .TextMatrix(0, 8) = "Amount Paid"
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPayment = Nothing
End Sub

Private Sub Grid_Click()
    With Grid
        If .TextMatrix(.RowSel, 3) = "" Then Exit Sub
        
        dtpDate.Value = .TextMatrix(.RowSel, 3)
        dcPaymentType.Text = .TextMatrix(.RowSel, 4)
        txtCardNumber.Text = .TextMatrix(.RowSel, 5)
        txtAuthCode.Text = .TextMatrix(.RowSel, 6)
        txtValidTill.Text = .TextMatrix(.RowSel, 7)
        txtAmountPaid.Text = .TextMatrix(.RowSel, 8)
    
        If Grid.Rows = 2 And Grid.TextMatrix(1, 2) = "" Then
            btnRemove.Visible = False
        Else
            btnRemove.Visible = True
            btnRemove.Top = (Grid.CellTop + Grid.Top) - 20
            btnRemove.Left = Grid.Left + 50
        End If
    End With
End Sub

Private Sub ResetFields()
    txtAmountPaid.Text = ""
    txtCardNumber.Text = ""
    txtAuthCode.Text = ""
    
    txtAmountPaid.SetFocus
End Sub

Private Sub DisplayForEditing()
    On Error GoTo err
    
    'Display the details
    Dim rsPayments As New Recordset

    cIRowCount = 0
    
    rsPayments.CursorLocation = adUseClient
    rsPayments.Open "SELECT * FROM qry_Payments WHERE FolioNumber='" & FolioNumber & "'", CN, adOpenStatic, adLockOptimistic
    
    If rsPayments.RecordCount > 0 Then
        rsPayments.MoveFirst
        While Not rsPayments.EOF
          cIRowCount = cIRowCount + 1     'increment
            With Grid
                If .Rows = 2 And .TextMatrix(1, 1) = "" Then
                    .TextMatrix(1, 1) = rsPayments!PaymentID
                    .TextMatrix(1, 2) = rsPayments!FolioNumber
                    .TextMatrix(1, 3) = rsPayments!Date
                    .TextMatrix(1, 4) = rsPayments!PaymentType
                    .TextMatrix(1, 5) = rsPayments!CardNumber
                    .TextMatrix(1, 6) = rsPayments!AuthorityNumber
                    .TextMatrix(1, 7) = rsPayments!ValidTill
                    .TextMatrix(1, 8) = toMoney(rsPayments!Amount)
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = rsPayments!PaymentID
                    .TextMatrix(.Rows - 1, 2) = rsPayments!FolioNumber
                    .TextMatrix(.Rows - 1, 3) = rsPayments!Date
                    .TextMatrix(.Rows - 1, 4) = rsPayments!PaymentType
                    .TextMatrix(.Rows - 1, 5) = rsPayments!CardNumber
                    .TextMatrix(.Rows - 1, 6) = rsPayments!AuthorityNumber
                    .TextMatrix(.Rows - 1, 7) = rsPayments!ValidTill
                    .TextMatrix(.Rows - 1, 8) = toMoney(rsPayments!Amount)
                End If
            End With
            
'            AmountPaid = rsPayments!Amount
            
            rsPayments.MoveNext
        Wend
        Grid.Row = 1
        Grid.ColSel = 8
        'Set fixed cols
        If State = adStateEditMode Then
            Grid.FixedRows = Grid.Row: 'Grid.SelectionMode = flexSelectionFree
            Grid.FixedCols = 1
        End If
    End If

    rsPayments.Close
    'Clear variables
    Set rsPayments = Nothing

    Exit Sub
err:
    If err.Number = 94 Then Resume Next
    
    prompt_err err, Name, "DisplayForEditing"
    Screen.MousePointer = vbDefault
End Sub

Private Sub DeleteItems()
    Dim CurrRow As Integer
    Dim rsPayments As New Recordset
    
'    If State = adStateAddMode Then Exit Sub
    
    rsPayments.CursorLocation = adUseClient
    rsPayments.Open "SELECT * FROM Payments WHERE FolioNumber='" & FolioNumber & "'", CN, adOpenStatic, adLockOptimistic
    If rsPayments.RecordCount > 0 Then
        rsPayments.MoveFirst
        While Not rsPayments.EOF
            CurrRow = getFlexPos(Grid, 1, rsPayments!PaymentID)
        
            'Add to grid
            With Grid
                If CurrRow < 0 Then
                    'Delete record if doesnt exist in flexgrid
                    DelRecwSQL "Payments", "PaymentID", "", True, rsPayments!PaymentID
                End If
            End With
            rsPayments.MoveNext
        Wend
    End If
End Sub

Public Sub RefreshBalance()
    txtBalance.Text = toMoney(Balance)
End Sub

Private Sub txtAmountPaid_GotFocus()
    HLText txtAmountPaid
End Sub

Private Sub txtAmountPaid_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtAmountPaid_Validate(Cancel As Boolean)
    txtAmountPaid.Text = toMoney(txtAmountPaid.Text)
End Sub
