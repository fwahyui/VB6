VERSION 5.00
Object = "{9EDDC69F-10E8-4DE7-9648-EC8A45F005C0}#1.0#0"; "b8Controls4.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmAccountReceivable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payments"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7380
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox bgHeader 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   457
      TabIndex        =   9
      Top             =   0
      Width           =   6855
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
         TabIndex        =   11
         Top             =   360
         Width           =   3900
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
         TabIndex        =   10
         Top             =   30
         Width           =   1395
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   30
         Picture         =   "frmAccountReceivable.frx":0000
         Top             =   30
         Width           =   480
      End
   End
   Begin VB.CommandButton btnRemove 
      Height          =   275
      Left            =   300
      Picture         =   "frmAccountReceivable.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Remove"
      Top             =   3270
      Visible         =   0   'False
      Width           =   275
   End
   Begin VB.TextBox txtCompany 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtBalance 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1290
      Width           =   1815
   End
   Begin VB.TextBox txtAmountPaid 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1290
      Width           =   1815
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   2460
      Width           =   825
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6120
      TabIndex        =   3
      Top             =   5490
      Width           =   1005
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   375
      Left            =   6330
      TabIndex        =   2
      Top             =   2460
      Width           =   795
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   5490
      Width           =   885
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4050
      TabIndex        =   0
      Text            =   "0.00"
      Top             =   2490
      Width           =   1245
   End
   Begin b8Controls4.b8Line b8Line5 
      Height          =   30
      Left            =   180
      TabIndex        =   12
      Top             =   5340
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   53
      BorderColor1    =   15592425
      BorderColor2    =   16777215
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   315
      Left            =   240
      TabIndex        =   13
      Top             =   2490
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      Format          =   78970881
      CurrentDate     =   39539
   End
   Begin b8Controls4.b8GradLine b8GradLine2 
      Height          =   240
      Left            =   180
      TabIndex        =   14
      Top             =   1830
      Width           =   7005
      _ExtentX        =   12356
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
      Left            =   2160
      TabIndex        =   15
      Top             =   2490
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   2190
      Left            =   240
      TabIndex        =   16
      Top             =   2970
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   3863
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
      ScaleWidth      =   491
      TabIndex        =   17
      Top             =   570
      Width           =   7365
      Begin b8Controls4.b8Line b8Line2 
         Height          =   30
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   53
         BorderColor1    =   15592425
         BorderColor2    =   16777215
      End
      Begin VB.Label lblPaid 
         BackStyle       =   0  'Transparent
         Caption         =   "PAID"
         BeginProperty Font 
            Name            =   "Andalus"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   4410
         TabIndex        =   26
         Top             =   150
         Visible         =   0   'False
         Width           =   795
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
         TabIndex        =   25
         Top             =   3030
         Width           =   45
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   195
         Left            =   4050
         TabIndex        =   24
         Top             =   1680
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Left            =   270
         TabIndex        =   23
         Top             =   1650
         Width           =   345
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Paid"
         Height          =   195
         Left            =   3450
         TabIndex        =   22
         Top             =   780
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Balance"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   810
         Width           =   585
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Type"
         Height          =   195
         Left            =   2160
         TabIndex        =   20
         Top             =   1680
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   300
         Width           =   660
      End
   End
End
Attribute VB_Name = "frmAccountReceivable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK               As Long
Public Company          As String
Public Balance          As Currency
Public AmountPaid       As Currency

Dim cIRowCount          As Integer
Dim Amount              As Currency

Private Sub btnRemove_Click()
    'Remove selected load product
    With Grid
        'Update the record count
        cIRowCount = cIRowCount - 1
        
        If .Rows = 2 Then Grid.Rows = Grid.Rows + 1
        .RemoveItem (.RowSel)
    End With

    btnRemove.Visible = False

    txtBalance.Text = toMoney(toNumber(txtBalance.Text) + toNumber(Grid.TextMatrix(Grid.RowSel, 5)))
    txtAmountPaid.Text = toMoney(txtAmountPaid.Text) - toNumber(Grid.TextMatrix(Grid.RowSel, 5))
    
    Grid_Click
End Sub

Private Sub cmdAdd_Click()
    If Trim(txtAmount.Text) = "0.00" Then Exit Sub

    If toNumber(txtAmount) > Balance Then
        MsgBox "Amount paid exceed balance. Please enter correct amount.", vbInformation
        txtAmount.SetFocus
        
        Exit Sub
    End If

    Dim CurrRow As Integer

    'Add to grid
    With Grid

        'Perform if the record is not exist
        If .Rows = 2 And .TextMatrix(1, 2) = "" Then
            .TextMatrix(1, 2) = dtpDate.Value
            .TextMatrix(1, 3) = dcPaymentType.BoundText
            .TextMatrix(1, 4) = dcPaymentType.Text
            .TextMatrix(1, 5) = toMoney(txtAmount.Text)
        Else
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 2) = dtpDate.Value
            .TextMatrix(.Rows - 1, 3) = dcPaymentType.BoundText
            .TextMatrix(.Rows - 1, 4) = dcPaymentType.Text
            .TextMatrix(.Rows - 1, 5) = toMoney(txtAmount.Text)

            .Row = .Rows - 1
        End If
        'Increase the record count
        cIRowCount = cIRowCount + 1
        
        txtBalance.Text = toMoney(Balance - toNumber(txtAmount.Text))
'        Balance = Balance - toNumber(txtAmountPaid.Text)
        
        txtAmountPaid.Text = toMoney(toNumber(txtAmountPaid.Text) + toNumber(txtAmount.Text))
        
        'Highlight the current row's column
        .ColSel = 5
        'Display a remove button
        Call Grid_Click
        
        Call ResetFields
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Public Sub cmdSave_Click()
    Dim rsPayments As New Recordset

    rsPayments.CursorLocation = adUseClient
    rsPayments.Open "SELECT * FROM [Accounts Receivable] WHERE AccRecID=" & PK, CN, adOpenStatic, adLockOptimistic
    
    With rsPayments
        .Fields("Debit") = txtAmountPaid.Text
        
        .Update
    End With
    
    DeleteItems
    
    Dim rsPaymentRec As New Recordset
    rsPaymentRec.CursorLocation = adUseClient
    rsPaymentRec.Open "SELECT * FROM [Payments Received] WHERE AccRecID=" & PK, CN, adOpenStatic, adLockOptimistic

    Dim c As Integer
    
    With Grid
        'Save the details of the records
        For c = 1 To cIRowCount
            .Row = c
            
            If .TextMatrix(c, 1) = "" Then
                rsPaymentRec.AddNew
                
                rsPaymentRec![AccRecID] = PK
            Else
                rsPaymentRec.Filter = "RecievedPaymentID = " & toNumber(.TextMatrix(c, 1))
            
'                If rsPaymentRec.RecordCount = 0 Then
'                    rsPaymentRec.AddNew
'
'                    rsPaymentRec![AccRecID] = PK
'                End If
            End If

            rsPaymentRec![Date] = .TextMatrix(c, 2)
            rsPaymentRec![PaymentTypeID] = .TextMatrix(c, 3)
            rsPaymentRec![Amount] = .TextMatrix(c, 5)

            rsPaymentRec.Update
        Next c
    End With

    'Clear variables
    c = 0
    Set rsPayments = Nothing
    Set rsPaymentRec = Nothing
    
    Unload frmAccountReceivable
End Sub

Private Sub cmdUpdate_Click()
    If Trim(txtAmount.Text) = "0.00" Then Exit Sub
    
    txtBalance.Text = toMoney(txtBalance.Text) + Amount
    txtAmountPaid.Text = toMoney(txtAmountPaid.Text) - Amount

    txtBalance.Text = toMoney(toNumber(txtBalance.Text) - toNumber(txtAmount.Text))
    txtAmountPaid.Text = toMoney(toNumber(txtAmountPaid.Text) + toNumber(txtAmount.Text))
    
    With Grid
        .TextMatrix(.RowSel, 2) = dtpDate.Value
        .TextMatrix(.RowSel, 3) = dcPaymentType.BoundText
        .TextMatrix(.RowSel, 4) = dcPaymentType.Text
        .TextMatrix(.RowSel, 5) = toMoney(txtAmount.Text)
    End With

    Call Grid_Click
    
    Call ResetFields
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    Call InitGrid
    
    bind_dc "SELECT * FROM [Payment Type]", "PaymentType", dcPaymentType, "paymentTypeID", True

    txtCompany.Text = Company
    txtBalance.Text = toMoney(Balance)
    txtAmountPaid.Text = toMoney(AmountPaid)
    
    dtpDate.Value = Date
    
    DisplayForEditing
    
    If Balance = "0.00" Then
        Grid.Height = 3000
        Grid.Top = 2160
        
        dtpDate.Visible = False
        dcPaymentType.Visible = False
        txtAmount.Visible = False
        cmdAdd.Visible = False
        cmdUpdate.Visible = False
        lblPaid.Visible = True
    End If
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
        .Cols = 6
        .ColSel = 5
        'Initialize the column size
        .ColWidth(0) = 315
        .ColWidth(1) = 0
        .ColWidth(2) = 1200
        .ColWidth(3) = 0
        .ColWidth(4) = 1200
        .ColWidth(5) = 1200

        'Initialize the column name
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "RecievedPaymentID"
        .TextMatrix(0, 2) = "Date"
        .TextMatrix(0, 3) = "Payment Type ID"
        .TextMatrix(0, 4) = "Payment Type"
        .TextMatrix(0, 5) = "Amount"
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmAccountReceivableList.RefreshRecords
    
    Set frmAccountReceivable = Nothing
End Sub

Private Sub Grid_Click()
    With Grid
        dtpDate.Value = .TextMatrix(.RowSel, 2)
        dcPaymentType.BoundText = .TextMatrix(.RowSel, 3)
        txtAmount.Text = toMoney(.TextMatrix(.RowSel, 5))
        
        Amount = .TextMatrix(.RowSel, 5)
        
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
    txtAmount.Text = ""
    
    txtAmount.SetFocus
End Sub

Private Sub DisplayForEditing()
    On Error GoTo err
    
    'Display the details
    Dim rsPayments As New Recordset

    cIRowCount = 0
    
    rsPayments.CursorLocation = adUseClient
    rsPayments.Open "SELECT * FROM [qry_Payment_Received] WHERE AccRecID=" & PK, CN, adOpenStatic, adLockOptimistic
    
    If rsPayments.RecordCount > 0 Then
        rsPayments.MoveFirst
        While Not rsPayments.EOF
          cIRowCount = cIRowCount + 1     'increment
            With Grid
                If .Rows = 2 And .TextMatrix(1, 1) = "" Then
                    .TextMatrix(1, 1) = rsPayments!RecievedPaymentID
                    .TextMatrix(1, 2) = rsPayments!Date
                    .TextMatrix(1, 3) = rsPayments!PaymentTypeID
                    .TextMatrix(1, 4) = rsPayments!PaymentType
                    .TextMatrix(1, 5) = toMoney(rsPayments!Amount)
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = rsPayments!RecievedPaymentID
                    .TextMatrix(.Rows - 1, 2) = rsPayments!Date
                    .TextMatrix(.Rows - 1, 3) = rsPayments!PaymentTypeID
                    .TextMatrix(.Rows - 1, 4) = rsPayments!PaymentType
                    .TextMatrix(.Rows - 1, 5) = toMoney(rsPayments!Amount)
                End If
            End With
            
'            AmountPaid = rsPayments!Amount
            
            rsPayments.MoveNext
        Wend
        Grid.Row = 1
        Grid.ColSel = 5
        'Set fixed cols

        Grid.FixedRows = Grid.Row: 'Grid.SelectionMode = flexSelectionFree
        Grid.FixedCols = 1
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
    
    rsPayments.CursorLocation = adUseClient
    rsPayments.Open "SELECT * FROM [Payments Received] WHERE AccRecID=" & PK, CN, adOpenStatic, adLockOptimistic
    If rsPayments.RecordCount > 0 Then
        rsPayments.MoveFirst
        While Not rsPayments.EOF
            CurrRow = getFlexPos(Grid, 1, rsPayments!RecievedPaymentID)
        
            'Add to grid
            With Grid
                If CurrRow < 0 Then
                    'Delete record if doesnt exist in flexgrid
                    DelRecwSQL "Payments", "RecievedPaymentID", "", True, rsPayments!RecievedPaymentID
                End If
            End With
            rsPayments.MoveNext
        Wend
    End If
End Sub

Private Sub txtAmount_GotFocus()
    HLText txtAmount
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtAmount_Validate(Cancel As Boolean)
    txtAmount.Text = toMoney(txtAmount.Text)
End Sub

