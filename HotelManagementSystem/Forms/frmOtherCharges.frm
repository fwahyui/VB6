VERSION 5.00
Object = "{9EDDC69F-10E8-4DE7-9648-EC8A45F005C0}#1.0#0"; "b8Controls4.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmOtherCharges 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Other Charges"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8730
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   8730
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnRemove 
      Height          =   275
      Left            =   300
      Picture         =   "frmOtherCharges.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Remove"
      Top             =   2160
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
      ScaleWidth      =   537
      TabIndex        =   9
      Top             =   0
      Width           =   8055
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
         Caption         =   "Other Charges"
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
         Width           =   2055
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   30
         Picture         =   "frmOtherCharges.frx":01B2
         Top             =   30
         Width           =   480
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   2220
      Left            =   240
      TabIndex        =   21
      Top             =   2010
      Width           =   8205
      _ExtentX        =   14473
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
      Height          =   4485
      Left            =   0
      ScaleHeight     =   299
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   579
      TabIndex        =   12
      Top             =   570
      Width           =   8685
      Begin VB.TextBox txtDescription 
         Height          =   315
         Left            =   3750
         TabIndex        =   2
         Top             =   990
         Width           =   1485
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   375
         Left            =   6480
         TabIndex        =   7
         Top             =   3930
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   375
         Left            =   7740
         TabIndex        =   6
         Top             =   930
         Width           =   795
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   375
         Left            =   7470
         TabIndex        =   8
         Top             =   3930
         Width           =   1005
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   6840
         TabIndex        =   5
         Top             =   930
         Width           =   825
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5310
         TabIndex        =   4
         Text            =   "0.00"
         Top             =   990
         Width           =   1455
      End
      Begin VB.TextBox txtGuestName 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   120
         Width           =   1815
      End
      Begin b8Controls4.b8Line b8Line2 
         Height          =   30
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   53
         BorderColor1    =   15592425
         BorderColor2    =   16777215
      End
      Begin b8Controls4.b8Line b8Line5 
         Height          =   30
         Left            =   150
         TabIndex        =   14
         Top             =   3780
         Width           =   8355
         _ExtentX        =   14737
         _ExtentY        =   53
         BorderColor1    =   15592425
         BorderColor2    =   16777215
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   315
         Left            =   240
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   990
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Format          =   36962305
         CurrentDate     =   39539
      End
      Begin MSDataListLib.DataCombo dcChargeType 
         Height          =   315
         Left            =   1860
         TabIndex        =   1
         Top             =   990
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Left            =   3810
         TabIndex        =   22
         Top             =   750
         Width           =   795
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Left            =   270
         TabIndex        =   19
         Top             =   750
         Width           =   345
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   195
         Left            =   5370
         TabIndex        =   18
         Top             =   720
         Width           =   540
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Charge Type"
         Height          =   195
         Left            =   1890
         TabIndex        =   17
         Top             =   750
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Guest Name"
         Height          =   195
         Left            =   90
         TabIndex        =   16
         Top             =   150
         Width           =   885
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
         TabIndex        =   15
         Top             =   3030
         Width           =   45
      End
   End
End
Attribute VB_Name = "frmOtherCharges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FolioNumber      As String
Public GuestName        As String
Public RefForm          As Form 'Calling form

Dim OtherCharges          As Currency
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
    
    Grid_Click
End Sub

Private Sub cmdAdd_Click()
    If Trim(txtAmount.Text) = "" Or Trim(txtAmount.Text) = "0.00" Then Exit Sub

    Dim CurrRow As Integer

    'Add to grid
    With Grid

        'Perform if the record is not exist
        If .Rows = 2 And .TextMatrix(1, 2) = "" Then
            .TextMatrix(1, 2) = FolioNumber
            .TextMatrix(1, 3) = dtpDate.Value
            .TextMatrix(1, 4) = dcChargeType.Text
            .TextMatrix(1, 5) = txtDescription.Text
            .TextMatrix(1, 6) = toNumber(toMoney(txtAmount.Text))
        Else
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 2) = FolioNumber
            .TextMatrix(.Rows - 1, 3) = dtpDate.Value
            .TextMatrix(.Rows - 1, 4) = dcChargeType.Text
            .TextMatrix(.Rows - 1, 5) = txtDescription.Text
            .TextMatrix(.Rows - 1, 6) = toNumber(toMoney(txtAmount.Text))

            .Row = .Rows - 1
        End If
        'Increase the record count
        cIRowCount = cIRowCount + 1
        
        'Highlight the current row's column
        .ColSel = 6
        'Display a remove button
        Call Grid_Click
        
        Call ResetFields
    End With
End Sub

Private Sub cmdClose_Click()
    Dim c As Integer
    
    With Grid
        'Save the details of the records
        For c = 1 To cIRowCount
            .Row = c
            
            OtherCharges = OtherCharges + .TextMatrix(c, 6)
        Next c
    End With

    'Clear variables
    c = 0
    RefForm.OtherCharges = OtherCharges
    
    Me.Hide
End Sub

Public Sub cmdSave_Click()
    Dim rsOtherCharges As New Recordset

    rsOtherCharges.CursorLocation = adUseClient
    rsOtherCharges.Open "SELECT * FROM [Other Charges] WHERE FolioNumber='" & FolioNumber & "'", CN, adOpenStatic, adLockOptimistic
    
    DeleteItems
    
    Dim c As Integer
    
    With Grid
        'Save the details of the records
        For c = 1 To cIRowCount
            .Row = c
            
            If .TextMatrix(c, 1) = "" Then
                rsOtherCharges.AddNew
                
                rsOtherCharges![FolioNumber] = FolioNumber
            Else
                rsOtherCharges.Filter = "OtherChargesID = " & toNumber(.TextMatrix(c, 1))
            
                If rsOtherCharges.RecordCount = 0 Then
                    rsOtherCharges.AddNew
                    
                    rsOtherCharges![FolioNumber] = FolioNumber
                End If
            End If

            rsOtherCharges![Date] = .TextMatrix(c, 3)
            rsOtherCharges![ChargeTypeID] = getValueAt("SELECT * FROM [Charge Type] WHERE ChargeType = '" & .TextMatrix(c, 4) & "'", "ChargeTypeID")
            rsOtherCharges![Description] = .TextMatrix(c, 5)
            rsOtherCharges![Amount] = .TextMatrix(c, 6)

            rsOtherCharges.Update

            OtherCharges = OtherCharges + .TextMatrix(c, 6)
        Next c
    End With

    'Clear variables
    c = 0
    Set rsOtherCharges = Nothing
    
    Unload frmOtherCharges
End Sub

Private Sub cmdUpdate_Click()
    With Grid
        .TextMatrix(.RowSel, 3) = dtpDate.Value
        .TextMatrix(.RowSel, 4) = dcChargeType.Text
        .TextMatrix(.RowSel, 5) = txtDescription.Text
        .TextMatrix(.RowSel, 6) = toMoney(txtAmount.Text)
    End With
End Sub

Private Sub Form_Activate()
    OtherCharges = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    Call InitGrid
    
    bind_dc "SELECT * FROM [Charge Type]", "ChargeType", dcChargeType, "ChargeTypeID", True

    txtGuestName.Text = GuestName
    
    dtpDate.Value = Date
    
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
        .Cols = 7
        .ColSel = 6
        'Initialize the column size
        .ColWidth(0) = 315
        .ColWidth(1) = 0
        .ColWidth(2) = 0
        .ColWidth(3) = 1200
        .ColWidth(4) = 1500
        .ColWidth(5) = 1700
        .ColWidth(6) = 1200

        'Initialize the column name
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Other Charges ID"
        .TextMatrix(0, 2) = "Folio Number"
        .TextMatrix(0, 3) = "Date"
        .TextMatrix(0, 4) = "Charge Type"
        .TextMatrix(0, 5) = "Description"
        .TextMatrix(0, 6) = "Amount"
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPayment = Nothing
End Sub

Private Sub Grid_Click()
    With Grid
        If .TextMatrix(.RowSel, 3) = "" Then Exit Sub
        
        dtpDate.Value = .TextMatrix(.RowSel, 3)
        dcChargeType.Text = .TextMatrix(.RowSel, 4)
        txtDescription.Text = .TextMatrix(.RowSel, 5)
        txtAmount.Text = .TextMatrix(.RowSel, 6)
    
        If Grid.Rows = 2 And Grid.TextMatrix(1, 2) = "" Then     '1 = Folio Number
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
   
    dcChargeType.SetFocus
End Sub

Private Sub DisplayForEditing()
    On Error GoTo err
    
    'Display the details
    Dim rsOtherCharges As New Recordset

    cIRowCount = 0
    
    rsOtherCharges.CursorLocation = adUseClient
    rsOtherCharges.Open "SELECT * FROM qry_Other_Charges WHERE FolioNumber='" & FolioNumber & "'", CN, adOpenStatic, adLockOptimistic
    
    If rsOtherCharges.RecordCount > 0 Then
        rsOtherCharges.MoveFirst
        While Not rsOtherCharges.EOF
          cIRowCount = cIRowCount + 1     'increment
            With Grid
                If .Rows = 2 And .TextMatrix(1, 1) = "" Then
                    .TextMatrix(1, 1) = rsOtherCharges!OtherChargesID
                    .TextMatrix(1, 2) = rsOtherCharges!FolioNumber
                    .TextMatrix(1, 3) = rsOtherCharges!Date
                    .TextMatrix(1, 4) = rsOtherCharges!ChargeType
                    .TextMatrix(1, 5) = rsOtherCharges!Description
                    .TextMatrix(1, 6) = toMoney(rsOtherCharges!Amount)
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = rsOtherCharges!OtherChargesID
                    .TextMatrix(.Rows - 1, 2) = rsOtherCharges!FolioNumber
                    .TextMatrix(.Rows - 1, 3) = rsOtherCharges!Date
                    .TextMatrix(.Rows - 1, 4) = rsOtherCharges!ChargeType
                    .TextMatrix(.Rows - 1, 5) = rsOtherCharges!Description
                    .TextMatrix(.Rows - 1, 6) = toMoney(rsOtherCharges!Amount)
                End If
            End With
           
            rsOtherCharges.MoveNext
        Wend
        Grid.Row = 1
        Grid.ColSel = 6
        'Set fixed cols
        If State = adStateEditMode Then
            Grid.FixedRows = Grid.Row: 'Grid.SelectionMode = flexSelectionFree
            Grid.FixedCols = 1
        End If
    End If

    rsOtherCharges.Close
    'Clear variables
    Set rsOtherCharges = Nothing

    Exit Sub
err:
    If err.Number = 94 Then Resume Next
    
    prompt_err err, Name, "DisplayForEditing"
    Screen.MousePointer = vbDefault
End Sub

Private Sub DeleteItems()
    Dim CurrRow As Integer
    Dim rsOtherCharges As New Recordset

    rsOtherCharges.CursorLocation = adUseClient
    rsOtherCharges.Open "SELECT * FROM [Other Charges] WHERE FolioNumber='" & FolioNumber & "'", CN, adOpenStatic, adLockOptimistic
    If rsOtherCharges.RecordCount > 0 Then
        rsOtherCharges.MoveFirst
        While Not rsOtherCharges.EOF
            CurrRow = getFlexPos(Grid, 1, rsOtherCharges!OtherChargesID)
        
            'Add to grid
            With Grid
                If CurrRow < 0 Then
                    'Delete record if doesnt exist in flexgrid
                    DelRecwSQL "[Other Charges]", "OtherChargesID", "", True, rsOtherCharges!OtherChargesID
                End If
            End With
            rsOtherCharges.MoveNext
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
