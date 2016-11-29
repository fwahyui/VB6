VERSION 5.00
Object = "{9EDDC69F-10E8-4DE7-9648-EC8A45F005C0}#1.0#0"; "b8Controls4.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmRatePerPeriod 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rate Per Period"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8430
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   8430
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
      ScaleWidth      =   439
      TabIndex        =   11
      Top             =   0
      Width           =   6585
      Begin VB.Image Image1 
         Height          =   480
         Left            =   30
         Picture         =   "frmRatePerPeriod.frx":0000
         Top             =   30
         Width           =   480
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rates"
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
         TabIndex        =   13
         Top             =   30
         Width           =   810
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
         Left            =   570
         TabIndex        =   12
         Top             =   360
         Width           =   3900
      End
   End
   Begin VB.PictureBox bgMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E3F9FB&
      Height          =   4875
      Left            =   0
      ScaleHeight     =   325
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   561
      TabIndex        =   7
      Top             =   570
      Width           =   8415
      Begin VB.TextBox txtChildrens 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4770
         TabIndex        =   3
         Text            =   "0.00"
         Top             =   390
         Width           =   1275
      End
      Begin VB.TextBox txtAdults 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3450
         TabIndex        =   2
         Text            =   "0.00"
         Top             =   390
         Width           =   1275
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   375
         Left            =   6150
         TabIndex        =   5
         Top             =   4350
         Width           =   885
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   315
         Left            =   6120
         TabIndex        =   4
         Top             =   390
         Width           =   795
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   375
         Left            =   7140
         TabIndex        =   6
         Top             =   4350
         Width           =   1005
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2100
         TabIndex        =   1
         Text            =   "0.00"
         Top             =   390
         Width           =   1275
      End
      Begin b8Controls4.b8Line b8Line2 
         Height          =   30
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   53
         BorderColor1    =   15592425
         BorderColor2    =   16777215
      End
      Begin b8Controls4.b8Line b8Line5 
         Height          =   30
         Left            =   210
         TabIndex        =   9
         Top             =   4200
         Width           =   8025
         _ExtentX        =   14155
         _ExtentY        =   53
         BorderColor1    =   15592425
         BorderColor2    =   16777215
      End
      Begin MSDataListLib.DataCombo dcRateType 
         Height          =   315
         Left            =   210
         TabIndex        =   0
         Top             =   390
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
         Height          =   3240
         Left            =   180
         TabIndex        =   14
         Top             =   840
         Width           =   8085
         _ExtentX        =   14261
         _ExtentY        =   5715
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
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Children's Rate"
         Height          =   195
         Left            =   4800
         TabIndex        =   18
         Top             =   150
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Adult's Rate"
         Height          =   195
         Left            =   3480
         TabIndex        =   17
         Top             =   150
         Width           =   855
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
         Height          =   195
         Left            =   2100
         TabIndex        =   16
         Top             =   150
         Width           =   345
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rate Type"
         Height          =   195
         Left            =   210
         TabIndex        =   15
         Top             =   150
         Width           =   750
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
         TabIndex        =   10
         Top             =   3030
         Width           =   45
      End
   End
End
Attribute VB_Name = "frmRatePerPeriod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public FolioNumber      As String
Dim cIRowCount          As Integer

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim rsRatePerPeriod As New Recordset

    rsRatePerPeriod.CursorLocation = adUseClient
    rsRatePerPeriod.Open "SELECT * FROM [Rate Per Period] WHERE FolioNumber='" & FolioNumber & "'", CN, adOpenStatic, adLockOptimistic
    
    Dim c As Integer
    
    With Grid
        'Save the details of the records
        For c = 1 To cIRowCount
            .Row = c
            
            rsRatePerPeriod.Filter = "[Date] = " & .TextMatrix(c, 1)
            
            If rsRatePerPeriod.RecordCount = 1 Then
                rsRatePerPeriod![RateTypeID] = .TextMatrix(c, 3)
                rsRatePerPeriod![Rate] = .TextMatrix(c, 5)
                rsRatePerPeriod![Adults] = .TextMatrix(c, 6)
                rsRatePerPeriod![Childrens] = .TextMatrix(c, 7)
    
                rsRatePerPeriod.Update
            End If
        Next c
    End With

    'Clear variables
    c = 0
    Set rsRatePerPeriod = Nothing
    
    Unload frmRatePerPeriod
End Sub

Private Sub cmdUpdate_Click()
    With Grid
        .TextMatrix(.RowSel, 3) = dcRateType.BoundText
        .TextMatrix(.RowSel, 4) = dcRateType.Text
        .TextMatrix(.RowSel, 5) = toMoney(txtRate.Text)
        .TextMatrix(.RowSel, 6) = toMoney(txtAdults.Text)
        .TextMatrix(.RowSel, 7) = toMoney(txtChildrens.Text)
    End With
End Sub

Private Sub dcRateType_Click(Area As Integer)
    Dim rsRoomRates As New ADODB.Recordset
    
    If dcRateType.BoundText = "" Then Exit Sub
    
    With rsRoomRates
        .Open "SELECT * FROM [Room Rates] WHERE RoomNumber = " & Grid.TextMatrix(Grid.RowSel, 2) & " AND RateTypeID = " & dcRateType.BoundText, CN, adOpenStatic, adLockOptimistic
    
        If .RecordCount > 0 Then
            txtRate.Text = toMoney(!RoomRate)
        End If
    End With
    
    rsRoomRates.Close
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    Call InitGrid
    
    bind_dc "SELECT * FROM [Rate Type]", "RateType", dcRateType, "RateTypeID", False

    DisplayForEditing
    
    DoEvents
End Sub

Private Sub DisplayForEditing()
    On Error GoTo err
    
    'Display the details
    Dim rsRatePerPeriod As New Recordset

    cIRowCount = 0
    
    rsRatePerPeriod.CursorLocation = adUseClient
    rsRatePerPeriod.Open "SELECT * FROM qry_Rate_Per_Period WHERE FolioNumber='" & FolioNumber & "' ORDER BY [Date]", CN, adOpenStatic, adLockOptimistic
    
    If rsRatePerPeriod.RecordCount > 0 Then
        rsRatePerPeriod.MoveFirst
        While Not rsRatePerPeriod.EOF
          cIRowCount = cIRowCount + 1     'increment
            With Grid
                If .Rows = 2 And .TextMatrix(1, 1) = "" Then
                    .TextMatrix(1, 1) = rsRatePerPeriod!Date
                    .TextMatrix(1, 2) = rsRatePerPeriod!RoomNumber
                    .TextMatrix(1, 3) = rsRatePerPeriod!RateTypeID
                    .TextMatrix(1, 4) = rsRatePerPeriod!RateType
                    .TextMatrix(1, 5) = toMoney(rsRatePerPeriod!Rate)
                    .TextMatrix(1, 6) = toMoney(rsRatePerPeriod!Adults)
                    .TextMatrix(1, 7) = toMoney(rsRatePerPeriod!Childrens)
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = rsRatePerPeriod!Date
                    .TextMatrix(.Rows - 1, 2) = rsRatePerPeriod!RoomNumber
                    .TextMatrix(.Rows - 1, 3) = rsRatePerPeriod!RateTypeID
                    .TextMatrix(.Rows - 1, 4) = rsRatePerPeriod!RateType
                    .TextMatrix(.Rows - 1, 5) = toMoney(rsRatePerPeriod!Rate)
                    .TextMatrix(.Rows - 1, 6) = toMoney(rsRatePerPeriod!Adults)
                    .TextMatrix(.Rows - 1, 7) = toMoney(rsRatePerPeriod!Childrens)
                End If
            End With
            
            rsRatePerPeriod.MoveNext
        Wend
        Grid.Row = 1
        Grid.ColSel = 7
        'Set fixed cols
'        If State = adStateEditMode Then
            Grid.FixedRows = Grid.Row: 'Grid.SelectionMode = flexSelectionFree
            Grid.FixedCols = 1
'        End If
    End If

    rsRatePerPeriod.Close
    'Clear variables
    Set rsRatePerPeriod = Nothing

    Exit Sub
err:
    If err.Number = 94 Then Resume Next
    
    prompt_err err, Name, "DisplayForEditing"
    Screen.MousePointer = vbDefault
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
        .Cols = 8
        .ColSel = 7
        'Initialize the column size
        .ColWidth(0) = 315
        .ColWidth(1) = 1200
        .ColWidth(2) = 1200
        .ColWidth(3) = 0
        .ColWidth(4) = 1200
        .ColWidth(5) = 900
        .ColWidth(6) = 1200
        .ColWidth(7) = 1200
        
        'Initialize the column name
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Date"
        .TextMatrix(0, 2) = "Room Number"
        .TextMatrix(0, 3) = "Rate Type ID"
        .TextMatrix(0, 4) = "Rate Type"
        .TextMatrix(0, 5) = "Rate"
        .TextMatrix(0, 6) = "Adult's Rate"
        .TextMatrix(0, 7) = "Children's Rate"
        
        'Set the column alignment
'        .ColAlignment(0) = vbLeftJustify
'        .ColAlignment(1) = vbLeftJustify
'        .ColAlignment(2) = vbLeftJustify
'        .ColAlignment(3) = flexAlignGeneral
'        .ColAlignment(4) = flexAlignGeneral
'        .ColAlignment(5) = vbRightJustify
'        .ColAlignment(6) = vbRightJustify
'        .ColAlignment(7) = vbRightJustify
'        .ColAlignment(8) = vbRightJustify
    End With
End Sub

Private Sub Grid_Click()
    With Grid
        dcRateType.BoundText = .TextMatrix(.RowSel, 3)
        txtRate.Text = .TextMatrix(.RowSel, 5)
        txtAdults.Text = .TextMatrix(.RowSel, 6)
        txtChildrens.Text = .TextMatrix(.RowSel, 7)
    End With
End Sub

Private Sub txtAdults_GotFocus()
    HLText txtAdults
End Sub

Private Sub txtAdults_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtAdults_Validate(Cancel As Boolean)
    txtAdults.Text = toMoney(txtAdults.Text)
End Sub

Private Sub txtChildrens_GotFocus()
    HLText txtChildrens
End Sub

Private Sub txtChildrens_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtChildrens_Validate(Cancel As Boolean)
    txtChildrens.Text = toMoney(txtChildrens.Text)
End Sub

Private Sub txtRate_GotFocus()
    HLText txtRate
End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtRate_Validate(Cancel As Boolean)
    txtRate.Text = toMoney(txtRate.Text)
End Sub
