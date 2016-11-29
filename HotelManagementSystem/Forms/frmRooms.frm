VERSION 5.00
Object = "{9EDDC69F-10E8-4DE7-9648-EC8A45F005C0}#1.0#0"; "b8Controls4.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmRooms 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rooms"
   ClientHeight    =   6120
   ClientLeft      =   105
   ClientTop       =   405
   ClientWidth     =   8910
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   8910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnRemove 
      Height          =   275
      Left            =   210
      Picture         =   "frmRooms.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Remove"
      Top             =   3180
      Visible         =   0   'False
      Width           =   275
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   2220
      Left            =   150
      TabIndex        =   15
      Top             =   3000
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
      TabIndex        =   11
      Top             =   0
      Width           =   8715
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
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rooms"
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
         TabIndex        =   12
         Top             =   30
         Width           =   990
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   30
         Picture         =   "frmRooms.frx":01B2
         Top             =   30
         Width           =   480
      End
   End
   Begin VB.PictureBox bgMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E3F9FB&
      Height          =   5535
      Left            =   0
      ScaleHeight     =   369
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   593
      TabIndex        =   16
      Top             =   570
      Width           =   8895
      Begin b8Controls4.b8Line b8Line5 
         Height          =   30
         Left            =   180
         TabIndex        =   18
         Top             =   4800
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   53
         BorderColor1    =   15592425
         BorderColor2    =   16777215
      End
      Begin VB.TextBox txtRoomNumber 
         DataField       =   "RoomNumber"
         DataMember      =   "Rooms"
         DataSource      =   "deHotel"
         Height          =   285
         Left            =   1650
         TabIndex        =   0
         Top             =   150
         Width           =   960
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   7800
         TabIndex        =   10
         Top             =   4950
         Width           =   885
      End
      Begin VB.TextBox txtRoomRate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2040
         TabIndex        =   4
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   375
         Left            =   7710
         TabIndex        =   8
         Top             =   2010
         Width           =   795
      End
      Begin VB.TextBox txtAdults 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5100
         TabIndex        =   6
         Top             =   2040
         Width           =   1245
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   375
         Left            =   6630
         TabIndex        =   9
         Top             =   4950
         Width           =   885
      End
      Begin VB.TextBox txtChildrens 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6420
         TabIndex        =   7
         Top             =   2040
         Width           =   1245
      End
      Begin VB.TextBox txtNoofPerson 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3570
         TabIndex        =   5
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CommandButton cmdUsrHistory 
         Caption         =   "Modification History"
         Height          =   315
         Left            =   240
         TabIndex        =   20
         Top             =   4950
         Width           =   1680
      End
      Begin b8Controls4.b8Line b8Line2 
         Height          =   30
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   53
         BorderColor1    =   15592425
         BorderColor2    =   16777215
      End
      Begin MSDataListLib.DataCombo dcRoomType 
         Bindings        =   "frmRooms.frx":0A7C
         Height          =   315
         Left            =   1650
         TabIndex        =   1
         Top             =   510
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   ""
         BoundColumn     =   "RoomTypeID"
         Text            =   ""
         Object.DataMember      =   ""
      End
      Begin MSDataListLib.DataCombo dcRoomStatus 
         Bindings        =   "frmRooms.frx":0A87
         DataField       =   "RoomStatusID"
         Height          =   315
         Left            =   1650
         TabIndex        =   2
         Top             =   915
         Width           =   2430
         _ExtentX        =   4286
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin b8Controls4.b8GradLine b8GradLine2 
         Height          =   240
         Left            =   150
         TabIndex        =   21
         Top             =   1380
         Width           =   8505
         _ExtentX        =   15002
         _ExtentY        =   423
         Color1          =   14737632
         Color2          =   16119285
         Caption         =   "Rate Details"
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
      Begin MSDataListLib.DataCombo dcRateType 
         Height          =   315
         Left            =   150
         TabIndex        =   3
         Top             =   2040
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Room Number"
         Height          =   195
         Index           =   1
         Left            =   585
         TabIndex        =   29
         Top             =   165
         Width           =   1020
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Room Type"
         Height          =   195
         Index           =   2
         Left            =   780
         TabIndex        =   28
         Top             =   555
         Width           =   825
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Room Status"
         Height          =   195
         Index           =   3
         Left            =   690
         TabIndex        =   27
         Top             =   960
         Width           =   915
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rate Type"
         Height          =   195
         Left            =   420
         TabIndex        =   26
         Top             =   1800
         Width           =   750
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Room Rate"
         Height          =   195
         Left            =   2160
         TabIndex        =   25
         Top             =   1770
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Adult's Rate"
         Height          =   195
         Left            =   5100
         TabIndex        =   24
         Top             =   1770
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Children's Rate"
         Height          =   195
         Left            =   6390
         TabIndex        =   23
         Top             =   1770
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. of Person"
         Height          =   195
         Left            =   3600
         TabIndex        =   22
         Top             =   1770
         Width           =   975
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
         TabIndex        =   19
         Top             =   3030
         Width           =   45
      End
   End
End
Attribute VB_Name = "frmRooms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public State                As FormState 'Variable used to determine on how the form used
Public PK                   As Long 'Variable used to get what record is going to edit

Dim RS                      As New Recordset
Dim cIRowCount              As Integer

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

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Public Sub cmdSave_Click()
On Error GoTo err

    Dim rsRoomRates As New Recordset

    CN.BeginTrans
    
    rsRoomRates.CursorLocation = adUseClient
    rsRoomRates.Open "SELECT * FROM [Room Rates] WHERE RoomNumber=" & PK, CN, adOpenStatic, adLockOptimistic

    If State = adStateAddMode Then
        If txtRoomNumber.Text = "" Then txtRoomNumber.SetFocus: Exit Sub
        
        RS.AddNew
        
        RS.Fields("RoomNumber") = txtRoomNumber.Text
        RS.Fields("AddedByFK") = CurrUser.USER_PK
    Else
        RS.Fields("DateModified") = Now
        RS.Fields("LastUserFK") = CurrUser.USER_PK
    End If

    With RS
        .Fields("RoomNumber") = txtRoomNumber.Text
        .Fields("RoomTypeID") = dcRoomType.BoundText
        .Fields("RoomStatusID") = dcRoomStatus.BoundText
        
        .Update
    End With
    
    Dim c As Integer
    
    With Grid
        'Save the details of the records
        For c = 1 To cIRowCount
            .Row = c
            
            If State = adStateAddMode Then
                rsRoomRates.AddNew
                
                rsRoomRates![RoomNumber] = txtRoomNumber.Text
            Else
                rsRoomRates.Filter = "[RateTypeID]=" & .TextMatrix(c, 1)
            End If

            If rsRoomRates.RecordCount > 0 Then
                rsRoomRates![RateTypeID] = .TextMatrix(c, 1)
                rsRoomRates![RoomRate] = .TextMatrix(c, 3)
                rsRoomRates![NoofPerson] = .TextMatrix(c, 4)
                rsRoomRates![ExtraAdultRate] = .TextMatrix(c, 5)
                rsRoomRates![ExtraChildRate] = .TextMatrix(c, 6)

                rsRoomRates.Update
            End If
        Next c
    End With

    'Save last rate entry to rate templates
    If State = adStateAddMode Then
        CN.Execute "DELETE RoomTypeID " & _
                    "From [Rate Templates] " & _
                    "WHERE RoomTypeID=" & dcRoomType.BoundText
                    
        CN.Execute "INSERT INTO [Rate Templates] ( RoomTypeID, RateTypeID, RoomRate, NoofPerson, ExtraAdultRate ) " & _
                    "SELECT Rooms.RoomTypeID, [Room Rates].RateTypeID, [Room Rates].RoomRate, [Room Rates].NoofPerson, [Room Rates].ExtraAdultRate " & _
                    "FROM [Room Rates] INNER JOIN Rooms ON [Room Rates].RoomNumber = Rooms.RoomNumber " & _
                    "WHERE [Room Rates].RoomNumber=" & txtRoomNumber.Text
    End If

    'Clear variables
    c = 0
    Set rsRoomRates = Nothing
    
    CN.CommitTrans
    
    Unload frmRooms

    Exit Sub
    
err:
    CN.RollbackTrans
    prompt_err err, Name, "cmdSave_Click"
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdUpdate_Click()
    With Grid
        .TextMatrix(.RowSel, 1) = dcRateType.BoundText
        .TextMatrix(.RowSel, 2) = dcRateType.Text
        .TextMatrix(.RowSel, 3) = toMoney(txtRoomRate.Text)
        .TextMatrix(.RowSel, 4) = txtNoofPerson.Text
        .TextMatrix(.RowSel, 5) = toMoney(txtAdults.Text)
        .TextMatrix(.RowSel, 6) = toMoney(txtChildrens.Text)
    End With
End Sub

Private Sub AddRoomRates()
On Error GoTo err

    CN.BeginTrans
    
    If State = adStateAddMode Then
        CN.Execute "INSERT INTO [Room Rates] ( RoomNumber, RateTypeID ) " & _
                    "SELECT " & txtRoomNumber.Text & ", [Rate Type].RateTypeID " & _
                    "FROM [Rate Type]"
    Else
        Dim rsRateType As New Recordset
        
        With rsRateType
            .Open "SELECT RateTypeID FROM [Rate Type]", CN, adOpenStatic, adLockOptimistic
            
            Do While Not .EOF
                If .Fields("RateTypeID") <> getValueAt("SELECT RateTypeID FROM [Room Rates] WHERE RoomNumber = " & txtRoomNumber.Text & " AND RateTypeID = " & .Fields("RateTypeID"), "RateTypeID") Then
                    CN.Execute "INSERT INTO [Room Rates] ( RoomNumber, RateTypeID ) " & _
                                "SELECT " & txtRoomNumber.Text & ", " & .Fields("RateTypeID")
                    
                    .Update
                End If
                
                .MoveNext
            Loop
            .Close
        End With
    End If
    
    CN.CommitTrans
    
    Exit Sub
    
err:
    CN.RollbackTrans
    prompt_err err, Name, "AddRoomRates"
    Screen.MousePointer = vbDefault
End Sub

Private Sub dcRoomType_Click(Area As Integer)
    AddFromRateTemplates
'     Call AddRoomRates
'     Call DisplayForEditing
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
On Error GoTo err

    Call InitGrid
    
    bind_dc "SELECT * FROM [Room Type]", "RoomType", dcRoomType, "RoomTypeID", True
    bind_dc "SELECT * FROM [Rate Type]", "RateType", dcRateType, "RateTypeID", True
    bind_dc "SELECT * FROM [Room Status]", "Status", dcRoomStatus, "StatusID", True

    RS.Open "SELECT * FROM Rooms WHERE RoomNumber = " & PK, CN, adOpenStatic, adLockOptimistic
    
    'Check the form state
    If State = adStateAddMode Then
        AddFromRateTemplates

        cmdUsrHistory.Enabled = False
    Else
        With RS
            txtRoomNumber.Text = PK
            dcRoomType.BoundText = .Fields("RoomTypeID")
            dcRoomStatus.BoundText = .Fields("RoomStatusID")
        End With
        
        Call AddRoomRates
        
        DisplayForEditing
    End If
    
    Exit Sub
    
err:
    CN.RollbackTrans
    prompt_err err, Name, "cmdSave_Click"
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
        .Cols = 7
        .ColSel = 6
        'Initialize the column size
        .ColWidth(0) = 315
        .ColWidth(1) = 0
        .ColWidth(2) = 1200
        .ColWidth(3) = 1200
        .ColWidth(4) = 1200
        .ColWidth(5) = 1200
        .ColWidth(6) = 1200

        'Initialize the column name
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Rate Type ID"
        .TextMatrix(0, 2) = "Rate Type"
        .TextMatrix(0, 3) = "Room Rate"
        .TextMatrix(0, 4) = "No. of Person"
        .TextMatrix(0, 5) = "Extra Adult's Rate"
        .TextMatrix(0, 6) = "Extra Children's Rate"
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmRoomsList.RefreshRecords
    
    Set frmRooms = Nothing
End Sub

Private Sub Grid_Click()
    With Grid
        dcRateType.BoundText = .TextMatrix(.RowSel, 1)
        txtRoomRate.Text = .TextMatrix(.RowSel, 3)
        txtNoofPerson.Text = .TextMatrix(.RowSel, 4)
        txtAdults.Text = .TextMatrix(.RowSel, 5)
        txtChildrens.Text = .TextMatrix(.RowSel, 6)
    
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
    txtRoomNumber.Text = ""
    
    txtRoomNumber.SetFocus
End Sub

Private Sub DisplayForEditing()
On Error GoTo err
    
    'Display the details
    Dim rsRoomRates As New Recordset

    cIRowCount = 0
    
    rsRoomRates.CursorLocation = adUseClient
    rsRoomRates.Open "SELECT * FROM qry_Room_Rates WHERE RoomNumber=" & PK, CN, adOpenStatic, adLockOptimistic
    
    If rsRoomRates.RecordCount > 0 Then
        rsRoomRates.MoveFirst
        While Not rsRoomRates.EOF
          cIRowCount = cIRowCount + 1     'increment
            With Grid
                If .Rows = 2 And .TextMatrix(1, 1) = "" Then
                    .TextMatrix(1, 1) = rsRoomRates!RateTypeID
                    .TextMatrix(1, 2) = rsRoomRates!RateType
                    .TextMatrix(1, 3) = toMoney(rsRoomRates!RoomRate)
                    .TextMatrix(1, 4) = rsRoomRates!NoofPerson
                    .TextMatrix(1, 5) = toMoney(rsRoomRates!ExtraAdultRate)
                    .TextMatrix(1, 6) = toMoney(rsRoomRates!ExtraChildRate)
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = rsRoomRates!RateTypeID
                    .TextMatrix(.Rows - 1, 2) = rsRoomRates!RateType
                    .TextMatrix(.Rows - 1, 3) = toMoney(rsRoomRates!RoomRate)
                    .TextMatrix(.Rows - 1, 4) = rsRoomRates!NoofPerson
                    .TextMatrix(.Rows - 1, 5) = toMoney(rsRoomRates!ExtraAdultRate)
                    .TextMatrix(.Rows - 1, 6) = toMoney(rsRoomRates!ExtraChildRate)
                End If
            End With
           
            rsRoomRates.MoveNext
        Wend
        Grid.Row = 1
        Grid.ColSel = 6
        'Set fixed cols
        If State = adStateEditMode Then
            Grid.FixedRows = Grid.Row: 'Grid.SelectionMode = flexSelectionFree
            Grid.FixedCols = 1
        End If
    End If

    rsRoomRates.Close
    'Clear variables
    Set rsRoomRates = Nothing

    Exit Sub
err:
    If err.Number = 94 Then Resume Next
    
    prompt_err err, Name, "DisplayForEditing"
    Screen.MousePointer = vbDefault
End Sub

Private Sub AddFromRateTemplates()
    Dim rsRateType As New Recordset
    Dim rsRateTemplates As New Recordset

    Grid.Clear
    InitGrid
    
    cIRowCount = 0
    
    rsRateTemplates.CursorLocation = adUseClient
    rsRateTemplates.Open "SELECT [Rate Templates].RoomTypeID, [Rate Templates].RateTypeID, [Rate Templates].RoomRate, [Rate Templates].NoofPerson, [Rate Templates].ExtraAdultRate " & _
                    "From [Rate Templates] " & _
                    "WHERE RoomTypeID= " & dcRoomType.BoundText, CN, adOpenStatic, adLockOptimistic
    
    rsRateType.CursorLocation = adUseClient
    rsRateType.Open "SELECT RateTypeID, RateType FROM [Rate Type]", CN, adOpenStatic, adLockOptimistic
    
    If rsRateType.RecordCount > 0 Then
        rsRateType.MoveFirst
        While Not rsRateType.EOF
          cIRowCount = cIRowCount + 1     'increment
          rsRateTemplates.Filter = "[RateTypeID]=" & rsRateType!RateTypeID
            With Grid
                If .Rows = 2 And .TextMatrix(1, 1) = "" Then
                    .TextMatrix(1, 1) = rsRateType!RateTypeID
                    .TextMatrix(1, 2) = rsRateType!RateType
                    If rsRateTemplates.RecordCount > 0 Then
                        .TextMatrix(1, 3) = toMoney(rsRateTemplates!RoomRate)
                        .TextMatrix(1, 4) = rsRateTemplates!NoofPerson
                        .TextMatrix(1, 5) = toMoney(rsRateTemplates!ExtraAdultRate)
                    End If
                    .TextMatrix(1, 6) = 0
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = rsRateType!RateTypeID
                    .TextMatrix(.Rows - 1, 2) = rsRateType!RateType
                    If rsRateTemplates.RecordCount > 0 Then
                        .TextMatrix(.Rows - 1, 3) = toMoney(rsRateTemplates!RoomRate)
                        .TextMatrix(.Rows - 1, 4) = rsRateTemplates!NoofPerson
                        .TextMatrix(.Rows - 1, 5) = toMoney(rsRateTemplates!ExtraAdultRate)
                    End If
                    .TextMatrix(.Rows - 1, 6) = 0
                End If
            End With
           
            rsRateType.MoveNext
        Wend
        Grid.Row = 1
        Grid.ColSel = 6
        'Set fixed cols
        If State = adStateEditMode Then
            Grid.FixedRows = Grid.Row: 'Grid.SelectionMode = flexSelectionFree
            Grid.FixedCols = 1
        End If
    End If
    
    rsRateType.Close
    rsRateTemplates.Close
    
    Set rsRateType = Nothing
    Set rsRateTemplates = Nothing
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

Private Sub txtNoofPerson_GotFocus()
    HLText txtNoofPerson
End Sub

Private Sub txtNoofPerson_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtNoofPerson_Validate(Cancel As Boolean)
    txtNoofPerson.Text = toNumber(txtNoofPerson.Text)
End Sub

Private Sub txtRoomRate_GotFocus()
    HLText txtRoomRate
End Sub

Private Sub txtRoomRate_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtRoomRate_Validate(Cancel As Boolean)
    txtRoomRate.Text = toMoney(txtRoomRate.Text)
End Sub
