VERSION 5.00
Object = "{9EDDC69F-10E8-4DE7-9648-EC8A45F005C0}#1.0#0"; "b8Controls4.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCheckOut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check Out"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8130
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   8130
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox bgMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E3F9FB&
      Height          =   6075
      Left            =   0
      ScaleHeight     =   405
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   541
      TabIndex        =   21
      Top             =   540
      Width           =   8115
      Begin VB.TextBox txtGuestName 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1620
         TabIndex        =   1
         Top             =   630
         Width           =   3015
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   5040
         TabIndex        =   13
         Text            =   "0.00"
         Top             =   3930
         Width           =   1815
      End
      Begin VB.TextBox txtSubTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   5040
         TabIndex        =   11
         Text            =   "0.00"
         Top             =   3090
         Width           =   1815
      End
      Begin VB.TextBox txtOtherCharges 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   5040
         TabIndex        =   10
         Text            =   "0.00"
         Top             =   2670
         Width           =   1815
      End
      Begin VB.TextBox txtRoomNumber 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   1620
         TabIndex        =   0
         Top             =   210
         Width           =   825
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   5070
         TabIndex        =   8
         Text            =   "0.00"
         Top             =   1830
         Width           =   1815
      End
      Begin VB.TextBox txtDiscount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   5040
         TabIndex        =   12
         Text            =   "0.00"
         Top             =   3510
         Width           =   1575
      End
      Begin VB.TextBox txtDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   1620
         TabIndex        =   4
         Text            =   "1"
         Top             =   2250
         Width           =   465
      End
      Begin VB.TextBox txtAdults 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   1620
         TabIndex        =   5
         Text            =   "1"
         Top             =   2670
         Width           =   465
      End
      Begin VB.TextBox txtChildrens 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   1620
         TabIndex        =   6
         Text            =   "0"
         Top             =   3090
         Width           =   465
      End
      Begin VB.TextBox txtTotalCharges 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   5040
         TabIndex        =   9
         Text            =   "0.00"
         Top             =   2250
         Width           =   1815
      End
      Begin VB.TextBox txtAmountPaid 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   5040
         TabIndex        =   14
         Text            =   "0.00"
         Top             =   4350
         Width           =   1815
      End
      Begin VB.TextBox txtBalance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   5040
         TabIndex        =   15
         Text            =   "0.00"
         Top             =   4770
         Width           =   1815
      End
      Begin VB.TextBox txtDateIn 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   1620
         TabIndex        =   2
         Top             =   1410
         Width           =   1245
      End
      Begin VB.CommandButton cmdCheckOut 
         Appearance      =   0  'Flat
         Caption         =   "Check Out"
         Height          =   315
         Left            =   4860
         TabIndex        =   16
         Top             =   5520
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "Cancel"
         Height          =   315
         Left            =   6150
         TabIndex        =   17
         Top             =   5520
         Width           =   1335
      End
      Begin b8Controls4.b8Line b8Line2 
         Height          =   30
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   53
         BorderColor1    =   15592425
         BorderColor2    =   16777215
      End
      Begin b8Controls4.b8Line b8Line5 
         Height          =   30
         Left            =   240
         TabIndex        =   23
         Top             =   5280
         Width           =   7425
         _ExtentX        =   13097
         _ExtentY        =   53
         BorderColor1    =   15592425
         BorderColor2    =   16777215
      End
      Begin MSComCtl2.DTPicker dtpDateOut 
         Height          =   345
         Left            =   5040
         TabIndex        =   7
         Top             =   1380
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   609
         _Version        =   393216
         Format          =   77266945
         CurrentDate     =   39536
      End
      Begin MSDataListLib.DataCombo dcRateType 
         Height          =   315
         Left            =   1620
         TabIndex        =   3
         Top             =   1860
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   315
         Left            =   6630
         TabIndex        =   41
         Top             =   3540
         Width           =   165
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Guest Name"
         Height          =   195
         Left            =   690
         TabIndex        =   40
         Top             =   660
         Width           =   885
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3600
         TabIndex        =   39
         Top             =   3960
         Width           =   1395
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Balance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3600
         TabIndex        =   38
         Top             =   4800
         Width           =   1395
      End
      Begin VB.Label lblAmountPaid 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Paid"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   3600
         TabIndex        =   37
         Top             =   4410
         Width           =   1395
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3600
         TabIndex        =   36
         Top             =   3120
         Width           =   1395
      End
      Begin VB.Label lblOtherCharges 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Other Charges"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   3600
         TabIndex        =   35
         Top             =   2700
         Width           =   1395
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  'Dot
         X1              =   248
         X2              =   248
         Y1              =   90
         Y2              =   340
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Room Number"
         Height          =   300
         Left            =   180
         TabIndex        =   34
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Date In"
         Height          =   300
         Left            =   180
         TabIndex        =   33
         Top             =   1440
         Width           =   1395
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Date Out"
         Height          =   300
         Left            =   3600
         TabIndex        =   32
         Top             =   1410
         Width           =   1395
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Rate Type"
         Height          =   300
         Left            =   180
         TabIndex        =   31
         Top             =   1860
         Width           =   1395
      End
      Begin VB.Label lblRatePerPeriod 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Rate/Period"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   3600
         TabIndex        =   30
         Top             =   1860
         Width           =   1395
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
         Height          =   300
         Left            =   3600
         TabIndex        =   29
         Top             =   3540
         Width           =   1395
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. of Days"
         Height          =   300
         Left            =   180
         TabIndex        =   28
         Top             =   2280
         Width           =   1395
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. of Adults"
         Height          =   300
         Left            =   180
         TabIndex        =   27
         Top             =   2700
         Width           =   1395
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. of Childrens"
         Height          =   300
         Left            =   180
         TabIndex        =   26
         Top             =   3120
         Width           =   1395
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Charges"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3600
         TabIndex        =   25
         Top             =   2280
         Width           =   1395
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
         TabIndex        =   24
         Top             =   3030
         Width           =   45
      End
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
      ScaleWidth      =   479
      TabIndex        =   18
      Top             =   0
      Width           =   7185
      Begin VB.Image Image1 
         Height          =   480
         Left            =   30
         Picture         =   "frmCheckOut.frx":0000
         Top             =   30
         Width           =   480
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Check Out"
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
         TabIndex        =   20
         Top             =   30
         Width           =   1470
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
         TabIndex        =   19
         Top             =   360
         Width           =   3900
      End
   End
End
Attribute VB_Name = "frmCheckOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RoomNumber           As Integer
Public AmountPaid           As Currency 'Amount paid from frmPayment
Public OtherCharges         As Currency
Public AutoCheckOut         As Boolean

Dim RS                      As New Recordset

Private Sub cmdCancel_Click()
On Error GoTo err

    CN.BeginTrans
    
    CN.Execute "DELETE FolioNumber " & _
                "From [Rate Per Period] " & _
                "WHERE FolioNumber='" & txtGuestName.Tag & "'"
                
    CN.Execute "INSERT INTO [Rate Per Period] " & _
                "SELECT [Rate Per Period Temp].* " & _
                "FROM [Rate Per Period Temp] " & _
                "Where ((([Rate Per Period Temp].FolioNumber) = '" & txtGuestName.Tag & "')) " & _
                "ORDER BY [Rate Per Period Temp].Date;"
    
    CN.CommitTrans
    
    Unload Me

    Exit Sub
err:
    CN.RollbackTrans
    prompt_err err, Name, "cmdCancel_Click"
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdCheckOut_Click()
    On Error GoTo err
    
    If txtBalance.Text <> "0.00" Then
        MsgBox "There's still remaining balance for this guest.", vbExclamation
        
        Exit Sub
    End If
    
    If MsgBox("Are you sure you want to Check Out?", vbYesNo + vbInformation) = vbNo Then Exit Sub
    
    CN.BeginTrans
    
    ChangeValue CN, "Rooms", "RoomStatusID", 3, True, "WHERE RoomNumber = " & txtRoomNumber.Text

    Call frmPayment.cmdSave_Click
    Call frmOtherCharges.cmdSave_Click

    With RS
        'Delete record from Inventory
        CN.Execute "DELETE ID, Status " & _
                    "From [Inventory] " & _
                    "WHERE ID='" & .Fields("FolioNumber") & "' AND Status='Check In'"
        
        .Fields("DateOut") = dtpDateOut.Value
        .Fields("OtherCharges") = txtOtherCharges.Text
        .Fields("Discount") = txtDiscount.Text
        .Fields("AmountPaid") = txtAmountPaid.Text
        .Fields("Days") = txtDays.Text
        .Fields("Status") = "Check Out"
        .Fields("CheckOutBy") = CurrUser.USER_PK
        
        .Update
    End With
    
    CN.CommitTrans
    
    Call PrintFolio
    
    RS.Close
    Set RS = Nothing
    
    Unload Me
    
    Exit Sub

err:
    CN.RollbackTrans
    prompt_err err, Name, "cmdCheckOut_Click"
    Screen.MousePointer = vbDefault
End Sub

Private Sub PrintFolio()
    With frmReports
        .strReport = "Folio"
        
        .strWhere = "{qry_RPT_Customers.FolioNumber} = '" & txtGuestName.Tag & "' AND {qry_RPT_Customers.Status} = 'Check Out'"

        frmReports.Show vbModal
    End With
End Sub

Private Sub dtpDateOut_Change()
    txtDays.Text = dtpDateOut.Value - CDate(txtDateIn.Text)
    
    Call ComputeRate
End Sub

Private Sub dtpDateOut_LostFocus()
    If CDate(txtDateIn.Text) > dtpDateOut.Value Then
        MsgBox "Check In date must be below check out date. Please enter another check out date.", vbInformation
        
        dtpDateOut.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
On Error GoTo err

    CN.BeginTrans

    RS.CursorLocation = adUseClient
    RS.Open "SELECT * FROM Transactions WHERE RoomNumber = " & RoomNumber & " AND Status = 'Check In'", CN, adOpenStatic, adLockOptimistic

    bind_dc "SELECT * FROM [Rate Type]", "RateType", dcRateType, "RateTypeID", True

    txtRoomNumber.Text = RoomNumber
    
    With RS
        txtGuestName.Tag = .Fields("FolioNumber")
        txtGuestName.Text = getValueAt("SELECT [Name] FROM qry_CheckIn WHERE FolioNumber = '" & .Fields("FolioNumber") & " '", "Name")
        txtDateIn.Text = .Fields("DateIn")
        If AutoCheckOut = True Then
            If .Fields("DateOut") >= Date Then
                dtpDateOut.Value = .Fields("DateOut")
            Else
                dtpDateOut.Value = Date
            End If
        Else
            dtpDateOut.Value = .Fields("DateOut")
        End If
        dcRateType.BoundText = .Fields("RateType")
        txtDays.Text = dtpDateOut.Value - CDate(txtDateIn.Text)
        txtAdults.Text = .Fields("Adults")
        txtChildrens.Text = .Fields("Childrens")
        txtRate.Text = toMoney(.Fields("Rate"))
        txtOtherCharges.Text = toMoney(.Fields("OtherCharges"))
        txtDiscount.Text = toMoney(.Fields("Discount"))
        txtAmountPaid.Text = toMoney(.Fields("AmountPaid"))
    End With
    
    dcRateType.Enabled = False
    
    Call ComputeAddRate
    Call ComputeRate

    
    CN.Execute "DELETE FolioNumber " & _
                "From [Rate Per Period Temp] " & _
                "WHERE FolioNumber='" & txtGuestName.Tag & "'"

    CN.Execute "INSERT INTO [Rate Per Period Temp] " & _
                "SELECT [Rate Per Period].* " & _
                "From [Rate Per Period] " & _
                "WHERE FolioNumber='" & txtGuestName.Tag & "'"
                
    CN.CommitTrans
    
    Exit Sub

err:
    CN.RollbackTrans
    prompt_err err, Name, "txtDays_Change"
    Screen.MousePointer = vbDefault
End Sub

Private Sub ComputeRate()
    txtTotalCharges.Text = toMoney(ComputeRatePerPeriod)
    txtSubTotal.Text = toMoney(toNumber(txtTotalCharges.Text) + toNumber(txtOtherCharges.Text))
    txtTotal.Text = toMoney(toNumber(txtSubTotal.Text) - (toNumber(txtSubTotal.Text) * toNumber(txtDiscount.Text) / 100))
    txtBalance.Text = toMoney(toNumber(txtTotal.Text) - toNumber(txtAmountPaid.Text))
End Sub

'Compute additional rate (no. of days & childrens)
Private Sub ComputeAddRate()
    Dim rsRoomRates As New ADODB.Recordset
    
    With rsRoomRates
        .Open "SELECT * FROM [Room Rates] WHERE RoomNumber = " & RoomNumber & " AND RateTypeID = " & dcRateType.BoundText, CN, adOpenStatic, adLockOptimistic
    
        If .RecordCount > 0 Then
            txtRate.Text = toMoney(!RoomRate)
            txtAdults.Tag = !ExtraAdultRate
            txtChildrens.Tag = !ExtraChildRate
        End If
    End With
    
    rsRoomRates.Close
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblRatePerPeriod.FontUnderline = False
    lblAmountPaid.FontUnderline = False
    lblOtherCharges.FontUnderline = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmRoomsWindow.RefreshRecords
    
    Unload frmPayment
    
    Set frmPayment = Nothing
    Set frmCheckOut = Nothing
End Sub

Private Sub lblAmountPaid_Click()
    With frmPayment
        .FolioNumber = txtGuestName.Tag
        .GuestName = txtGuestName.Text
        .Balance = txtBalance.Text
        .RefreshBalance
        Set .RefForm = Me
        
        .Show vbModal
        
        txtAmountPaid.Text = toMoney(AmountPaid)
    End With
End Sub

Private Sub lblAmountPaid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetHandCur True
    lblAmountPaid.FontUnderline = True
End Sub

Private Function ComputeRatePerPeriod() As Currency
    Dim rsRoomRates As New ADODB.Recordset
    
    With rsRoomRates
        .Open "SELECT * FROM [Rate Per Period] WHERE FolioNumber = '" & txtGuestName.Tag & "'", CN, adOpenStatic, adLockOptimistic
    
        Do Until .EOF
            ComputeRatePerPeriod = ComputeRatePerPeriod + toMoney(!Rate) + toMoney(!Adults) + toMoney(!Childrens)
            
            .MoveNext
        Loop
    End With
    
    rsRoomRates.Close
End Function

Private Sub lblOtherCharges_Click()
    With frmOtherCharges
        .FolioNumber = txtGuestName.Tag
        .GuestName = txtGuestName.Text
        
        Set .RefForm = Me
        
        .Show vbModal
        
        txtOtherCharges.Text = toMoney(OtherCharges)
    End With
End Sub

Private Sub lblOtherCharges_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetHandCur True
    lblOtherCharges.FontUnderline = True
End Sub

Private Sub lblRatePerPeriod_Click()
    With frmRatePerPeriod
        .FolioNumber = txtGuestName.Tag
        
        .Show vbModal
        
        Call ComputeRate
    End With
End Sub

Private Sub lblRatePerPeriod_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetHandCur True
    lblRatePerPeriod.FontUnderline = True
End Sub

Private Sub txtAmountPaid_Change()
    txtBalance.Text = toMoney(toNumber(txtTotal.Text) - toNumber(txtAmountPaid.Text))
End Sub

Private Sub txtDays_Change()
On Error GoTo err
    
    Dim rsRatePerPeriod As New ADODB.Recordset
    Dim tmpDate As Date
    Dim minNoofPerson As Integer
    
    tmpDate = txtDateIn.Text
    
    If txtAdults.Tag = "" Then Exit Sub
    
    CN.BeginTrans
    
    CN.Execute "DELETE [Date] " & _
                "FROM [Rate Per Period] " & _
                "WHERE [Date]>#" & dtpDateOut - 1 & "#"

    Dim intAdults As Integer
    
    minNoofPerson = getValueAt("SELECT * FROM [Room Rates] WHERE RoomNumber = " & RoomNumber & " AND RateTypeID = " & dcRateType.BoundText, "NoofPerson")
    
    If txtAdults.Text = minNoofPerson Then
        intAdults = 0
    Else
        intAdults = CInt(txtAdults.Text) - minNoofPerson
    End If

    With rsRatePerPeriod
        .Open "SELECT * FROM [Rate Per Period] WHERE FolioNumber = '" & txtGuestName.Tag & "' ORDER BY [Date]", CN, adOpenStatic, adLockOptimistic

        Do Until tmpDate > dtpDateOut.Value - 1
            .Filter = "[Date] = #" & tmpDate & "#"
            
            If .RecordCount = 0 Then
                .AddNew
                
                .Fields("FolioNumber") = txtGuestName.Tag
                .Fields("Date") = tmpDate
                .Fields("RoomNumber") = txtRoomNumber.Text
                .Fields("RateTypeID") = dcRateType.BoundText
                .Fields("Rate") = txtRate.Text
                .Fields("Adults") = txtAdults.Tag * intAdults
                .Fields("Childrens") = toMoney(txtChildrens.Tag) * toNumber(txtChildrens.Text)
                
                .Update
            End If
            tmpDate = tmpDate + 1
        Loop
    End With
    
    CN.CommitTrans
    
    rsRatePerPeriod.Close

    Exit Sub

err:
    CN.RollbackTrans
    prompt_err err, Name, "txtDays_Change"
    Screen.MousePointer = vbDefault
End Sub
