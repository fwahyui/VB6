VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmChangeRoom 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5805
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6405
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   6405
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
      ScaleWidth      =   417
      TabIndex        =   10
      Top             =   0
      Width           =   6255
      Begin VB.Image Image1 
         Height          =   480
         Left            =   30
         Picture         =   "frmChangeRoom.frx":0000
         Top             =   30
         Width           =   480
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Change Room"
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
         Width           =   1995
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
         TabIndex        =   11
         Top             =   360
         Width           =   3900
      End
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4950
      TabIndex        =   9
      Top             =   5190
      Width           =   1215
   End
   Begin VB.CommandButton CmdChangeRoom 
      Caption         =   "Change Room"
      Height          =   375
      Left            =   3690
      TabIndex        =   8
      Top             =   5190
      Width           =   1215
   End
   Begin VB.TextBox txtTo 
      Height          =   375
      Left            =   5190
      TabIndex        =   7
      Top             =   930
      Width           =   645
   End
   Begin VB.TextBox txtFrom 
      Height          =   345
      Left            =   3780
      TabIndex        =   5
      Top             =   930
      Width           =   765
   End
   Begin MSComCtl2.DTPicker dtpStartDate 
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   900
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   661
      _Version        =   393216
      Format          =   79495169
      CurrentDate     =   39564
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   3435
      Left            =   60
      TabIndex        =   0
      Top             =   1650
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   6059
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Room Number"
         Object.Width           =   2937
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Room Type"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Room Status"
         Object.Width           =   3334
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "To"
      Height          =   285
      Left            =   4830
      TabIndex        =   6
      Top             =   960
      Width           =   315
   End
   Begin VB.Label Label2 
      Caption         =   "From"
      Height          =   345
      Left            =   3270
      TabIndex        =   4
      Top             =   930
      Width           =   465
   End
   Begin VB.Label Label1 
      Caption         =   "Starting Date"
      Height          =   345
      Left            =   60
      TabIndex        =   3
      Top             =   900
      Width           =   1065
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Available Rooms"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   210
      Left            =   135
      TabIndex        =   1
      Top             =   1350
      Width           =   4815
   End
   Begin VB.Shape shpBar 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   60
      Top             =   1350
      Width           =   6195
   End
End
Attribute VB_Name = "frmChangeRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    frmCheckIn.blnChangeRoom = False
    
    Unload Me
End Sub

Private Sub CmdChangeRoom_Click()
    With frmCheckIn
        If dtpStartDate.Value < .dtpDateIn.Value Then
            MsgBox "Start date must be higher than the 'Date In' from Check in form", vbInformation
            
            Exit Sub
        End If
        
        If dtpStartDate.Value > .dtpDateOut.Value Then
            .dtpDateOut.Value = dtpStartDate.Value
        ElseIf dtpStartDate.Value = .dtpDateOut.Value Then
            .dtpDateOut.Value = dtpStartDate.Value + 1
        End If
        
        .txtRoomNumber.Text = txtTo.Text
        .txtRate = toMoney(getValueAt("SELECT RoomRate FROM [Room Rates] WHERE RoomNumber = " & .txtRoomNumber.Text & " AND RateTypeID = " & .dcRateType.BoundText, "RoomRate"))
        
        Call .dtpDateOut_Change
        
        .blnChangeRoom = True
        Unload Me
    End With
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    Dim RS As New Recordset
    
    dtpStartDate.Value = Date
    
    RS.Open "SELECT * FROM qry_Rooms_Window WHERE Status = 'Vacant'", CN, adOpenStatic, adLockOptimistic
    
    FillListView lvList, RS, 3, 2, False, True
    
    Call lvList_Click
    
    RS.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmChangeRoom = Nothing
End Sub

Private Sub lvList_Click()
    txtTo.Text = lvList.SelectedItem
End Sub
