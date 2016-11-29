VERSION 5.00
Object = "{9EDDC69F-10E8-4DE7-9648-EC8A45F005C0}#1.0#0"; "b8Controls4.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRPTRoomHistory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Room History"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboRoomNumber 
      Height          =   315
      Left            =   1770
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   270
      Width           =   1815
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   405
      Left            =   3240
      TabIndex        =   1
      Top             =   2160
      Width           =   945
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   405
      Left            =   2130
      TabIndex        =   0
      Top             =   2160
      Width           =   945
   End
   Begin b8Controls4.b8Line b8Line1 
      Height          =   30
      Left            =   30
      TabIndex        =   2
      Top             =   1800
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   53
      BorderColor1    =   11325655
      BorderColor2    =   16185592
   End
   Begin MSComCtl2.DTPicker dtpBegDate 
      Height          =   375
      Left            =   900
      TabIndex        =   3
      Top             =   1080
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "MMM-dd-yyyy"
      Format          =   77266947
      CurrentDate     =   39156
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   375
      Left            =   2580
      TabIndex        =   4
      Top             =   1080
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "MMM-dd-yyyy"
      Format          =   77266947
      CurrentDate     =   39156
   End
   Begin VB.Label Label2 
      Caption         =   "From"
      Height          =   255
      Left            =   900
      TabIndex        =   8
      Top             =   840
      Width           =   825
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Date"
      Height          =   315
      Left            =   300
      TabIndex        =   7
      Top             =   1110
      Width           =   525
   End
   Begin VB.Label Label4 
      Caption         =   "To"
      Height          =   255
      Left            =   2610
      TabIndex        =   6
      Top             =   810
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Room"
      Height          =   195
      Left            =   1200
      TabIndex        =   5
      Top             =   270
      Width           =   420
   End
End
Attribute VB_Name = "frmRPTRoomHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    With frmReports
        .strReport = "Room History"
        
        If cboRoomNumber.Text = "All Rooms" Then
            .strWhere = "{qry_rpt_Room_History.DateIn} IN #" & dtpBegDate.Value & "# TO #" & dtpEndDate.Value & "# OR {qry_rpt_Room_History.DateOut} IN #" & dtpBegDate.Value & "# TO #" & dtpEndDate.Value & "#"
        Else
            .strWhere = "{qry_rpt_Room_History.RoomNumber} = " & cboRoomNumber.Text & " AND ({qry_rpt_Room_History.DateIn} IN #" & dtpBegDate.Value & "# TO #" & dtpEndDate.Value & "# OR {qry_rpt_Room_History.DateOut} IN #" & dtpBegDate.Value & "# TO #" & dtpEndDate.Value & "#)"
        End If

        frmReports.Show vbModal
    End With

End Sub

Private Sub Form_Load()
On Error GoTo err_Form_Load

    Dim rsRooms As New Recordset
    
    cboRoomNumber.AddItem "All Rooms"
    
    With rsRooms
        .Open "SELECT RoomNumber FROM Rooms ORDER BY RoomNumber ASC", CN, adOpenStatic, adLockOptimistic
    
        Do Until .EOF
            cboRoomNumber.AddItem !RoomNumber
            
            .MoveNext
        Loop
    End With
    
    cboRoomNumber.ListIndex = 0
    
    dtpBegDate.Value = Date
    dtpEndDate.Value = Date
    
    Exit Sub
    
err_Form_Load:
    MsgBox err.Description, vbInformation
End Sub

