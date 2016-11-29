VERSION 5.00
Object = "{9EDDC69F-10E8-4DE7-9648-EC8A45F005C0}#1.0#0"; "b8Controls4.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRPTDueReservation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Due Reservation"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   405
      Left            =   3360
      TabIndex        =   1
      Top             =   1560
      Width           =   945
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   405
      Left            =   2250
      TabIndex        =   0
      Top             =   1560
      Width           =   945
   End
   Begin MSComCtl2.DTPicker dtpCheckIn 
      Height          =   375
      Left            =   1890
      TabIndex        =   2
      Top             =   360
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   661
      _Version        =   393216
      Format          =   77266945
      CurrentDate     =   39592
   End
   Begin b8Controls4.b8Line b8Line1 
      Height          =   30
      Left            =   150
      TabIndex        =   3
      Top             =   1320
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   53
      BorderColor1    =   11325655
      BorderColor2    =   16185592
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Check Out Date"
      Height          =   195
      Left            =   600
      TabIndex        =   4
      Top             =   390
      Width           =   1155
   End
End
Attribute VB_Name = "frmRPTDueReservation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    With frmReports
        .strReport = "Due Reservation"
        
        .strWhere = "{qry_rpt_Due_Reservation.DateIn} = #" & dtpCheckIn.Value & "#"

        frmReports.Show vbModal
    End With
End Sub

Private Sub Form_Load()
    dtpCheckIn.Value = Date
End Sub
