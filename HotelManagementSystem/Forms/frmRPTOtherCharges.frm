VERSION 5.00
Object = "{9EDDC69F-10E8-4DE7-9648-EC8A45F005C0}#1.0#0"; "b8Controls4.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRPTOtherCharges 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Other Charges"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   5220
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   405
      Left            =   2460
      TabIndex        =   1
      Top             =   2250
      Width           =   945
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   405
      Left            =   3570
      TabIndex        =   0
      Top             =   2250
      Width           =   945
   End
   Begin b8Controls4.b8Line b8Line1 
      Height          =   30
      Left            =   360
      TabIndex        =   2
      Top             =   1890
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   53
      BorderColor1    =   11325655
      BorderColor2    =   16185592
   End
   Begin MSComCtl2.DTPicker dtpBegDate 
      Height          =   375
      Left            =   1230
      TabIndex        =   3
      Top             =   1170
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
      Left            =   2910
      TabIndex        =   4
      Top             =   1170
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "MMM-dd-yyyy"
      Format          =   77266947
      CurrentDate     =   39156
   End
   Begin MSDataListLib.DataCombo dcChargeType 
      Height          =   315
      Left            =   2010
      TabIndex        =   8
      Top             =   330
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Other Charges"
      Height          =   195
      Left            =   960
      TabIndex        =   9
      Top             =   360
      Width           =   1020
   End
   Begin VB.Label Label4 
      Caption         =   "To"
      Height          =   255
      Left            =   2940
      TabIndex        =   7
      Top             =   900
      Width           =   825
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Date"
      Height          =   315
      Left            =   630
      TabIndex        =   6
      Top             =   1200
      Width           =   525
   End
   Begin VB.Label Label2 
      Caption         =   "From"
      Height          =   255
      Left            =   1230
      TabIndex        =   5
      Top             =   930
      Width           =   825
   End
End
Attribute VB_Name = "frmRPTOtherCharges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    With frmReports
        .strReport = "Account Receivable"
        
        If dcChargeType.Text <> "" Then
            .strWhere = "{qry_rpt_Other_Charges.ChargeType} = '" & dcChargeType.Text & "' AND {qry_rpt_Other_Charges.Date} IN #" & dtpBegDate.Value & "# TO #" & dtpEndDate.Value & "#"
        Else
            .strWhere = "{qry_rpt_Other_Charges.Date} IN #" & dtpBegDate.Value & "# TO #" & dtpEndDate.Value & "#"
        End If

        frmReports.Show vbModal
    End With
End Sub

Private Sub Form_Load()
    bind_dc "SELECT * FROM [Charge Type]", "ChargeType", dcChargeType, "ChargeTypeID"
    
    dtpBegDate.Value = Date
    dtpEndDate.Value = Date
End Sub
