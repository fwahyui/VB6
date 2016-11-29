VERSION 5.00
Object = "{9EDDC69F-10E8-4DE7-9648-EC8A45F005C0}#1.0#0"; "b8Controls4.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmRPTAccRec 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Account Receivable"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   405
      Left            =   3450
      TabIndex        =   2
      Top             =   1500
      Width           =   945
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   405
      Left            =   2340
      TabIndex        =   1
      Top             =   1500
      Width           =   945
   End
   Begin b8Controls4.b8Line b8Line1 
      Height          =   30
      Left            =   240
      TabIndex        =   4
      Top             =   1260
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   53
      BorderColor1    =   11325655
      BorderColor2    =   16185592
   End
   Begin MSDataListLib.DataCombo dcCompany 
      Height          =   315
      Left            =   1590
      TabIndex        =   0
      Top             =   540
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Company"
      Height          =   195
      Left            =   840
      TabIndex        =   3
      Top             =   600
      Width           =   660
   End
End
Attribute VB_Name = "frmRPTAccRec"
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
                
        frmReports.Show vbModal
    End With

End Sub

Private Sub Form_Load()
    bind_dc "SELECT * FROM [Company] WHERE Company Is Not Null ORDER BY Company", "Company", dcCompany, "CompanyID", True
End Sub
