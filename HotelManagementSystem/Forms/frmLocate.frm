VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmLocate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Locate Database"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4980
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPath 
      Height          =   315
      Left            =   600
      TabIndex        =   3
      Top             =   1320
      Width           =   3615
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   345
      Left            =   4260
      TabIndex        =   2
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   315
      Left            =   2490
      TabIndex        =   1
      Top             =   1770
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   3600
      TabIndex        =   0
      Top             =   1770
      Width           =   1035
   End
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   3000
      Top             =   270
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Provider"
      Height          =   1005
      Left            =   90
      TabIndex        =   4
      Top             =   150
      Width           =   4605
      Begin VB.OptionButton optJet4 
         Caption         =   "Microsoft Jet OLEDB 4.0"
         Height          =   285
         Left            =   150
         TabIndex        =   6
         Top             =   300
         Value           =   -1  'True
         Width           =   2655
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ODBC"
         Enabled         =   0   'False
         Height          =   285
         Left            =   150
         TabIndex        =   5
         Top             =   630
         Width           =   2655
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Path"
      Height          =   285
      Left            =   90
      TabIndex        =   7
      Top             =   1350
      Width           =   765
   End
End
Attribute VB_Name = "frmLocate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowse_Click()
  dlgOpen.ShowOpen
  txtPath.Text = dlgOpen.FileName
End Sub

Private Sub cmdCancel_Click()
  If CN.State = 1 Then CN.Close
  End
End Sub

Private Sub cmdOK_Click()
  If Trim(txtPath.Text) = "" Then
    'do nothing
  Else
    DBPath = txtPath.Text
    SetINI "Configuration", "Path", DBPath
    
    
    Unload Me
    'If InvalidDB Then frmLogin.show 1
  End If
End Sub

Private Sub Form_Load()
  With dlgOpen
    .Filter = "All Files (*.*)|*.*| MS Access Files (*.mdb)|*.mdb|"
  End With
End Sub

