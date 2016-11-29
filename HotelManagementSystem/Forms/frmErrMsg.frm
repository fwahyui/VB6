VERSION 5.00
Begin VB.Form frmErrMsg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "3JNet - Unhandled error"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   5730
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   345
      Left            =   4470
      TabIndex        =   0
      Top             =   3780
      Width           =   1155
   End
   Begin VB.Image Image1 
      Height          =   4260
      Left            =   0
      Picture         =   "frmErrMsg.frx":0000
      Top             =   0
      Width           =   5745
   End
   Begin VB.Label lblMsg 
      BackStyle       =   0  'Transparent
      Caption         =   "Unknown Error!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   3345
      Left            =   2130
      TabIndex        =   1
      Top             =   660
      Width           =   3315
   End
End
Attribute VB_Name = "frmErrMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mModuleName As String
Dim mRoutineName As String
Dim mDetail As String


'Public Sub ShowForm(ByVal sModuleName As String, ByVal sRoutineName As String, ByVal sDetail As String)
'
'
'    'set param
'    mModuleName = sModuleName
'    mRoutineName = sRoutineName
'    mDetail = sDetail
'
'    lblMsg.Caption = "Module: '" & sModuleName & "'" & vbNewLine & _
'                    "Procedure: '" & sRoutineName & "'" & vbNewLine & _
'                        "Detail: '" & sDetail & "'"
'
'
'    On Error Resume Next
'    Me.Show vbModal
'    err.Clear
'    Unload Me
'End Sub

'Private Sub cmdClose_Click()
'    Form_SaveError
'End Sub

'Private Sub Form_SaveError()
'
'
'    Dim vRS As New ADODB.Recordset
'    Dim sSQL As String
'
'
'    sSQL = "SELECT * FROM tblErrLog"
'
'    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
'        GoTo RAE
'    End If
'
'    'add new record
'    With vRS
'        .AddNew
'        .Fields("DateOccured") = Now
'        .Fields("Module") = mModuleName
'        .Fields("Routine") = mRoutineName
'        .Fields("Detail") = mDetail
'        .Update
'    End With
'
'RAE:
'    Set vRS = Nothing
'    'close this form
'    Unload Me
'End Sub



