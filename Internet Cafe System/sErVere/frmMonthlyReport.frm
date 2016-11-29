VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMonthlyReport 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2265
   ClientLeft      =   3015
   ClientTop       =   3390
   ClientWidth     =   5055
   ControlBox      =   0   'False
   Icon            =   "frmMonthlyReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   615
      Left            =   3600
      Picture         =   "frmMonthlyReport.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export"
      Height          =   615
      Left            =   2160
      Picture         =   "frmMonthlyReport.frx":0156
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "X"
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   375
   End
   Begin VB.ComboBox cmbYear 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmMonthlyReport.frx":04E0
      Left            =   2880
      List            =   "frmMonthlyReport.frx":04F3
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.ComboBox cmbMonth 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmMonthlyReport.frx":0515
      Left            =   960
      List            =   "frmMonthlyReport.frx":053D
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   840
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   0
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   ".xls"
      DialogTitle     =   "Export"
      Filter          =   "Microsoft Excel 2K (*.xls)|*.xls|"
      Flags           =   4
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00000000&
      Caption         =   " MONTHLY REPORT"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Year:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   3
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Month:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "frmMonthlyReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdExport_Click()
On Error GoTo Monger
  cmdPrint.Enabled = False
  cmdExport.Enabled = False
  
  Select Case cmbMonth.Text
  Case "January": cmbMonth.Tag = "1"
  Case "February": cmbMonth.Tag = "2"
  Case "March": cmbMonth.Tag = "3"
  Case "April": cmbMonth.Tag = "4"
  Case "May": cmbMonth.Tag = "5"
  Case "June": cmbMonth.Tag = "6"
  Case "July": cmbMonth.Tag = "7"
  Case "August": cmbMonth.Tag = "8"
  Case "September": cmbMonth.Tag = "9"
  Case "October": cmbMonth.Tag = "10"
  Case "November": cmbMonth.Tag = "11"
  Case "December": cmbMonth.Tag = "12"
  End Select
  
  Dialog.FileName = App.Path & "\" & cmbMonth.Text & "_" & cmbYear.Text
  Export = True
  Dialog.ShowSave
  If Dialog.FileName = "" Then GoTo Monger
  SavePath = Dialog.FileName
    
  Screen.MousePointer = vbHourglass
  If Excel_Monthly(Val(cmbMonth.Tag), Val(cmbYear.Text)) = True Then
    oWB.Close 0
    oXL.Quit
    Unload Me
    MsgBox "The report was successfully exported in excel format!"
  End If

Monger:
  cmdPrint.Enabled = True
  cmdExport.Enabled = True
  Screen.MousePointer = vbNormal
End Sub

Private Sub cmdPrint_Click()
On Error GoTo Monger
  cmdPrint.Enabled = False
  cmdExport.Enabled = False
  
  Select Case cmbMonth.Text
  Case "January": cmbMonth.Tag = "1"
  Case "February": cmbMonth.Tag = "2"
  Case "March": cmbMonth.Tag = "3"
  Case "April": cmbMonth.Tag = "4"
  Case "May": cmbMonth.Tag = "5"
  Case "June": cmbMonth.Tag = "6"
  Case "July": cmbMonth.Tag = "7"
  Case "August": cmbMonth.Tag = "8"
  Case "September": cmbMonth.Tag = "9"
  Case "October": cmbMonth.Tag = "10"
  Case "November": cmbMonth.Tag = "11"
  Case "December": cmbMonth.Tag = "12"
  End Select

  Export = False
  Screen.MousePointer = vbHourglass
  If Excel_Monthly(Val(cmbMonth.Tag), Val(cmbYear.Text)) = True Then
    oWB.Close 0
    oXL.Quit
    Unload Me
  End If

Monger:
  cmdPrint.Enabled = True
  cmdExport.Enabled = True
  Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Load()
  Set oXL = CreateObject("Excel.Application")
  oXL.Visible = False

  Select Case Format(Now, "m")
  Case 1: cmbMonth.Text = "January"
  Case 2: cmbMonth.Text = "February"
  Case 3: cmbMonth.Text = "March"
  Case 4: cmbMonth.Text = "April"
  Case 5: cmbMonth.Text = "May"
  Case 6: cmbMonth.Text = "June"
  Case 7: cmbMonth.Text = "July"
  Case 8: cmbMonth.Text = "August"
  Case 9: cmbMonth.Text = "September"
  Case 10: cmbMonth.Text = "October"
  Case 11: cmbMonth.Text = "November"
  Case 12: cmbMonth.Text = "December"
  End Select
  cmbYear.AddItem (Format(Now, "yyyy"))
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set oSheet = Nothing
  Set oWB = Nothing
  Set oXL = Nothing
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  MoveForm Me
End Sub
