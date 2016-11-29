VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDailyReport 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2385
   ClientLeft      =   3765
   ClientTop       =   1890
   ClientWidth     =   4515
   ControlBox      =   0   'False
   Icon            =   "frmDailyReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4515
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker Cal 
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   600
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      _Version        =   393216
      Format          =   97517569
      CurrentDate     =   40552
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "X"
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export"
      Height          =   615
      Left            =   1080
      Picture         =   "frmDailyReport.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   840
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   ".xls"
      DialogTitle     =   "Export"
      Filter          =   "Microsoft Excel 2K (*.xls)|*.xls|"
      Flags           =   4
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   615
      Left            =   2520
      Picture         =   "frmDailyReport.frx":0396
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00000000&
      Caption         =   " DAILY REPORT"
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
      TabIndex        =   1
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmDailyReport"
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
  cmdExport.Enabled = False
  cmdPrint.Enabled = False
  Dialog.FileName = App.Path & "\" & Format(Cal.Value, "mmmm") & "_" & Trim(Str(Cal.Day)) & "_" & Trim(Str(Cal.Year)) & " - " & Format(Cal.Value, "dddd")
  Export = True
  Dialog.ShowSave
  If Dialog.FileName = "" Then GoTo Monger
  SavePath = Dialog.FileName
    
  Screen.MousePointer = vbHourglass
  If Excel_Daily = True Then
    oWB.Close 0
    Set oXL = Nothing
    Set oSheet = Nothing
    Unload Me
    MsgBox "The report was successfully exported in excel format!"
  End If
Monger:
  cmdExport.Enabled = True
  cmdPrint.Enabled = True
  Screen.MousePointer = vbNormal
End Sub

Private Sub cmdPrint_Click()
On Error GoTo Monger
  
  cmdExport.Enabled = False
  cmdPrint.Enabled = False
  Export = False
  Screen.MousePointer = vbHourglass
  If Excel_Daily = True Then
    oWB.Close 0
    oXL.Quit
    Unload Me
  End If

Monger:
  cmdExport.Enabled = True
  cmdPrint.Enabled = True
  Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Load()
  Set oXL = CreateObject("Excel.Application")
  oXL.Visible = False
  Cal.Value = Format(Now, "ddddd")
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set oSheet = Nothing
  Set oWB = Nothing
  Set oXL = Nothing
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  MoveForm Me
End Sub
