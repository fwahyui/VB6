VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmcaribim 
   Caption         =   "Daftar Masalah Dan Penanganan"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8085
   ScaleWidth      =   9000
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1065
      ScaleWidth      =   8970
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   9000
      Begin VB.Image Image1 
         Height          =   855
         Left            =   75
         Picture         =   "frmcaribim.frx":0000
         Top             =   75
         Width           =   810
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Daftar Masalah Dan Penanganan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   495
         Index           =   1
         Left            =   1080
         TabIndex        =   9
         Top             =   240
         Width           =   7035
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   8775
      Begin VB.CommandButton Command3 
         Caption         =   "&Keluar"
         Height          =   375
         Left            =   7200
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "L&ihat"
         Height          =   375
         Left            =   6360
         TabIndex        =   7
         Top             =   360
         Width           =   735
      End
      Begin MSComCtl2.DTPicker tg1 
         Height          =   375
         Left            =   3120
         TabIndex        =   4
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16449537
         CurrentDate     =   39847
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Lihat"
         Height          =   375
         Left            =   2280
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   360
         Width           =   2055
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   5175
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   9128
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.DTPicker tg2 
         Height          =   375
         Left            =   4920
         TabIndex        =   5
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16449537
         CurrentDate     =   39847
      End
      Begin VB.Label Label1 
         Caption         =   "S/D"
         Height          =   255
         Index           =   0
         Left            =   4560
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmcaribim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private oBim As New DLLBK.cBK
Dim rs As New ADODB.Recordset

Private Sub Command1_Click()
On Error GoTo salah
Dim pnis As Long
pnis = Text1.text
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
    Set rs = oBim.List(False, "where NIs = " & pnis & "")
    Set DataGrid1.DataSource = rs
salah:
    If Err.Number <> 0 Then
        MsgBox "Parameter Yang Anda Masukkan Tidak Sesuai Dengan Database Kami", vbCritical
    End If
End Sub

Private Sub Command2_Click()
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
    Set rs = oBim.List(False, "where tgl between '" & tg1.Value & "' and '" & tg2.Value & "'")
    Set DataGrid1.DataSource = rs
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set oBim = New DLLBK.cBK
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Frame1.Top = 1200
    Frame1.Height = Me.Height - Command1.Height - 1200 - 750
    Frame1.Width = Me.Width - 360
    DataGrid1.Top = 840
    DataGrid1.Height = Frame1.Height - 1000
    DataGrid1.Width = Frame1.Width - 360
    Command3.Left = Frame1.Width - 240 - Command3.Width
End Sub
