VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmjawab 
   Caption         =   "Identifikasi Potensi Siswa"
   ClientHeight    =   9045
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9045
   ScaleWidth      =   11805
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtnis 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1320
      Width           =   2295
   End
   Begin VB.TextBox txtnamasiswa 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1800
      Width           =   5415
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1065
      ScaleWidth      =   11775
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   11805
      Begin VB.Image Image1 
         Height          =   855
         Left            =   75
         Picture         =   "frmjawab.frx":0000
         Top             =   75
         Width           =   810
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Identifikasi Potensi Siswa"
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
         TabIndex        =   4
         Top             =   120
         Width           =   7035
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Beri Tanda Centang Pada Chekbox Apabila Ya"
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   3
         Top             =   600
         Width           =   4335
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Selesai"
      Height          =   495
      Left            =   7200
      TabIndex        =   1
      Top             =   9000
      Width           =   1335
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   5460
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   14940
      _ExtentX        =   26353
      _ExtentY        =   9631
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      RecordNavigatorString=   "Baris ke:|dari"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      TabKeyBehavior  =   1
      UseEvenOddColor =   -1  'True
      MethodHoldFields=   -1  'True
      GroupByBoxVisible=   0   'False
      RowHeaders      =   -1  'True
      HeaderFontName  =   "Arial"
      HeaderFontBold  =   -1  'True
      HeaderFontSize  =   9.75
      HeaderFontWeight=   700
      ColumnHeaderHeight=   330
      IntProp1        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmjawab.frx":24C8
      Column(2)       =   "frmjawab.frx":2590
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmjawab.frx":2634
      FormatStyle(2)  =   "frmjawab.frx":278C
      FormatStyle(3)  =   "frmjawab.frx":283C
      FormatStyle(4)  =   "frmjawab.frx":28F0
      FormatStyle(5)  =   "frmjawab.frx":29C8
      FormatStyle(6)  =   "frmjawab.frx":2A80
      ImageCount      =   0
      PrinterProperties=   "frmjawab.frx":2B60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NIS"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   585
   End
End
Attribute VB_Name = "frmjawab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private oSiswa As DLLBK.Csiswa
Private oSoal As DLLBK.cSoal
Dim rsa As New ADODB.Recordset
Private Sub Form_Load()
    Set oSiswa = New DLLBK.Csiswa
    Set oSoal = New DLLBK.cSoal
    txtnis.text = ""
    txtnamasiswa.text = ""
    GridEX1.Left = 120
End Sub

Private Sub txtnis_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If oSiswa.Load(txtnis.text) > 0 Then
            txtnamasiswa.text = oSiswa.Nama
            Call isi
            Dim rs As New ADODB.Recordset
            Set rs = Nothing
            rs.CursorLocation = adUseClient
            rs.Open "select a.NIS,a.NoSoal,b.Soal,a.Jawaban from TbJawaban a, tbsoal b where a.NIS=" & txtnis.text & " and a.nosoal=b.nosoal", koneksi, adOpenDynamic, adLockOptimistic
            Set GridEX1.ADORecordset = rs
            Call formatGrid
        End If
    End If
End Sub
Private Sub formatGrid()
On Error Resume Next
GridEX1.Columns(1).Width = 100
GridEX1.Columns(1).Visible = False
GridEX1.Columns(2).Width = 1000
GridEX1.Columns(3).Width = GridEX1.Width - 3000
GridEX1.Columns(4).Width = 1300
End Sub

Private Sub isi()
'On Error GoTo Hell
On Error Resume Next
    Dim isi As New ADODB.Recordset
    sql = "select * from tbsoal"
    Set rsa = Nothing
    rsa.CursorLocation = adUseClient
    rsa.Open sql, koneksi
    
    For a = 0 To rsa.RecordCount - 1
        sql = "insert into tbjawaban (NIS,NoSoal,Jawaban) values (" & txtnis.text & "," & rsa!Nosoal & ",0)"
        Set isi = Nothing
        isi.Open sql, koneksi
        rsa.MoveNext
    Next
'Hell:
'    If Err.Number = -2147467259 Then
'        MsgBox "Siswa Yang Bersangkutan Sudah Melakukan Tes"
'    End If
End Sub
Private Sub Form_Resize()
On Error Resume Next
Dim ScaleFactorX As Single
Dim ScaleFactorY As Single

If Not DoResize Then
    DoResize = True
    Exit Sub
End If
RePosForm = False
If (WindowState = vbMaximized) Then
    ScaleFactorX = Me.Width - CurrentWidth
    CurrentWidth = Me.Width
    ScaleFactorY = Me.Height - CurrentHeight
    CurrentHeight = Me.Height
    ResizeControl ScaleFactorX, ScaleFactorY
Else
    ScaleFactorX = Me.Width - CurrentWidth
    ScaleFactorY = Me.Height - CurrentHeight
    CurrentHeight = Me.Height
    CurrentWidth = Me.Width
    ResizeControl ScaleFactorX, ScaleFactorY
End If
End Sub
Private Sub ResizeControl(ByVal SFX As Single, ByVal SFY As Single)
On Error GoTo Hell
''   === Place controls resize here ==============
With GridEX1
    .Width = .Width + SFX
    .Height = .Height + SFY
End With
'   === End Controls ============================
If RePosForm Then
    Move Left + SFX, Top + SFY, Width + SFX, Height + SFY
End If

Exit Sub
Hell:
    MsgBox Err.Description, vbInformation
End Sub

