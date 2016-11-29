VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmentryKelas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Entry Kelas"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1095
      Index           =   1
      Left            =   0
      ScaleHeight     =   1065
      ScaleWidth      =   6645
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   6675
      Begin VB.Image Image1 
         Height          =   855
         Left            =   75
         Picture         =   "frmentryKelas.frx":0000
         Top             =   75
         Width           =   810
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Entry Kelas"
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
         Index           =   23
         Left            =   1080
         TabIndex        =   13
         Top             =   120
         Width           =   7035
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Masukkan Semua Data Kelas"
         Height          =   375
         Index           =   2
         Left            =   1080
         TabIndex        =   12
         Top             =   600
         Width           =   4335
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFC0C0&
      Height          =   3615
      Index           =   0
      Left            =   120
      ScaleHeight     =   3555
      ScaleWidth      =   6315
      TabIndex        =   3
      Top             =   1560
      Width           =   6375
      Begin VB.ComboBox txtkelas 
         BackColor       =   &H00C0E0FF&
         Height          =   315
         ItemData        =   "frmentryKelas.frx":24C8
         Left            =   2280
         List            =   "frmentryKelas.frx":24D5
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox txtruangan 
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
         Left            =   2280
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   2040
         Width           =   3615
      End
      Begin VB.TextBox txtjumlahsiswa 
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
         Left            =   2280
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   2520
         Width           =   3615
      End
      Begin VB.TextBox txtidkel 
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
         Left            =   2280
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1080
         Width           =   3615
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   10
         Top             =   600
         Width           =   3615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID Kelas"
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
         Left            =   600
         TabIndex        =   9
         Top             =   1080
         Width           =   840
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kelas"
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
         Left            =   600
         TabIndex        =   8
         Top             =   1560
         Width           =   510
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ruangan"
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
         Left            =   600
         TabIndex        =   7
         Top             =   2040
         Width           =   840
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Siswa"
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
         Left            =   600
         TabIndex        =   6
         Top             =   2520
         Width           =   1275
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tahun Ajaran"
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
         Left            =   600
         TabIndex        =   5
         Top             =   600
         Width           =   1290
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   14
      Top             =   1095
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   635
      ButtonWidth     =   1746
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Simpan"
            Key             =   "Simpan"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Baru     "
            Key             =   "New"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Tutup"
            Key             =   "Exit"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   8730
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmentryKelas.frx":24E6
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmentryKelas.frx":3578
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmentryKelas.frx":460A
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmentryKelas.frx":569C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmentryKelas.frx":672E
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmentryKelas.frx":77C0
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmentryKelas.frx":8852
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmentryKelas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private oKelas As New DLLBK.Ckelas
Private oTahun As New DLLBK.cTahun
Dim DataMode As ENUM_DATA_MODE

Private Sub Form_Load()
    Set oKelas = New DLLBK.Ckelas
    Set oTahun = New DLLBK.cTahun
    lbl.Caption = strTahun
    New_data
End Sub

Private Sub cariIDKelas()
    Dim rs As New ADODB.Recordset
    Set rs = Nothing
    rs.CursorLocation = adUseClient
    Set rs = oKelas.List
    Dim Jumlah As Integer
    Jumlah = rs.RecordCount + 1
    txtidkel = Jumlah
End Sub
Private Sub mnuSave()
On Error GoTo Hell
If DataMode = EN_NEW Then
    MsgBox "Data harus diisi dulu" & vbCrLf & "Simpan data dibatalkan", vbExclamation, "Simpan Data"
    Exit Sub
ElseIf DataMode = EN_SAVED Then
    MsgBox "Tidak ada data yang berubah" & vbCrLf & "Simpan data dibatalkan", vbExclamation, "Simpan Data"
    Exit Sub
End If
If SaveData(PidTahun, txtidkel.text, txtkelas.text, txtruangan.text, txtjumlahsiswa.text) > 0 Then
    MsgBox "Data BERHASIL disimpan", vbInformation, "Simpan Data"
    Form_Load
Else
    MsgBox "Data GAGAL disimpan", vbCritical, "Simpan Data"
End If
Exit Sub
Hell:
    MsgBox Err.Description, vbCritical, "Internal"
End Sub
Private Function SaveData(IDTahunAjaran As Integer, IDKelas As Integer, Kelas As String, _
                Ruangan As String, JumlahSiswa As Integer) As Integer
                Dim rs As New ADODB.Recordset
If DataMode = EN_NEW_CHANGED Then
    Set rs = Nothing
    rs.CursorLocation = adUseClient
    Set rs = oKelas.List(True, "Where kelas='" & Kelas & "' and Ruangan = '" & Ruangan & "'")
    If rs.RecordCount > 0 Then
        If MsgBox("Data Sudah Ada..." & vbCrLf & "Apakah Anda Ingin Mengubahnya dengan data yang baru ???", vbQuestion + vbYesNo) = vbYes Then
        DataMode = EN_LOAD_CHANGED
        Else
            MsgBox "Data Tidak Dirubah", vbInformation
            Exit Function
        End If
    Else
        oKelas.Add IDTahunAjaran, IDKelas, Kelas, Ruangan, JumlahSiswa
    End If
ElseIf DataMode = EN_LOAD_CHANGED Then
    oKelas.Edit IDTahunAjaran, IDKelas, Kelas, Ruangan, JumlahSiswa
End If
DataMode = EN_SAVED
SaveData = 1
End Function

Private Sub New_data()
    On Error Resume Next
    DataMode = EN_NEW
    txtidkel.text = ""
    txtjumlahsiswa.text = ""
    txtruangan.text = ""
    With txtkelas
        .Clear
        .AddItem "X"
        .AddItem "XI"
        .AddItem "XII"
    End With
    cariIDKelas
End Sub
Public Sub EditData(pNIP As Integer)
On Error Resume Next
If oKelas.Load(pNIP) > 0 Then
    txtidkel.text = oKelas.IDKelas
    txtjumlahsiswa.text = oKelas.JumlahSiswa
    txtruangan.text = oKelas.Ruangan
    txtkelas.text = oKelas.Kelas
Else
    MsgBox "Data tidak ditemukan", vbExclamation, "Load Data"
    Unload Me
End If
DataMode = EN_SAVED
End Sub
Private Sub ChangeData()
If DataMode = EN_NEW Then
    DataMode = EN_NEW_CHANGED
ElseIf DataMode = EN_SAVED Then
    DataMode = EN_LOAD_CHANGED
End If
End Sub

Private Sub txtjumlahsiswa_Change()
    ChangeData
End Sub

Private Sub txtkelas_Change()
    ChangeData
End Sub

Private Sub txtruangan_Change()
    ChangeData
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Hell
Select Case UCase(Button.key)
    Case "SIMPAN": Call mnuSave
    Case "NEW": Call Form_Load
    Case "EXIT":
        Unload Me
End Select
Exit Sub
Hell:
    MsgBox Err.Description, vbCritical, "Internal"
    Resume Next
End Sub
