VERSION 5.00
Begin VB.Form frmkonseling 
   Caption         =   "Bimbingan Dan Konseling"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Keluar"
      Height          =   375
      Left            =   7680
      TabIndex        =   7
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   6240
      TabIndex        =   6
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   5895
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   8775
      Begin VB.TextBox txtket 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Text            =   "frmkonseling.frx":0000
         Top             =   4680
         Width           =   3855
      End
      Begin VB.TextBox txtpenangan 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Text            =   "frmkonseling.frx":0006
         Top             =   3720
         Width           =   3855
      End
      Begin VB.TextBox txtmasalah 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Text            =   "frmkonseling.frx":000C
         Top             =   2760
         Width           =   3855
      End
      Begin VB.TextBox txtnamasiswa 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   2160
         Width           =   3855
      End
      Begin VB.TextBox txtnis 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox txtnamaguru 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "txtnamaguru"
         Top             =   1200
         Width           =   3855
      End
      Begin VB.TextBox txtnip 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   720
         Width           =   3855
      End
      Begin VB.TextBox txtid 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Keterangan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   240
         TabIndex        =   19
         Top             =   4680
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Penanganan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   18
         Top             =   3840
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Masalah"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   17
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Nama"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   16
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "NIS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   15
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Nama"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   14
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "NIP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmkonseling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private oBim As New DLLBK.cBK
Private oSis As New DLLBK.Csiswa
Private oGuru As New DLLBK.cGuru
Dim DataMode As ENUM_DATA_MODE
Dim id As Long


Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    Form_Load
End Sub

Private Sub Command3_Click()
    Call mnuSave
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
If SaveData("" & txtid.text, Now, "" & txtnis.text, "" & txtnip.text, "" & txtmasalah.text, "" & txtpenangan.text, "" & txtket.text) > 0 Then
    MsgBox "Data BERHASIL disimpan", vbInformation, "Simpan Data"
    Call New_data
    Form_Load
Else
    MsgBox "Data GAGAL disimpan", vbCritical, "Simpan Data"
End If
Exit Sub
Hell:
    MsgBox Err.Description, vbCritical, "Internal"
    'Resume Next
End Sub
Private Function SaveData(IDKonseling As Long, tgl As Date, NIS As Integer, NIP As String, _
Permasalahan As String, Penanganan As String, Keterangan As String) As Integer
If DataMode = EN_NEW_CHANGED Then
    oBim.Add IDKonseling, tgl, NIS, NIP, Permasalahan, Penanganan, Keterangan
ElseIf DataMode = EN_LOAD_CHANGED Then
    oBim.Edit IDKonseling, tgl, NIS, NIP, Permasalahan, Penanganan, Keterangan
End If
DataMode = EN_SAVED
SaveData = 1
End Function
Private Sub Form_Load()
    Set oBim = New DLLBK.cBK
    Set oSis = New DLLBK.Csiswa
    Set oGuru = New DLLBK.cGuru
    New_data
    oBim.List (True)
    id = Format(Now, "DD")
    id = id & Format(Now, "MM")
    id = id & Format(Now, "YYYY")
    txtid.text = id & oBim.Jumlah + 1
End Sub

Private Sub New_data()
    txtid.text = ""
    txtnamaguru.text = ""
    txtnamasiswa.text = ""
    txtnip.text = ""
    txtnis.text = ""
    txtmasalah.text = ""
    txtpenangan.text = ""
    txtket.text = ""
End Sub

Public Sub EditData(pID As Long)
On Error Resume Next
If oBim.Load(pID) > 0 Then
    txtnip.text = oBim.NIP
    txtnis.text = oBim.NIS
    txtmasalah.text = oBim.Permasalahan
    txtpenangan.text = oBim.Penanganan
    txtket.text = oBim.Keterangan
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

Private Sub txtket_Change()
    ChangeData
End Sub

Private Sub txtmasalah_Change()
    ChangeData
End Sub

Private Sub txtnamaguru_Change()
    ChangeData
End Sub

Private Sub txtnamasiswa_Change()
    ChangeData
End Sub

Private Sub txtnip_Change()
    ChangeData
End Sub

Private Sub txtnip_KeyPress(KeyAscii As Integer)
    If txtnip.text = "" Then Exit Sub
    If KeyAscii = 13 Then
        If oGuru.Load(txtnip.text) > 0 Then
            txtnamaguru.text = oGuru.Nama
            txtnis.SetFocus
        Else
            MsgBox "Data Guru Tidak Ditemukan", vbInformation
        End If
    End If
End Sub

Private Sub txtnip_LostFocus()
    Call txtnip_KeyPress(13)
End Sub

Private Sub txtnis_Change()
    ChangeData
End Sub

Private Sub txtnis_KeyPress(KeyAscii As Integer)
    If txtnis.text = "" Then Exit Sub
    If KeyAscii = 13 Then
        If oSis.Load(txtnis.text) > 0 Then
            txtnamasiswa.text = oSis.Nama
            txtmasalah.SetFocus
        Else
            MsgBox "Data Siswa Tidak Ditemukan", vbInformation
        End If
    End If
End Sub

Private Sub txtnis_LostFocus()
    Call txtnis_KeyPress(13)
End Sub

Private Sub txtpenangan_Change()
    ChangeData
End Sub
