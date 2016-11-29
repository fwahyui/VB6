VERSION 5.00
Begin VB.Form frmguru 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Entry Data Guru"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   5775
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   120
         TabIndex        =   12
         Top             =   2520
         Width           =   5415
         Begin VB.CommandButton Command3 
            Caption         =   "&Simpan"
            Height          =   375
            Left            =   240
            TabIndex        =   4
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton Command2 
            Caption         =   "&Batal"
            Height          =   375
            Left            =   2160
            TabIndex        =   5
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Keluar"
            Height          =   375
            Left            =   3960
            TabIndex        =   6
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.TextBox txtnip 
         Appearance      =   0  'Flat
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
         Left            =   1320
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   240
         Width           =   3855
      End
      Begin VB.TextBox txtjabatan 
         Appearance      =   0  'Flat
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
         Left            =   1320
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   2040
         Width           =   3855
      End
      Begin VB.TextBox txtalamat 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Text            =   "frmguru.frx":0000
         Top             =   1200
         Width           =   4095
      End
      Begin VB.TextBox txtnama 
         Appearance      =   0  'Flat
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
         Left            =   1320
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   720
         Width           =   3855
      End
      Begin VB.Label Label4 
         Caption         =   "Jabatan"
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
         Left            =   240
         TabIndex        =   11
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Alamat"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label2 
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
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "NIP"
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
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmguru"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private oGuru As New DLLBK.cGuru
Dim DataMode As ENUM_DATA_MODE

Private Sub Command1_Click()
    Unload Me
    Set oGuru = Nothing
End Sub

Private Sub Command2_Click()
    New_data
End Sub

Private Sub Command3_Click()
    Call mnuSave
End Sub

Private Sub Form_Load()
    Set oGuru = New cGuru
    Call New_data
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
If SaveData("" & txtnip.text, "" & txtnama.text, "" & txtalamat.text, "" & txtjabatan.text) > 0 Then
    MsgBox "Data BERHASIL disimpan", vbInformation, "Simpan Data"
    Call New_data
Else
    MsgBox "Data GAGAL disimpan", vbCritical, "Simpan Data"
End If
Exit Sub
Hell:
    MsgBox Err.Description, vbCritical, "Internal"
    'Resume Next
End Sub
Private Function SaveData(pNIP As String, pNama As String, pAlamat As String, pJabatan As String) As Integer
If DataMode = EN_NEW_CHANGED Then
    oGuru.Add pNIP, pNama, pAlamat, pJabatan
ElseIf DataMode = EN_LOAD_CHANGED Then
    oGuru.Edit pNIP, pNama, pAlamat, pJabatan
End If
DataMode = EN_SAVED
SaveData = 1
End Function

Private Sub New_data()
    On Error Resume Next
    DataMode = EN_NEW
    txtalamat.text = ""
    txtnama.text = ""
    txtnip.text = ""
    txtjabatan.text = ""
    txtnip.SetFocus
End Sub
Public Sub EditData(pNIP As String)
On Error Resume Next
If oGuru.Load(pNIP) > 0 Then
    txtalamat.text = oGuru.Alamat
    txtnama.text = oGuru.Nama
    txtnip.text = oGuru.NIP
    txtjabatan.text = oGuru.Jabatan
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

Private Sub txtalamat_Change()
    Call ChangeData
End Sub

Private Sub txtjabatan_Change()
    Call ChangeData
End Sub

Private Sub txtnama_Change()
    Call ChangeData
End Sub

Private Sub txtnip_Change()
    Call ChangeData
End Sub


