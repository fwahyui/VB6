VERSION 5.00
Begin VB.Form frmentrysoal 
   Caption         =   "Form1"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   7380
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Keluar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Simpan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Batal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFC0C0&
      Height          =   1695
      Left            =   120
      ScaleHeight     =   1635
      ScaleWidth      =   7035
      TabIndex        =   2
      Top             =   120
      Width           =   7095
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
         Height          =   735
         Left            =   1560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Text            =   "frmentrysoal.frx":0000
         Top             =   840
         Width           =   5415
      End
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
         Left            =   1560
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Soal"
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
         TabIndex        =   4
         Top             =   840
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NO"
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
         TabIndex        =   3
         Top             =   360
         Width           =   390
      End
   End
End
Attribute VB_Name = "frmentrysoal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private oSoal As New DLLBK.cSoal
Dim DataMode As ENUM_DATA_MODE

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    Call mnuSave
End Sub

Private Sub Command3_Click()
    Form_Load
End Sub

Private Sub Form_Load()
    Set oSoal = New DLLBK.cSoal
    Call New_data
End Sub
Private Sub New_data()
    txtnamasiswa.text = ""
    txtnis.text = ""
    DataMode = EN_NEW
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
If SaveData(txtnis.text, txtnamasiswa.text) > 0 Then
    MsgBox "Data BERHASIL disimpan", vbInformation, "Simpan Data"
    Form_Load
Else
    MsgBox "Data GAGAL disimpan", vbCritical, "Simpan Data"
End If
Exit Sub
Hell:
    MsgBox Err.Description, vbCritical, "Internal"
    'Resume Next
End Sub
Private Function SaveData(pNo As Integer, pSoal As String) As Integer
If DataMode = EN_NEW_CHANGED Then
    oSoal.Add pNo, pSoal
ElseIf DataMode = EN_LOAD_CHANGED Then
    oSoal.Edit pNo, pSoal
End If
DataMode = EN_SAVED
SaveData = 1
End Function

Public Sub EditData(pNo As Integer)
On Error Resume Next
If oSoal.Load(pNo) > 0 Then
    txtnis.text = oSoal.No
    txtnamasiswa.text = oSoal.Soal
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

Private Sub txtnamasiswa_Change()
    Call ChangeData
End Sub

Private Sub txtnis_Change()
    Call ChangeData
End Sub
