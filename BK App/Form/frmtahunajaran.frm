VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmtahunajaran 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Setting Tahun Ajaran"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFC0C0&
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   4875
      TabIndex        =   3
      Top             =   960
      Width           =   4935
      Begin VB.CommandButton Command1 
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   1680
         TabIndex        =   6
         Top             =   840
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   600
         TabIndex        =   4
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy"
         Format          =   57147395
         UpDown          =   -1  'True
         CurrentDate     =   39851
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   2760
         TabIndex        =   5
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy"
         Format          =   57147395
         UpDown          =   -1  'True
         CurrentDate     =   39851
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   2520
         X2              =   2280
         Y1              =   240
         Y2              =   600
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   1
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   5115
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5145
      Begin VB.Image Image1 
         Height          =   600
         Left            =   75
         Picture         =   "frmtahunajaran.frx":0000
         Stretch         =   -1  'True
         Top             =   75
         Width           =   570
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Setting Tahun Ajaran"
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
         Index           =   0
         Left            =   720
         TabIndex        =   2
         Top             =   120
         Width           =   7035
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Pilih Tahun Ajaran Yang Akan Digunakan"
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   1
         Top             =   480
         Width           =   4335
      End
   End
End
Attribute VB_Name = "frmtahunajaran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst As New ADODB.Recordset
Private oTahun As New DLLBK.cTahun
Private Sub Command1_Click()
Dim tahun As String
    tahun = DTPicker1.Year & "/" & DTPicker2.Year
    Set rst = Nothing
    rst.CursorLocation = adUseClient
    If oTahun.Load(tahun) = 0 Then
        oTahun.Add (tahun)
        MsgBox "Tahun Ajaran Telah Di Aktifkna", vbInformation
        PidTahun = oTahun.IDTahunAjaran
        Unload Me
    Else
        oTahun.Edit (tahun)
        MsgBox "Tahun Ajaran Telah Di Aktifkna", vbInformation
        PidTahun = oTahun.IDTahunAjaran
        Unload Me
    End If
    LihatTahunAktiv
End Sub

Private Sub Form_Load()
    Set oTahun = New DLLBK.cTahun
    DTPicker1.Year = Mid(strTahun, 1, 4)
    DTPicker2.Year = Mid(strTahun, 6, 4)
End Sub
