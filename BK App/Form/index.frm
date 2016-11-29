VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm index 
   BackColor       =   &H8000000C&
   Caption         =   "Operasional Penanganan Masalah Siswa Dalam Sekolah"
   ClientHeight    =   7185
   ClientLeft      =   165
   ClientTop       =   840
   ClientWidth     =   10005
   Icon            =   "index.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "index.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "index.frx":7264
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "index.frx":DAC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "index.frx":14328
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "index.frx":1AB8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "index.frx":213EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "index.frx":27C4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "index.frx":2E4B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "index.frx":34D12
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Login"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Daftar Kelas"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Daftar Siswa"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Daftar Guru"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Daftar Soal"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Potensi Siswa"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Masalah Dan Penanganan"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Help"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Keluar"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnaplikas 
      Caption         =   "&Aplikasi"
      Begin VB.Menu MNLOGIN 
         Caption         =   "&Login"
      End
      Begin VB.Menu mnlogout 
         Caption         =   "&LogOut"
      End
      Begin VB.Menu e 
         Caption         =   "-"
      End
      Begin VB.Menu mnsetingtahunajaran 
         Caption         =   "Seting Tahun Ajaran"
         Shortcut        =   ^T
      End
      Begin VB.Menu l 
         Caption         =   "-"
      End
      Begin VB.Menu mnkeluar 
         Caption         =   "&Keluar"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnmaster 
      Caption         =   "&Master"
      Begin VB.Menu mnentrykelas 
         Caption         =   "Entry Kelas"
      End
      Begin VB.Menu mndaftarkelas 
         Caption         =   "Daftar KElas"
      End
      Begin VB.Menu u 
         Caption         =   "-"
      End
      Begin VB.Menu mnentrydatasiswa 
         Caption         =   "&Entry Data Siswa"
      End
      Begin VB.Menu mndaftarsiswa2 
         Caption         =   "Daftar Siswa"
      End
      Begin VB.Menu mndaftarsiswa 
         Caption         =   "&Daftar Konseling dan Potensi Siswa"
      End
      Begin VB.Menu r 
         Caption         =   "-"
      End
      Begin VB.Menu mnentrydataguru 
         Caption         =   "Entry Data &Guru"
      End
      Begin VB.Menu mndaftardataguru 
         Caption         =   "Daftar Data G&uru"
      End
      Begin VB.Menu t 
         Caption         =   "-"
      End
      Begin VB.Menu mnentrySoal 
         Caption         =   "Entry &Soal"
      End
      Begin VB.Menu mndaftarsoal 
         Caption         =   "Daftar Soal"
      End
   End
   Begin VB.Menu mnoperasional 
      Caption         =   "&Operasional"
      Begin VB.Menu mnlayanankonseling 
         Caption         =   "&Layanan Konseling"
      End
      Begin VB.Menu mndaftarkasusdan 
         Caption         =   "Dafar &Kasus Dan Penanggulangan"
      End
      Begin VB.Menu y 
         Caption         =   "-"
      End
      Begin VB.Menu mnpotensisiswa 
         Caption         =   "Potensi Siswa"
      End
   End
   Begin VB.Menu mnlaporan 
      Caption         =   "&Laporan"
      Begin VB.Menu mnlaporanpersiswa 
         Caption         =   "Laporan Daftar Konseling Tiap Siswa"
      End
   End
End
Attribute VB_Name = "index"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mndaftardataguru_Click()
    frmdaftarguru.Show
End Sub

Private Sub mndaftarkasusdan_Click()
    frmcaribim.Show
End Sub

Private Sub mndaftarkelas_Click()
    frmdaftarkelas.Show
End Sub

Private Sub mndaftarsiswa_Click()
'    frmdaftarsiswa.Show
    frmtree.Show
End Sub

Private Sub mndaftarsoal_Click()
    frmdaftarsoal.Show
End Sub

Private Sub mnentrydataguru_Click()
    frmguru.Show 1
End Sub

Private Sub mnentrydatasiswa_Click()
    frmentrysiswa.Show
End Sub

Private Sub mnentrykelas_Click()
    frmentryKelas.Show
End Sub

Private Sub mnentrySoal_Click()
    frmentrysoal.Show
End Sub

Private Sub mnkeluar_Click()
    End
End Sub

Private Sub mnlaporanpersiswa_Click()
    FRMSETUPPRINT.Show
End Sub

Private Sub mnlayanankonseling_Click()
    frmkonseling.Show
End Sub

Private Sub mnpotensisiswa_Click()
    frmjawab.Show
End Sub

Private Sub mnsetingtahunajaran_Click()
    frmtahunajaran.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 3:
        Case 5: frmdaftarkelas.Show
        Case 6: frmtree.Show
        Case 7: frmdaftarguru.Show
        Case 8: frmdaftarsoal.Show
        Case 9: frmjawab.Show
        Case 10: frmkonseling.Show
        Case 12
        Case 13: If MsgBox("Apakah Anda Yakin Akan Keluar???", vbYesNo + vbQuestion) = vbYes Then End
    End Select
End Sub
