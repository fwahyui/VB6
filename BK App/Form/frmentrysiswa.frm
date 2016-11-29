VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmentrysiswa 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entry Data Siswa"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1095
      Index           =   1
      Left            =   0
      ScaleHeight     =   1065
      ScaleWidth      =   9465
      TabIndex        =   96
      TabStop         =   0   'False
      Top             =   0
      Width           =   9495
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Masukkan Semua Data Siswa"
         Height          =   375
         Index           =   2
         Left            =   1080
         TabIndex        =   98
         Top             =   600
         Width           =   4335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Entry Siswa"
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
         TabIndex        =   97
         Top             =   120
         Width           =   7035
      End
      Begin VB.Image Image1 
         Height          =   855
         Left            =   75
         Picture         =   "frmentrysiswa.frx":0000
         Top             =   75
         Width           =   810
      End
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
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7080
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
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7080
      Width           =   1335
   End
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
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7080
      Width           =   1335
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   10186
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Data Pribadi"
      TabPicture(0)   =   "frmentrysiswa.frx":24C8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Picture1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Orang Tua"
      TabPicture(1)   =   "frmentrysiswa.frx":24E4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Pendidikan"
      TabPicture(2)   =   "frmentrysiswa.frx":2500
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Prestasi"
      TabPicture(3)   =   "frmentrysiswa.frx":251C
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Picture5"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Kegiatan"
      TabPicture(4)   =   "frmentrysiswa.frx":2538
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Picture6"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Lain-lain"
      TabPicture(5)   =   "frmentrysiswa.frx":2554
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Picture7"
      Tab(5).ControlCount=   1
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFC0C0&
         Height          =   5295
         Index           =   0
         Left            =   60
         ScaleHeight     =   5235
         ScaleWidth      =   9075
         TabIndex        =   78
         Top             =   480
         Width           =   9135
         Begin VB.ComboBox Combo1 
            BackColor       =   &H00C0E0FF&
            Height          =   315
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   100
            Top             =   120
            Width           =   1695
         End
         Begin VB.TextBox txtnotelpsiswa 
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
            Left            =   2400
            TabIndex        =   86
            Text            =   "Text1"
            Top             =   4680
            Width           =   5535
         End
         Begin VB.TextBox txtalamatsiswa 
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
            Height          =   885
            Left            =   2400
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   85
            Text            =   "frmentrysiswa.frx":2570
            Top             =   3600
            Width           =   5535
         End
         Begin VB.ComboBox txtagama 
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
            ItemData        =   "frmentrysiswa.frx":2576
            Left            =   2400
            List            =   "frmentrysiswa.frx":2589
            Style           =   2  'Dropdown List
            TabIndex        =   84
            Top             =   3120
            Width           =   1455
         End
         Begin VB.OptionButton txtcew 
            BackColor       =   &H00FFC0C0&
            Caption         =   "&Perempuan"
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
            Left            =   4080
            TabIndex        =   83
            Top             =   2640
            Width           =   1815
         End
         Begin VB.OptionButton txtlaki 
            BackColor       =   &H00FFC0C0&
            Caption         =   "&Laki-laki"
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
            Left            =   2400
            TabIndex        =   82
            Top             =   2640
            Width           =   1695
         End
         Begin VB.TextBox txttempatlahir 
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
            Left            =   2400
            TabIndex        =   81
            Text            =   "Text1"
            Top             =   1560
            Width           =   5415
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
            Left            =   2400
            TabIndex        =   80
            Text            =   "Text1"
            Top             =   1080
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
            Left            =   2400
            TabIndex        =   79
            Text            =   "Text1"
            Top             =   600
            Width           =   2295
         End
         Begin MSComCtl2.DTPicker txttgllahir 
            Height          =   375
            Left            =   2400
            TabIndex        =   87
            Top             =   2160
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            CalendarBackColor=   16777215
            CalendarTitleBackColor=   16777215
            Format          =   96403457
            CurrentDate     =   39847
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "KELAS"
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
            Index           =   24
            Left            =   600
            TabIndex        =   99
            Top             =   120
            Width           =   735
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Telp"
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
            TabIndex        =   95
            Top             =   4680
            Width           =   810
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Height          =   255
            Left            =   600
            TabIndex        =   94
            Top             =   3600
            Width           =   675
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Agama"
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
            TabIndex        =   93
            Top             =   3120
            Width           =   660
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jenis Kelamin"
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
            TabIndex        =   92
            Top             =   2640
            Width           =   1305
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal Lahir"
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
            TabIndex        =   91
            Top             =   2160
            Width           =   1350
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tempat Lahir"
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
            TabIndex        =   90
            Top             =   1560
            Width           =   1305
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
            Left            =   600
            TabIndex        =   89
            Top             =   1080
            Width           =   585
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
            Left            =   600
            TabIndex        =   88
            Top             =   600
            Width           =   405
         End
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H00FFC0C0&
         Height          =   5295
         Left            =   -74940
         ScaleHeight     =   5235
         ScaleWidth      =   9075
         TabIndex        =   56
         Top             =   360
         Width           =   9135
         Begin VB.TextBox txtpenyakit3 
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
            Left            =   1680
            TabIndex        =   65
            Text            =   "Text1"
            Top             =   4680
            Width           =   3735
         End
         Begin VB.TextBox txtpenyakit2 
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
            Left            =   1680
            TabIndex        =   64
            Text            =   "Text1"
            Top             =   4200
            Width           =   3735
         End
         Begin VB.TextBox txtpenyakit1 
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
            Left            =   1680
            TabIndex        =   63
            Text            =   "Text1"
            Top             =   3720
            Width           =   3735
         End
         Begin VB.TextBox txtjumlahsaudarakandung 
            Alignment       =   1  'Right Justify
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
            Left            =   3720
            TabIndex        =   62
            Text            =   "Text1"
            Top             =   1200
            Width           =   1695
         End
         Begin VB.TextBox txtjumlahsaudara 
            Alignment       =   1  'Right Justify
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
            Left            =   3720
            TabIndex        =   61
            Text            =   "Text1"
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox txtanakke 
            Alignment       =   1  'Right Justify
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
            Left            =   3720
            TabIndex        =   60
            Text            =   "Text1"
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txttinggibadan 
            Alignment       =   1  'Right Justify
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
            Left            =   3720
            TabIndex        =   59
            Text            =   "Text1"
            Top             =   2640
            Width           =   1695
         End
         Begin VB.TextBox txtberatbadan 
            Alignment       =   1  'Right Justify
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
            Left            =   3720
            TabIndex        =   58
            Text            =   "Text1"
            Top             =   2160
            Width           =   1695
         End
         Begin VB.TextBox txtjumlahsaudaratiri 
            Alignment       =   1  'Right Justify
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
            Left            =   3720
            TabIndex        =   57
            Text            =   "Text1"
            Top             =   1680
            Width           =   1695
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Penyakit Yang Pernah Diderita Oleh Siswa"
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
            Left            =   1320
            TabIndex        =   77
            Top             =   3240
            Width           =   4080
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1"
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
            Index           =   9
            Left            =   1320
            TabIndex        =   76
            Top             =   3720
            Width           =   90
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2"
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
            Index           =   10
            Left            =   1320
            TabIndex        =   75
            Top             =   4200
            Width           =   105
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "3"
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
            Index           =   11
            Left            =   1320
            TabIndex        =   74
            Top             =   4680
            Width           =   105
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tinggi Badan"
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
            Index           =   1
            Left            =   1320
            TabIndex        =   73
            Top             =   2640
            Width           =   1290
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Berat Badan"
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
            Index           =   1
            Left            =   1320
            TabIndex        =   72
            Top             =   2160
            Width           =   1170
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Saudara Tiri"
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
            Index           =   1
            Left            =   1320
            TabIndex        =   71
            Top             =   1680
            Width           =   1155
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Saudara Kandung"
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
            Index           =   1
            Left            =   1320
            TabIndex        =   70
            Top             =   1200
            Width           =   1665
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Saudara"
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
            Index           =   1
            Left            =   1320
            TabIndex        =   69
            Top             =   720
            Width           =   1485
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Anak KE"
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
            Index           =   12
            Left            =   1320
            TabIndex        =   68
            Top             =   240
            Width           =   870
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kg"
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
            Index           =   13
            Left            =   5640
            TabIndex        =   67
            Top             =   2280
            Width           =   270
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cm"
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
            Index           =   14
            Left            =   5640
            TabIndex        =   66
            Top             =   2760
            Width           =   330
         End
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H00FFC0C0&
         Height          =   5295
         Left            =   -74940
         ScaleHeight     =   5235
         ScaleWidth      =   9075
         TabIndex        =   48
         Top             =   360
         Width           =   9135
         Begin VB.TextBox txtkegiatan2 
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
            Left            =   2040
            TabIndex        =   55
            Text            =   "Text1"
            Top             =   2400
            Width           =   6255
         End
         Begin VB.TextBox txtkegiatan3 
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
            Left            =   2040
            TabIndex        =   50
            Text            =   "Text1"
            Top             =   2880
            Width           =   6255
         End
         Begin VB.TextBox txtkegiatan1 
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
            Left            =   2040
            TabIndex        =   49
            Text            =   "Text1"
            Top             =   1920
            Width           =   6255
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Masukkan Kegiatan Yang Pernah Dilakukan Oleh Siswa"
            BeginProperty Font 
               Name            =   "Garamond"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1440
            TabIndex        =   54
            Top             =   1200
            Width           =   5880
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1"
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
            Index           =   6
            Left            =   1440
            TabIndex        =   53
            Top             =   1920
            Width           =   90
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2"
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
            Index           =   7
            Left            =   1440
            TabIndex        =   52
            Top             =   2400
            Width           =   105
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "3"
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
            Index           =   8
            Left            =   1440
            TabIndex        =   51
            Top             =   2880
            Width           =   105
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00FFC0C0&
         Height          =   5295
         Left            =   -74940
         ScaleHeight     =   5235
         ScaleWidth      =   9075
         TabIndex        =   36
         Top             =   360
         Width           =   9135
         Begin VB.TextBox txtprestasi3 
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
            Left            =   1680
            TabIndex        =   41
            Text            =   "Text1"
            Top             =   2520
            Width           =   7095
         End
         Begin VB.TextBox txtprestasi2 
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
            Left            =   1680
            TabIndex        =   40
            Text            =   "Text1"
            Top             =   2040
            Width           =   7095
         End
         Begin VB.TextBox txtprestasi1 
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
            Left            =   1680
            TabIndex        =   39
            Text            =   "Text1"
            Top             =   1560
            Width           =   7095
         End
         Begin VB.TextBox txtprestasi5 
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
            Left            =   1680
            TabIndex        =   38
            Text            =   "Text1"
            Top             =   3480
            Width           =   7095
         End
         Begin VB.TextBox txtprestasi4 
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
            Left            =   1680
            TabIndex        =   37
            Text            =   "Text1"
            Top             =   3000
            Width           =   7095
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Masukkan Prestasi Yang Pernah Diraih Oleh Siswa"
            BeginProperty Font 
               Name            =   "Garamond"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1080
            TabIndex        =   47
            Top             =   720
            Width           =   5340
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1"
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
            Index           =   1
            Left            =   1080
            TabIndex        =   46
            Top             =   1560
            Width           =   90
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2"
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
            Index           =   2
            Left            =   1080
            TabIndex        =   45
            Top             =   2040
            Width           =   105
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "3"
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
            Index           =   3
            Left            =   1080
            TabIndex        =   44
            Top             =   2520
            Width           =   105
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "4"
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
            Index           =   4
            Left            =   1080
            TabIndex        =   43
            Top             =   3000
            Width           =   105
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "5"
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
            Index           =   5
            Left            =   1080
            TabIndex        =   42
            Top             =   3480
            Width           =   105
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00FFC0C0&
         Height          =   5295
         Left            =   -74940
         ScaleHeight     =   5235
         ScaleWidth      =   9075
         TabIndex        =   17
         Top             =   360
         Width           =   9135
         Begin VB.Frame Frame3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Formal"
            BeginProperty Font 
               Name            =   "Garamond"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   3255
            Left            =   120
            TabIndex        =   27
            Top             =   840
            Width           =   4215
            Begin VB.TextBox txtsd 
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
               Left            =   1200
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   31
               Text            =   "frmentrysiswa.frx":25B4
               Top             =   360
               Width           =   2895
            End
            Begin VB.TextBox txtalamatsd 
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
               Height          =   765
               Left            =   1200
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   30
               Text            =   "frmentrysiswa.frx":25BA
               Top             =   840
               Width           =   2895
            End
            Begin VB.TextBox txtsmp 
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
               Left            =   1200
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   29
               Text            =   "frmentrysiswa.frx":25C0
               Top             =   1680
               Width           =   2895
            End
            Begin VB.TextBox txtalamatsmp 
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
               Height          =   765
               Left            =   1200
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   28
               Text            =   "frmentrysiswa.frx":25C6
               Top             =   2160
               Width           =   2895
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "SD"
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
               Index           =   15
               Left            =   240
               TabIndex        =   35
               Top             =   360
               Width           =   300
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
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
               Height          =   255
               Index           =   16
               Left            =   240
               TabIndex        =   34
               Top             =   840
               Width           =   675
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "SMP"
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
               Index           =   17
               Left            =   240
               TabIndex        =   33
               Top             =   1680
               Width           =   465
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
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
               Height          =   255
               Index           =   18
               Left            =   240
               TabIndex        =   32
               Top             =   2160
               Width           =   675
            End
         End
         Begin VB.Frame Frame4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Non-Formal"
            BeginProperty Font 
               Name            =   "Garamond"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   3255
            Left            =   4440
            TabIndex        =   18
            Top             =   840
            Width           =   4575
            Begin VB.TextBox txtnonformal1 
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
               Left            =   1920
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   22
               Text            =   "frmentrysiswa.frx":25CC
               Top             =   360
               Width           =   2535
            End
            Begin VB.TextBox txtalamatnonformal1 
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
               Height          =   765
               Left            =   1920
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   21
               Text            =   "frmentrysiswa.frx":25D2
               Top             =   840
               Width           =   2535
            End
            Begin VB.TextBox txtnonformal2 
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
               Left            =   1920
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   20
               Text            =   "frmentrysiswa.frx":25D8
               Top             =   1680
               Width           =   2535
            End
            Begin VB.TextBox txtalamatnonformal2 
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
               Height          =   765
               Left            =   1920
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   19
               Text            =   "frmentrysiswa.frx":25DE
               Top             =   2160
               Width           =   2535
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Non Formal 1"
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
               Index           =   19
               Left            =   240
               TabIndex        =   26
               Top             =   360
               Width           =   1320
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
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
               Height          =   255
               Index           =   20
               Left            =   240
               TabIndex        =   25
               Top             =   840
               Width           =   675
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
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
               Height          =   255
               Index           =   22
               Left            =   240
               TabIndex        =   24
               Top             =   2160
               Width           =   675
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Non Formal 2"
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
               Index           =   21
               Left            =   240
               TabIndex        =   23
               Top             =   1680
               Width           =   1335
            End
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFC0C0&
         Height          =   5295
         Left            =   -74940
         ScaleHeight     =   5235
         ScaleWidth      =   9075
         TabIndex        =   4
         Top             =   360
         Width           =   9135
         Begin VB.TextBox txtnamaortu 
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
            Left            =   2160
            TabIndex        =   10
            Text            =   "Text1"
            Top             =   960
            Width           =   5415
         End
         Begin VB.TextBox txthubungan 
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
            Left            =   2160
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   1440
            Width           =   5415
         End
         Begin VB.TextBox txtpendidikanortu 
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
            Left            =   2160
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   3360
            Width           =   5415
         End
         Begin VB.ComboBox txtagamaortu 
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
            ItemData        =   "frmentrysiswa.frx":25E4
            Left            =   2160
            List            =   "frmentrysiswa.frx":25F7
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1920
            Width           =   1695
         End
         Begin VB.TextBox txtalamatortu 
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
            Height          =   885
            Left            =   2160
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Text            =   "frmentrysiswa.frx":2622
            Top             =   2400
            Width           =   5535
         End
         Begin VB.TextBox txtpekerjaanortu 
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
            Left            =   2160
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   3840
            Width           =   5415
         End
         Begin VB.Label Label16 
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
            Left            =   720
            TabIndex        =   16
            Top             =   960
            Width           =   585
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hubungan"
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
            Left            =   720
            TabIndex        =   15
            Top             =   1440
            Width           =   1020
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Height          =   255
            Left            =   720
            TabIndex        =   14
            Top             =   2400
            Width           =   675
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Agama"
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
            Left            =   720
            TabIndex        =   13
            Top             =   1920
            Width           =   660
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pendidikan"
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
            Left            =   720
            TabIndex        =   12
            Top             =   3360
            Width           =   1065
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pekerjaan"
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
            Left            =   720
            TabIndex        =   11
            Top             =   3840
            Width           =   930
         End
      End
   End
End
Attribute VB_Name = "frmentrysiswa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private oSiswa As New DLLBK.Csiswa
Private oKelas As New DLLBK.Ckelas
Dim IDD As Integer
Dim DataMode As ENUM_DATA_MODE
Private Sub ChangeData()
If DataMode = EN_NEW Then
    DataMode = EN_NEW_CHANGED
ElseIf DataMode = EN_SAVED Then
    DataMode = EN_LOAD_CHANGED
End If
End Sub

Private Sub mnuSave()
Dim jeniskelamin As String
 If txtcew.Value = True Then
    jeniskelamin = "P"
 Else
    jeniskelamin = "L"
 End If
On Error GoTo Hell
If DataMode = EN_NEW Then
    MsgBox "Data harus diisi dulu" & vbCrLf & "Simpan data dibatalkan", vbExclamation, "Simpan Data"
    Exit Sub
ElseIf DataMode = EN_SAVED Then
    MsgBox "Tidak ada data yang berubah" & vbCrLf & "Simpan data dibatalkan", vbExclamation, "Simpan Data"
    Exit Sub
End If
If SaveData(txtnis.text, "" & txtnamasiswa.text, "" & txttempatlahir.text, _
txttgllahir.Value, "" & jeniskelamin, "" & txtagama.text, "" & txtalamatsiswa.text, _
"" & txtnotelpsiswa.text, "" & txtnamaortu.text, "" & txthubungan.text, "" & txtagamaortu.text, _
"" & txtalamatortu.text, "" & txtpendidikanortu.text, "" & txtpekerjaanortu.text, "" & txtsd.text, _
"" & txtalamatsd.text, "" & txtsmp.text, "" & txtalamatsmp.text, "" & txtnonformal1.text, _
"" & txtalamatnonformal1.text, "" & txtnonformal2.text, "" & txtalamatnonformal2.text, _
"" & txtprestasi1.text, "" & txtprestasi2.text, "" & txtprestasi3.text, "" & txtprestasi4.text, _
"" & txtprestasi5.text, "" & txtkegiatan1.text, "" & txtkegiatan2.text, "" & txtkegiatan3.text, _
0 & txtanakke.text, 0 & txtjumlahsaudara.text, 0 & txtjumlahsaudarakandung.text, _
0 & txtjumlahsaudaratiri.text, 0 & txtberatbadan.text, 0 & txttinggibadan.text, _
0 & txtpenyakit1.text, 0 & txtpenyakit2.text, 0 & txtpenyakit3.text) = 1 Then
    
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
Private Function SaveData(NIS As Long, Nama As String, _
Tempatlahir As String, tgllahir As Date, jk As String, agama As String, Alamat As String, Notelp As String, _
namaortu As String, hubungan As String, agamaortu As String, alamatortu As String, pendidikanortu As String, pekerjaanortu As String, _
sd As String, alamatsd As String, smp As String, alamatsmp As String, _
nonformalsatu As String, alamatnonformalsatu As String, nonformaldua As String, _
alamatnonformaldua As String, _
prestasi As String, prestasidua As String, prestasitiga As String, prestasiempat As String, _
prestasilima As String, kegiatan As String, kegiatan1 As String, kegiatan2 As String, anakke As Integer, _
jumlahsaudara As Integer, saudarakandung As Integer, saudaratiri As Integer, beratbadan As Integer, tinggibandan As Integer, penyakit1 As String, _
penyakit2 As String, penyakit3 As String, IDKelas As Integer) As Integer

If DataMode = EN_NEW_CHANGED Then
    oSiswa.Add NIS, Nama, Tempatlahir, tgllahir, jk, agama, Alamat, Notelp, namaortu, hubungan, agamaortu, alamatortu, pendidikanortu, pekerjaanortu, sd, alamatsd, smp, alamatsmp, nonformalsatu, alamatnonformalsatu, nonformaldua, alamatnonformaldua, prestasi, prestasidua, prestasitiga, prestasiempat, prestasilima, kegiatan, kegiatan1, kegiatan2, anakke, jumlahsaudara, saudarakandung, saudaratiri, beratbadan, tinggibandan, penyakit1, penyakit2, penyakit3, IDD
ElseIf DataMode = EN_LOAD_CHANGED Then
    oSiswa.Edit NIS, Nama, Tempatlahir, tgllahir, jk, agama, Alamat, Notelp, namaortu, hubungan, agamaortu, alamatortu, pendidikanortu, pekerjaanortu, sd, alamatsd, smp, alamatsmp, nonformalsatu, alamatnonformalsatu, nonformaldua, alamatnonformaldua, prestasi, prestasidua, prestasitiga, prestasiempat, prestasilima, kegiatan, kegiatan1, kegiatan2, anakke, jumlahsaudara, saudarakandung, saudaratiri, beratbadan, tinggibandan, penyakit1, penyakit2, penyakit3, IDD
End If
DataMode = EN_SAVED
SaveData = 1

End Function


Public Sub EditData(pnis As Long)
On Error Resume Next
If oSiswa.Load(pnis) > 0 Then
    txtagama.text = oSiswa.agama
    txtagamaortu.text = oSiswa.agamaortu
    txtalamatnonformal1.text = oSiswa.alamatnonformalsatu
    txtalamatnonformal2.text = oSiswa.alamatnonformaldua
    txtalamatortu.text = oSiswa.alamatortu
    txtalamatsd.text = oSiswa.alamatsd
    txtalamatsiswa.text = oSiswa.Alamat
    txtalamatsmp.text = oSiswa.alamatsmp
    txtanakke.text = oSiswa.anakke
    txtberatbadan.text = oSiswa.beratbadan
    If oSiswa.jk = "L" Then
        txtlaki.Value = True
    Else
        txtcew.Value = True
    End If
    txthubungan.text = oSiswa.hubungan
    txtjumlahsaudara.text = oSiswa.jumlahsaudara
    txtjumlahsaudarakandung.text = oSiswa.saudarakandung
    txtjumlahsaudaratiri.text = oSiswa.saudaratiri
    txtkegiatan1.text = oSiswa.kegiatan
    txtkegiatan2.text = oSiswa.kegiatan1
    txtkegiatan3.text = oSiswa.kegiatan2
    txtnamaortu.text = oSiswa.namaortu
    txtnamasiswa.text = oSiswa.Nama
    txtnis.text = oSiswa.NIS
    txtnonformal1.text = oSiswa.nonformalsatu
    txtnonformal2.text = oSiswa.nonformaldua
    txtnotelpsiswa.text = oSiswa.Notelp
    txtpekerjaanortu.text = oSiswa.pekerjaanortu
    txtpendidikanortu.text = oSiswa.pendidikanortu
    txtpenyakit1.text = oSiswa.penyakit1
    txtpenyakit2.text = oSiswa.penyakit2
    txtpenyakit3.text = oSiswa.penyakit3
    txtprestasi1.text = oSiswa.prestasi
    txtprestasi2.text = oSiswa.prestasidua
    txtprestasi3.text = oSiswa.prestasitiga
    txtprestasi4.text = oSiswa.prestasiempat
    txtprestasi5.text = oSiswa.prestasilima
    txtsd.text = oSiswa.sd
    txtsmp.text = oSiswa.smp
    txttempatlahir.text = oSiswa.Tempatlahir
    txttgllahir.Value = oSiswa.tgllahir
    txttinggibadan.text = oSiswa.tinggibandan
Else
    MsgBox "Data tidak ditemukan", vbExclamation, "Load Data"
    Unload Me
End If
DataMode = EN_SAVED
End Sub

Private Sub New_data()
    DataMode = EN_NEW
    Call IsiCombo(PidTahun)
    txtalamatnonformal1.text = ""
    txtalamatnonformal2.text = ""
    txtalamatortu.text = ""
    txtalamatsd.text = ""
    txtalamatsiswa.text = ""
    txtalamatsmp.text = ""
    txtanakke.text = ""
    txtberatbadan.text = ""
    txtcew.Value = False
    txthubungan.text = ""
    txtjumlahsaudara.text = ""
    txtjumlahsaudarakandung.text = ""
    txtjumlahsaudaratiri.text = ""
    txtkegiatan1.text = ""
    txtkegiatan2.text = ""
    txtkegiatan3.text = ""
    txtlaki.Value = False
    txtnamaortu.text = ""
    txtnamasiswa.text = ""
    txtnis.text = ""
    txtnonformal1.text = ""
    txtnonformal2.text = ""
    txtnotelpsiswa.text = ""
    txtpekerjaanortu.text = ""
    txtpendidikanortu.text = ""
    txtpenyakit1.text = ""
    txtpenyakit2.text = ""
    txtpenyakit3.text = ""
    txtprestasi1.text = ""
    txtprestasi2.text = ""
    txtprestasi3.text = ""
    txtprestasi4.text = ""
    txtprestasi5.text = ""
    txtsd.text = ""
    txtsmp.text = ""
    txttempatlahir.text = ""
    txttgllahir.Value = Now
    txttinggibadan.text = ""
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    If Combo1.text = "" Then
        MsgBox "Kelas Siswa Belum Dipilih", vbInformation
        Exit Sub
    Else
        mnuSave
    End If
End Sub

Private Sub Command3_Click()
    New_data
End Sub

Private Sub Form_Load()
    Call form_ditengah(Index, Me)
    Call New_data
    Set oSiswa = New DLLBK.Csiswa
    Set oKelas = New DLLBK.Ckelas
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, y As Single)

End Sub

Private Sub SSTab1_DblClick()
    Call ChangeData
End Sub

Private Sub txtagama_Change()
    Call ChangeData
End Sub

Private Sub txtagamaortu_Change()
    Call ChangeData
End Sub

Private Sub txtalamatnonformal1_Change()
    Call ChangeData
End Sub

Private Sub txtalamatnonformal2_Change()
    Call ChangeData
End Sub

Private Sub txtalamatortu_Change()
    Call ChangeData
End Sub

Private Sub txtalamatsd_Change()
    Call ChangeData
End Sub

Private Sub txtalamatsiswa_Change()
    Call ChangeData
End Sub

Private Sub txtalamatsmp_Change()
    Call ChangeData
End Sub

Private Sub txtanakke_Change()
    Call ChangeData
End Sub

Private Sub txtberatbadan_Change()
    Call ChangeData
End Sub

Private Sub txtcew_Click()
    Call ChangeData
End Sub

Private Sub txthubungan_Change()
    Call ChangeData
End Sub

Private Sub txtjumlahsaudara_Change()
    Call ChangeData
End Sub

Private Sub txtjumlahsaudarakandung_Change()
    Call ChangeData
End Sub

Private Sub txtjumlahsaudaratiri_Change()
    Call ChangeData
End Sub

Private Sub txtkegiatan1_Change()
    Call ChangeData
End Sub

Private Sub txtkegiatan3_Change()
    Call ChangeData
End Sub

Private Sub txtlaki_Click()
    Call ChangeData
End Sub

Private Sub txtnamaortu_Change()
    Call ChangeData
End Sub

Private Sub txtnamasiswa_Change()
    Call ChangeData
End Sub

Private Sub txtnis_Change()
    Call ChangeData
End Sub

Private Sub txtnonformal1_Change()
    Call ChangeData
End Sub

Private Sub txtnonformal2_Change()
    Call ChangeData
End Sub

Private Sub txtnotelpsiswa_Change()
    Call ChangeData
End Sub

Private Sub txtpekerjaanortu_Change()
    Call ChangeData
End Sub

Private Sub txtpendidikanortu_Change()
    Call ChangeData
End Sub

Private Sub txtpenyakit1_Change()
    Call ChangeData
End Sub

Private Sub txtpenyakit2_Change()
    Call ChangeData
End Sub

Private Sub txtpenyakit3_Change()
    Call ChangeData
End Sub

Private Sub txtprestasi1_Change()
    Call ChangeData
End Sub

Private Sub txtprestasi2_Change()
    Call ChangeData
End Sub

Private Sub txtprestasi3_Change()
    Call ChangeData
End Sub

Private Sub txtprestasi4_Change()
    Call ChangeData
End Sub

Private Sub txtprestasi5_Change()
    Call ChangeData
End Sub

Private Sub txtsd_Change()
    Call ChangeData
End Sub

Private Sub txtsmp_Change()
    Call ChangeData
End Sub

Private Sub txttempatlahir_Change()
    Call ChangeData
End Sub

Private Sub txttgllahir_Click()
    Call ChangeData
End Sub

Private Sub txttinggibadan_Change()
    Call ChangeData
End Sub
Private Sub IsiCombo(PidTahun As Integer)
Dim rsss As New ADODB.Recordset
    Set rsss = Nothing
    rsss.CursorLocation = adUseClient
        Set rsss = oKelas.ListCombo("where IDTahunAjaran = " & PidTahun & "")
        For i = 0 To rsss.RecordCount - 1
            Combo1.AddItem rsss!Kelas & "-" & rsss!Ruangan
            rsss.MoveNext
        Next
End Sub
Private Sub carikelas()

End Sub
