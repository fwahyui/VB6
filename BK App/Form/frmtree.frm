VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{70709D0B-CC7E-4203-B043-629D9B84D0A8}#3.0#0"; "XChart.ocx"
Begin VB.Form frmtree 
   AutoRedraw      =   -1  'True
   Caption         =   "Data Siswa"
   ClientHeight    =   9765
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   14430
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9765
   ScaleWidth      =   14430
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "&Refresh Data"
      Height          =   255
      Left            =   120
      TabIndex        =   146
      Top             =   840
      Width           =   3735
   End
   Begin VB.PictureBox Picture4 
      Align           =   1  'Align Top
      BackColor       =   &H00808080&
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   14370
      TabIndex        =   145
      Top             =   0
      Width           =   14430
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9495
      Left            =   3960
      TabIndex        =   1
      Top             =   840
      Width           =   10845
      _ExtentX        =   19129
      _ExtentY        =   16748
      _Version        =   393216
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Bimbingan"
      TabPicture(0)   =   "frmtree.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Jawaban Siswa"
      TabPicture(1)   =   "frmtree.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Potensi Siswa"
      TabPicture(2)   =   "frmtree.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "XChart1"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   6855
         Left            =   -74760
         TabIndex        =   7
         Top             =   480
         Width           =   7815
         Begin VB.TextBox th 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6960
            TabIndex        =   143
            Top             =   6360
            Width           =   735
         End
         Begin VB.TextBox tg 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6120
            TabIndex        =   141
            Top             =   6360
            Width           =   735
         End
         Begin VB.TextBox tf 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5160
            TabIndex        =   139
            Top             =   6360
            Width           =   735
         End
         Begin VB.TextBox te 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4200
            TabIndex        =   137
            Top             =   6360
            Width           =   735
         End
         Begin VB.TextBox td 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3240
            TabIndex        =   135
            Top             =   6360
            Width           =   735
         End
         Begin VB.TextBox tc 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2160
            TabIndex        =   133
            Top             =   6360
            Width           =   735
         End
         Begin VB.TextBox tb 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1200
            TabIndex        =   131
            Top             =   6360
            Width           =   735
         End
         Begin VB.TextBox ta 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            TabIndex        =   129
            Top             =   6360
            Width           =   735
         End
         Begin VB.CheckBox yesno 
            Caption         =   "8"
            Height          =   255
            Index           =   119
            Left            =   7080
            TabIndex        =   127
            Top             =   5520
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "7"
            Height          =   255
            Index           =   118
            Left            =   6240
            TabIndex        =   126
            Top             =   5520
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "6"
            Height          =   255
            Index           =   117
            Left            =   5280
            TabIndex        =   125
            Top             =   5520
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "5"
            Height          =   255
            Index           =   116
            Left            =   4320
            TabIndex        =   124
            Top             =   5520
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "4"
            Height          =   255
            Index           =   115
            Left            =   3360
            TabIndex        =   123
            Top             =   5520
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "3"
            Height          =   255
            Index           =   114
            Left            =   2280
            TabIndex        =   122
            Top             =   5520
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "2"
            Height          =   255
            Index           =   113
            Left            =   1320
            TabIndex        =   121
            Top             =   5520
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "1"
            Height          =   255
            Index           =   112
            Left            =   240
            TabIndex        =   120
            Top             =   5520
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "8"
            Height          =   255
            Index           =   111
            Left            =   7080
            TabIndex        =   119
            Top             =   5160
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "7"
            Height          =   255
            Index           =   110
            Left            =   6240
            TabIndex        =   118
            Top             =   5160
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "6"
            Height          =   255
            Index           =   109
            Left            =   5280
            TabIndex        =   117
            Top             =   5160
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "5"
            Height          =   255
            Index           =   108
            Left            =   4320
            TabIndex        =   116
            Top             =   5160
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "4"
            Height          =   255
            Index           =   107
            Left            =   3360
            TabIndex        =   115
            Top             =   5160
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "3"
            Height          =   255
            Index           =   106
            Left            =   2280
            TabIndex        =   114
            Top             =   5160
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "2"
            Height          =   255
            Index           =   105
            Left            =   1320
            TabIndex        =   113
            Top             =   5160
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "1"
            Height          =   255
            Index           =   104
            Left            =   240
            TabIndex        =   112
            Top             =   5160
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "8"
            Height          =   255
            Index           =   103
            Left            =   7080
            TabIndex        =   111
            Top             =   4800
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "7"
            Height          =   255
            Index           =   102
            Left            =   6240
            TabIndex        =   110
            Top             =   4800
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "6"
            Height          =   255
            Index           =   101
            Left            =   5280
            TabIndex        =   109
            Top             =   4800
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "5"
            Height          =   255
            Index           =   100
            Left            =   4320
            TabIndex        =   108
            Top             =   4800
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "4"
            Height          =   255
            Index           =   99
            Left            =   3360
            TabIndex        =   107
            Top             =   4800
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "3"
            Height          =   255
            Index           =   98
            Left            =   2280
            TabIndex        =   106
            Top             =   4800
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "2"
            Height          =   255
            Index           =   97
            Left            =   1320
            TabIndex        =   105
            Top             =   4800
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "1"
            Height          =   255
            Index           =   96
            Left            =   240
            TabIndex        =   104
            Top             =   4800
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "8"
            Height          =   255
            Index           =   95
            Left            =   7080
            TabIndex        =   103
            Top             =   4440
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "7"
            Height          =   255
            Index           =   94
            Left            =   6240
            TabIndex        =   102
            Top             =   4440
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "6"
            Height          =   255
            Index           =   93
            Left            =   5280
            TabIndex        =   101
            Top             =   4440
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "5"
            Height          =   255
            Index           =   92
            Left            =   4320
            TabIndex        =   100
            Top             =   4440
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "4"
            Height          =   255
            Index           =   91
            Left            =   3360
            TabIndex        =   99
            Top             =   4440
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "3"
            Height          =   255
            Index           =   90
            Left            =   2280
            TabIndex        =   98
            Top             =   4440
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "2"
            Height          =   255
            Index           =   89
            Left            =   1320
            TabIndex        =   97
            Top             =   4440
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "1"
            Height          =   255
            Index           =   88
            Left            =   240
            TabIndex        =   96
            Top             =   4440
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "8"
            Height          =   255
            Index           =   87
            Left            =   7080
            TabIndex        =   95
            Top             =   4080
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "7"
            Height          =   255
            Index           =   86
            Left            =   6240
            TabIndex        =   94
            Top             =   4080
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "6"
            Height          =   255
            Index           =   85
            Left            =   5280
            TabIndex        =   93
            Top             =   4080
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "5"
            Height          =   255
            Index           =   84
            Left            =   4320
            TabIndex        =   92
            Top             =   4080
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "4"
            Height          =   255
            Index           =   83
            Left            =   3360
            TabIndex        =   91
            Top             =   4080
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "3"
            Height          =   255
            Index           =   82
            Left            =   2280
            TabIndex        =   90
            Top             =   4080
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "2"
            Height          =   255
            Index           =   81
            Left            =   1320
            TabIndex        =   89
            Top             =   4080
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "1"
            Height          =   255
            Index           =   80
            Left            =   240
            TabIndex        =   88
            Top             =   4080
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "8"
            Height          =   255
            Index           =   79
            Left            =   7080
            TabIndex        =   87
            Top             =   3720
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "7"
            Height          =   255
            Index           =   78
            Left            =   6240
            TabIndex        =   86
            Top             =   3720
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "6"
            Height          =   255
            Index           =   77
            Left            =   5280
            TabIndex        =   85
            Top             =   3720
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "5"
            Height          =   255
            Index           =   76
            Left            =   4320
            TabIndex        =   84
            Top             =   3720
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "4"
            Height          =   255
            Index           =   75
            Left            =   3360
            TabIndex        =   83
            Top             =   3720
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "3"
            Height          =   255
            Index           =   74
            Left            =   2280
            TabIndex        =   82
            Top             =   3720
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "2"
            Height          =   255
            Index           =   73
            Left            =   1320
            TabIndex        =   81
            Top             =   3720
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "1"
            Height          =   255
            Index           =   72
            Left            =   240
            TabIndex        =   80
            Top             =   3720
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "8"
            Height          =   255
            Index           =   71
            Left            =   7080
            TabIndex        =   79
            Top             =   3360
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "7"
            Height          =   255
            Index           =   70
            Left            =   6240
            TabIndex        =   78
            Top             =   3360
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "6"
            Height          =   255
            Index           =   69
            Left            =   5280
            TabIndex        =   77
            Top             =   3360
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "5"
            Height          =   255
            Index           =   68
            Left            =   4320
            TabIndex        =   76
            Top             =   3360
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "4"
            Height          =   255
            Index           =   67
            Left            =   3360
            TabIndex        =   75
            Top             =   3360
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "3"
            Height          =   255
            Index           =   66
            Left            =   2280
            TabIndex        =   74
            Top             =   3360
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "2"
            Height          =   255
            Index           =   65
            Left            =   1320
            TabIndex        =   73
            Top             =   3360
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "1"
            Height          =   255
            Index           =   64
            Left            =   240
            TabIndex        =   72
            Top             =   3360
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "8"
            Height          =   255
            Index           =   63
            Left            =   7080
            TabIndex        =   71
            Top             =   3000
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "7"
            Height          =   255
            Index           =   62
            Left            =   6240
            TabIndex        =   70
            Top             =   3000
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "6"
            Height          =   255
            Index           =   61
            Left            =   5280
            TabIndex        =   69
            Top             =   3000
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "5"
            Height          =   255
            Index           =   60
            Left            =   4320
            TabIndex        =   68
            Top             =   3000
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "4"
            Height          =   255
            Index           =   59
            Left            =   3360
            TabIndex        =   67
            Top             =   3000
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "3"
            Height          =   255
            Index           =   58
            Left            =   2280
            TabIndex        =   66
            Top             =   3000
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "2"
            Height          =   255
            Index           =   57
            Left            =   1320
            TabIndex        =   65
            Top             =   3000
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "1"
            Height          =   255
            Index           =   56
            Left            =   240
            TabIndex        =   64
            Top             =   3000
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "8"
            Height          =   255
            Index           =   55
            Left            =   7080
            TabIndex        =   63
            Top             =   2640
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "7"
            Height          =   255
            Index           =   54
            Left            =   6240
            TabIndex        =   62
            Top             =   2640
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "6"
            Height          =   255
            Index           =   53
            Left            =   5280
            TabIndex        =   61
            Top             =   2640
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "5"
            Height          =   255
            Index           =   52
            Left            =   4320
            TabIndex        =   60
            Top             =   2640
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "4"
            Height          =   255
            Index           =   51
            Left            =   3360
            TabIndex        =   59
            Top             =   2640
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "3"
            Height          =   255
            Index           =   50
            Left            =   2280
            TabIndex        =   58
            Top             =   2640
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "2"
            Height          =   255
            Index           =   49
            Left            =   1320
            TabIndex        =   57
            Top             =   2640
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "1"
            Height          =   255
            Index           =   48
            Left            =   240
            TabIndex        =   56
            Top             =   2640
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "8"
            Height          =   255
            Index           =   47
            Left            =   7080
            TabIndex        =   55
            Top             =   2280
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "7"
            Height          =   255
            Index           =   46
            Left            =   6240
            TabIndex        =   54
            Top             =   2280
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "6"
            Height          =   255
            Index           =   45
            Left            =   5280
            TabIndex        =   53
            Top             =   2280
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "5"
            Height          =   255
            Index           =   44
            Left            =   4320
            TabIndex        =   52
            Top             =   2280
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "4"
            Height          =   255
            Index           =   43
            Left            =   3360
            TabIndex        =   51
            Top             =   2280
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "3"
            Height          =   255
            Index           =   42
            Left            =   2280
            TabIndex        =   50
            Top             =   2280
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "2"
            Height          =   255
            Index           =   41
            Left            =   1320
            TabIndex        =   49
            Top             =   2280
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "1"
            Height          =   255
            Index           =   40
            Left            =   240
            TabIndex        =   48
            Top             =   2280
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "8"
            Height          =   255
            Index           =   39
            Left            =   7080
            TabIndex        =   47
            Top             =   1920
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "7"
            Height          =   255
            Index           =   38
            Left            =   6240
            TabIndex        =   46
            Top             =   1920
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "6"
            Height          =   255
            Index           =   37
            Left            =   5280
            TabIndex        =   45
            Top             =   1920
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "5"
            Height          =   255
            Index           =   36
            Left            =   4320
            TabIndex        =   44
            Top             =   1920
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "4"
            Height          =   255
            Index           =   35
            Left            =   3360
            TabIndex        =   43
            Top             =   1920
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "3"
            Height          =   255
            Index           =   34
            Left            =   2280
            TabIndex        =   42
            Top             =   1920
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "2"
            Height          =   255
            Index           =   33
            Left            =   1320
            TabIndex        =   41
            Top             =   1920
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "1"
            Height          =   255
            Index           =   32
            Left            =   240
            TabIndex        =   40
            Top             =   1920
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "8"
            Height          =   255
            Index           =   31
            Left            =   7080
            TabIndex        =   39
            Top             =   1560
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "7"
            Height          =   255
            Index           =   30
            Left            =   6240
            TabIndex        =   38
            Top             =   1560
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "6"
            Height          =   255
            Index           =   29
            Left            =   5280
            TabIndex        =   37
            Top             =   1560
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "5"
            Height          =   255
            Index           =   28
            Left            =   4320
            TabIndex        =   36
            Top             =   1560
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "4"
            Height          =   255
            Index           =   27
            Left            =   3360
            TabIndex        =   35
            Top             =   1560
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "3"
            Height          =   255
            Index           =   26
            Left            =   2280
            TabIndex        =   34
            Top             =   1560
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "2"
            Height          =   255
            Index           =   25
            Left            =   1320
            TabIndex        =   33
            Top             =   1560
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "1"
            Height          =   255
            Index           =   24
            Left            =   240
            TabIndex        =   32
            Top             =   1560
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "8"
            Height          =   255
            Index           =   23
            Left            =   7080
            TabIndex        =   31
            Top             =   1200
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "7"
            Height          =   255
            Index           =   22
            Left            =   6240
            TabIndex        =   30
            Top             =   1200
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "6"
            Height          =   255
            Index           =   21
            Left            =   5280
            TabIndex        =   29
            Top             =   1200
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "5"
            Height          =   255
            Index           =   20
            Left            =   4320
            TabIndex        =   28
            Top             =   1200
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "4"
            Height          =   255
            Index           =   19
            Left            =   3360
            TabIndex        =   27
            Top             =   1200
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "3"
            Height          =   255
            Index           =   18
            Left            =   2280
            TabIndex        =   26
            Top             =   1200
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "2"
            Height          =   255
            Index           =   17
            Left            =   1320
            TabIndex        =   25
            Top             =   1200
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "1"
            Height          =   255
            Index           =   16
            Left            =   240
            TabIndex        =   24
            Top             =   1200
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "8"
            Height          =   255
            Index           =   15
            Left            =   7080
            TabIndex        =   23
            Top             =   840
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "7"
            Height          =   255
            Index           =   14
            Left            =   6240
            TabIndex        =   22
            Top             =   840
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "6"
            Height          =   255
            Index           =   13
            Left            =   5280
            TabIndex        =   21
            Top             =   840
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "5"
            Height          =   255
            Index           =   12
            Left            =   4320
            TabIndex        =   20
            Top             =   840
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "4"
            Height          =   255
            Index           =   11
            Left            =   3360
            TabIndex        =   19
            Top             =   840
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "3"
            Height          =   255
            Index           =   10
            Left            =   2280
            TabIndex        =   18
            Top             =   840
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "2"
            Height          =   255
            Index           =   9
            Left            =   1320
            TabIndex        =   17
            Top             =   840
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "1"
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   16
            Top             =   840
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "8"
            Height          =   255
            Index           =   7
            Left            =   7080
            TabIndex        =   15
            Top             =   480
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "7"
            Height          =   255
            Index           =   6
            Left            =   6240
            TabIndex        =   14
            Top             =   480
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "6"
            Height          =   255
            Index           =   5
            Left            =   5280
            TabIndex        =   13
            Top             =   480
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "5"
            Height          =   255
            Index           =   4
            Left            =   4320
            TabIndex        =   12
            Top             =   480
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "4"
            Height          =   255
            Index           =   3
            Left            =   3360
            TabIndex        =   11
            Top             =   480
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "3"
            Height          =   255
            Index           =   2
            Left            =   2280
            TabIndex        =   10
            Top             =   480
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "2"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   9
            Top             =   480
            Width           =   615
         End
         Begin VB.CheckBox yesno 
            Caption         =   "1"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   8
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Jumlah H:"
            Height          =   255
            Index           =   6
            Left            =   6960
            TabIndex        =   142
            Top             =   6000
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Jumlah G:"
            Height          =   255
            Index           =   5
            Left            =   6120
            TabIndex        =   140
            Top             =   6000
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Jumlah F:"
            Height          =   255
            Index           =   4
            Left            =   5160
            TabIndex        =   138
            Top             =   6000
            Width           =   735
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah E:"
            Height          =   195
            Index           =   3
            Left            =   4200
            TabIndex        =   136
            Top             =   6000
            Width           =   690
         End
         Begin VB.Label Label3 
            Caption         =   "Jumlah D:"
            Height          =   255
            Index           =   2
            Left            =   3240
            TabIndex        =   134
            Top             =   6000
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Jumlah C:"
            Height          =   255
            Index           =   1
            Left            =   2160
            TabIndex        =   132
            Top             =   6000
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "Jumlah B:"
            Height          =   255
            Left            =   1200
            TabIndex        =   130
            Top             =   6000
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Jumlah A:"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   128
            Top             =   6000
            Width           =   735
         End
         Begin VB.Line Line1 
            BorderWidth     =   3
            X1              =   120
            X2              =   7560
            Y1              =   5880
            Y2              =   5880
         End
      End
      Begin VB.Frame Frame1 
         Height          =   7815
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   9975
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3495
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Text            =   "frmtree.frx":0054
            Top             =   4080
            Width           =   9735
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3375
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Text            =   "frmtree.frx":005A
            Top             =   360
            Width           =   9735
         End
         Begin VB.Label Label2 
            Caption         =   "Penanganan"
            Height          =   375
            Left            =   120
            TabIndex        =   6
            Top             =   3840
            Width           =   2415
         End
         Begin VB.Label Label1 
            Caption         =   "Permasalahan"
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   120
            Width           =   2415
         End
      End
      Begin ActiveChart.XChart XChart1 
         Height          =   9675
         Left            =   -74880
         TabIndex        =   144
         Top             =   600
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   17066
         uTopMargin      =   600
         uBottomMargin   =   750
         uLeftMargin     =   750
         uRightMargin    =   750
         uContentBorder  =   -1  'True
         uSelectable     =   -1  'True
         uHotTracking    =   -1  'True
         uSelectedColumn =   -1
         uChartTitle     =   "Potensi Siswa"
         uChartSubTitle  =   "Italy"
         uAxisXOn        =   -1  'True
         uAxisYOn        =   -1  'True
         uColorBars      =   0   'False
         uIntersectMajor =   3
         uIntersectMinor =   1
         uMaxYValue      =   15
         uDisplayDescript=   -1  'True
         uXAxisLabel     =   "Klasifikasi Potensi Siswa"
         uYAxislabel     =   "Hasil Yang Diperoleh"
         BackColor       =   16777152
         ForeColor       =   0
         MinY            =   0
         BarColor        =   255
         SelectedBarColor=   16711680
         MajorGridColor  =   0
         MinorGridColor  =   0
         LegendBackColor =   16777088
         LegendForeColor =   0
         InfoBackColor   =   12648447
         InfoForeColor   =   16711680
         XAxisLabelColor =   0
         YAxisLabelColor =   0
         XAxisItemsColor =   0
         YAxisItemsColor =   0
         ChartTitleColor =   4210752
         ChartSubTitleColor=   0
         ChartType       =   4
         MenuType        =   0
         MenuItems       =   "&Save as...|&Print|&Copy|Selection &information|&Legend|&Properties|&Hide"
         InfoItems       =   ""
         SaveAsCaption   =   "Masifa"
         AutoRedraw      =   -1  'True
         BarWidthPercentage=   50
         BarSymbol       =   "*"
         BarPictureTile  =   -1  'True
         PictureTile     =   0   'False
         MinorGridOn     =   0   'False
         MajorGridOn     =   -1  'True
         LineWidth       =   1
         LineColor       =   16711680
         BarSymbolColor  =   255
         BarFillStyle    =   0
         LineStyle       =   0
         BarShadow       =   -1  'True
         BarShadowColor  =   0
         MeanOn          =   0   'False
         MeanColor       =   65535
         MeanCaption     =   ""
         DataFormat      =   "##.00"
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2520
      Top             =   7320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtree.frx":0060
            Key             =   "laporan"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtree.frx":68C2
            Key             =   "chart"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtree.frx":D124
            Key             =   "siswa"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtree.frx":13986
            Key             =   "data"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtree.frx":1A1E8
            Key             =   "tahun"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtree.frx":20A4A
            Key             =   "kelas"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   10821
      _Version        =   393217
      Indentation     =   998
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
End
Attribute VB_Name = "frmtree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer, b As Integer, c As Integer, d As Integer, e As Integer, f As Integer, g As Integer, h As Integer
Dim rs As New ADODB.Recordset, rsbim As New ADODB.Recordset, rsa As New ADODB.Recordset
Public Sub BuildTree(tvw As TreeView)
Dim nd As Node
Dim pkey As String, pkey2 As String, pkey3 As String, pkey4 As String
Dim key As String
Dim text As String
Dim nilai1 As Integer, nilai2 As Integer
Dim rsk As New ADODB.Recordset
sql = "select * from tbtahunajaran"
Set rsa = Nothing
rsa.CursorLocation = adUseClient
rsa.Open sql, koneksi
For o = 1 To rsa.RecordCount
    key = "t" & rsa!tahunajaran 'tahunajaran
    text = rsa!tahunajaran
    Set nd = tvw.Nodes.Add(, , key, text, "tahun")
    sql = "select * from tbkelas where idtahunajaran=" & rsa!IDTahunAjaran & ""
    Set rsk = Nothing
    rsk.CursorLocation = adUseClient
    rsk.Open sql, koneksi
    pkey3 = key
    For k = 1 To rsk.RecordCount
        key = "k" & rsk!IDKelas
        text = "Kelas" & " ( " & rsk!Kelas & "-" & rsk!Ruangan & " ) "
        Set nd = tvw.Nodes.Add(pkey3, tvwChild, key, text, "kelas")
        sql = "Select * from tbsiswa where IDKELAS=" & rsk!IDKelas & " order by nis asc"
        Set rs = Nothing
        rs.CursorLocation = adUseClient
        rs.Open sql, koneksi
        pkey4 = key
        For i = 1 To rs.RecordCount
            key = "a" & rs!NIS
            text = rs!NIS & " ( " & rs!Nama & " )"
            Set nd = tvw.Nodes.Add(pkey4, tvwChild, key, text, "siswa")
            pkey = key
            For a = 0 To 1
                If a = 0 Then
                    text = "Bimbingan"
                    key = "b" & rs!NIS
                    Set nd = tvw.Nodes.Add(pkey, tvwChild, key, text, "laporan")
                    sql = "SELECT IDKonseling,tgl,NIS,NIP,Permasalahan,Penanganan,Keterangan FROM TBKonseling where nis = " & rs!NIS & ""
                    Set rsbim = Nothing
                    rsbim.CursorLocation = adUseClient
                    rsbim.Open sql, koneksi
                    pkey2 = key
                    For b = 0 To rsbim.RecordCount - 1
                        nilai1 = b + 1
                        text = rsbim!tgl
                        key = "d" & rsbim!IDKonseling
                        Set nd = tvw.Nodes.Add(pkey2, tvwChild, key, text, "data")
                        rsbim.MoveNext
                    Next
                    rsbim.Close
                Else
                    text = "Potensi Siswa"
                    key = "c" & rs!NIS
                    Set nd = tvw.Nodes.Add(pkey, tvwChild, key, text, "chart")
                End If
            Next
            rs.MoveNext
        Next
        rsk.MoveNext
    Next
    rsa.MoveNext
Next
End Sub

Private Sub Command1_Click()
    Unload Me
    Me.Show
End Sub

Private Sub Form_Load()
    Call BuildTree(TreeView1)
    Call captioncek
    Text1.text = ""
    Text2.text = ""
End Sub
Private Sub Form_Resize()
On Error Resume Next
    Command1.Top = 840
    Command1.Left = 120
    Command1.Width = TreeView1.Width
    TreeView1.Top = 1080
    TreeView1.Left = 120
    TreeView1.Height = Me.Height - 1700
    SSTab1.Top = Command1.Top
    SSTab1.Height = TreeView1.Height + Command1.Height
    XChart1.Height = SSTab1.Height - 1000
    XChart1.Width = SSTab1.Width - 240
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
'On Error GoTo Hell
    If Left(Node.key, 1) = "d" Then
        cari_masalah (Mid(Node.key, 2, 100))
        SSTab1.Tab = 0
    ElseIf Left(Node.key, 1) = "c" Then
        XChart1.ChartSubTitle = Node.Parent.text
        SSTab1.Tab = 1
        Call isicap(Mid(Node.key, 2, 100))
        Call Cari_Nilai
        PrepareData
        XChart1.AutoRedraw = True
'        XChart1.ShowLegend (True)
        XChart1.Refresh
'        XChart1.ShowLegend (True)
    End If
'Hell:

End Sub
Private Function cari_masalah(pID As Long)
Dim Masalah As New ADODB.Recordset
    sql = "SELECT IDKonseling,tgl,NIS,NIP,Permasalahan,Penanganan,Keterangan FROM TBKonseling where IDKonseling = " & pID & ""
    Set Masalah = Nothing
    Masalah.CursorLocation = adUseClient
    Masalah.Open sql, koneksi
    
    If Not Masalah.EOF Then
        Text1.text = Masalah!Permasalahan
        Text2.text = Masalah!Penanganan
    End If

End Function

Private Sub captioncek()
    For i = 0 To 119
        yesno(i).Caption = i + 1
        yesno(i).Enabled = False
        yesno(i).Value = 0
    Next
    
End Sub

Private Sub isicap(pnis As Long)
    sql = "select * from Tbjawaban where NIS = " & pnis & " order by Nosoal"
    Set rsbim = Nothing
    rsbim.CursorLocation = adUseClient
    rsbim.Open sql, koneksi
    If Not rsbim.EOF Then
        Call captioncek
        For i = 0 To rsbim.RecordCount - 1
            If rsbim!jawaban = -1 Then
                yesno(i).Value = 1
                rsbim.MoveNext
            Else
                yesno(i).Value = 0
                rsbim.MoveNext
            End If
        Next
    Else
        For i = 0 To 119
            yesno(i).Value = 0
            
        Next
    End If
End Sub

Private Sub Cari_Nilai()
Dim AA As Integer
a = 0
b = 0
c = 0
d = 0
e = 0
f = 0
g = 0
h = 0
For i = 0 To 119
    'A
    If i = 0 Or i = 8 Or i = 16 Or i = 24 Or i = 32 Or _
        i = 40 Or i = 48 Or i = 56 Or i = 64 Or i = 72 Or i = 80 Or _
        i = 88 Or i = 96 Or i = 104 Or i = 112 Then
        If yesno(i).Value = 1 Then
            a = a + 1
        Else
            a = a
        End If
    'B
    ElseIf i = 1 Or i = 9 Or i = 17 Or i = 25 Or i = 33 Or _
        i = 41 Or i = 49 Or i = 57 Or i = 65 Or i = 73 Or i = 81 Or _
        i = 89 Or i = 97 Or i = 105 Or i = 113 Then
        If yesno(i).Value = 1 Then
            b = b + 1
        Else
            b = b
        End If
    'C
    ElseIf i = 2 Or i = 10 Or i = 18 Or i = 26 Or i = 34 Or _
        i = 42 Or i = 50 Or i = 58 Or i = 66 Or i = 74 Or i = 82 Or _
        i = 90 Or i = 98 Or i = 106 Or i = 114 Then
        If yesno(i).Value = 1 Then
            c = c + 1
        Else
            c = c
        End If
    'D
    ElseIf i = 3 Or i = 11 Or i = 19 Or i = 27 Or i = 35 Or _
        i = 43 Or i = 51 Or i = 59 Or i = 67 Or i = 75 Or i = 83 Or _
        i = 91 Or i = 99 Or i = 107 Or i = 115 Then
        If yesno(i).Value = 1 Then
            d = d + 1
        Else
            d = d
        End If
    'E
    ElseIf i = 4 Or i = 12 Or i = 20 Or i = 28 Or i = 36 Or _
        i = 44 Or i = 52 Or i = 60 Or i = 68 Or i = 76 Or i = 84 Or _
        i = 92 Or i = 100 Or i = 108 Or i = 116 Then
        If yesno(i).Value = 1 Then
            e = e + 1
        Else
            e = e
        End If
    'F
    ElseIf i = 5 Or i = 13 Or i = 21 Or i = 29 Or i = 37 Or _
        i = 45 Or i = 53 Or i = 61 Or i = 69 Or i = 77 Or i = 85 Or _
        i = 93 Or i = 101 Or i = 109 Or i = 117 Then
        If yesno(i).Value = 1 Then
            f = f + 1
        Else
            f = f
        End If
    'G
    ElseIf i = 6 Or i = 14 Or i = 22 Or i = 30 Or i = 38 Or _
        i = 46 Or i = 54 Or i = 62 Or i = 70 Or i = 78 Or i = 86 Or _
        i = 94 Or i = 102 Or i = 110 Or i = 118 Then
        If yesno(i).Value = 1 Then
            g = g + 1
        Else
            g = g
        End If
    'H
    ElseIf i = 7 Or i = 15 Or i = 23 Or i = 31 Or i = 39 Or _
        i = 47 Or i = 55 Or i = 63 Or i = 71 Or i = 79 Or i = 87 Or _
        i = 95 Or i = 103 Or i = 111 Or i = 119 Then
        If yesno(i).Value = 1 Then
            h = h + 1
        Else
            h = h
        End If
    End If
Next
    ta.text = a
    tb.text = b
    tc.text = c
    td.text = d
    te.text = e
    tf.text = f
    tg.text = g
    th.text = h
End Sub
Private Sub PrepareData()
    Dim X As Integer
    Dim intSign As Integer
    Dim oChartItem As ChartItem
    Dim varMonths As Variant
    Dim varMonthsExt As Variant
    varMonths = Array("A", "B", "C", "D", "E", "F", "G", "H")
    varMonthsExt = Array("Verbal-Linguistik", "Logis-Matematis", "Spasial", "Kinestetik", "Musikal", "Interpersonal", "Intrapersonal", "Natural")
    XChart1.AutoRedraw = True
    XChart1.chartType = xcBarLine
    XChart1.Clear
    For X = 1 To 8
        If X = 1 Then
            oChartItem.Value = a
        ElseIf X = 2 Then
            oChartItem.Value = b
        ElseIf X = 3 Then
            oChartItem.Value = c
        ElseIf X = 4 Then
            oChartItem.Value = d
        ElseIf X = 5 Then
            oChartItem.Value = e
        ElseIf X = 6 Then
            oChartItem.Value = f
        ElseIf X = 7 Then
            oChartItem.Value = g
        ElseIf X = 8 Then
            oChartItem.Value = h
        End If
        oChartItem.ItemID = X
        oChartItem.XAxisDescription = varMonths(X - 1)
        oChartItem.SelectedDescription = varMonthsExt(X - 1)
        XChart1.AddItem oChartItem
    Next X
End Sub
