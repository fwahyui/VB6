VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "MWARNET - SERVER"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11160
   Icon            =   "server.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "server.frx":FA8A
   ScaleHeight     =   7500
   ScaleWidth      =   11160
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   840
      TabIndex        =   0
      Top             =   1800
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   8705
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "MAIN"
      TabPicture(0)   =   "server.frx":18058
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Image1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label40"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame7"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "CmListCls"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame11"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txsecurity"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "TXDATE"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "TXTIME"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "CmChat"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Cmcaptured"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Txpassordhide"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Cmsaved"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "CmTimeronline"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "CmCalculator"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Cmprintbill"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Command2"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Command1"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Combo2"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Command3"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Check2"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).ControlCount=   23
      TabCaption(1)   =   "REPORT"
      TabPicture(1)   =   "server.frx":18074
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "Image14"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "NETWORK"
      TabPicture(2)   =   "server.frx":18090
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame15"
      Tab(2).Control(1)=   "Frame10"
      Tab(2).Control(2)=   "Frame4"
      Tab(2).Control(3)=   "Frame3"
      Tab(2).Control(4)=   "Image10"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "SETTING"
      TabPicture(3)   =   "server.frx":180AC
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Image13"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame9"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Frame12"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Frame14"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Frame13"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).ControlCount=   5
      Begin VB.Frame Frame15 
         Caption         =   "NOTE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   -68520
         TabIndex        =   185
         Top             =   3240
         Width           =   2775
         Begin VB.Label Label37 
            Caption         =   "Regards, Manik Artawan"
            Height          =   255
            Left            =   240
            TabIndex        =   187
            Top             =   1080
            Width           =   2295
         End
         Begin VB.Label Label32 
            Caption         =   "This tools just for test only, please don't use this tools to spy your customer !"
            Height          =   615
            Left            =   240
            TabIndex        =   186
            Top             =   360
            Width           =   2295
         End
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Emergency stop"
         Height          =   195
         Left            =   5400
         TabIndex        =   183
         Top             =   400
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "X"
         Height          =   255
         Left            =   7800
         TabIndex        =   182
         Top             =   380
         Width           =   255
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "server.frx":180C8
         Left            =   6960
         List            =   "server.frx":180DB
         TabIndex        =   181
         Text            =   "PC01"
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         Height          =   375
         Left            =   8280
         TabIndex        =   126
         ToolTipText     =   "Application exit"
         Top             =   4320
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Moved"
         Height          =   375
         Left            =   8280
         TabIndex        =   179
         ToolTipText     =   "Calculator"
         Top             =   3960
         Width           =   975
      End
      Begin VB.CommandButton Cmprintbill 
         Caption         =   "Note Print"
         Height          =   375
         Left            =   8280
         TabIndex        =   178
         ToolTipText     =   "Calculator"
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton CmCalculator 
         Caption         =   "Calculator"
         Height          =   375
         Left            =   8280
         TabIndex        =   125
         ToolTipText     =   "Calculator"
         Top             =   3240
         Width           =   975
      End
      Begin VB.CommandButton CmTimeronline 
         Caption         =   "Timer"
         Height          =   375
         Left            =   7200
         TabIndex        =   151
         ToolTipText     =   "Print receipt"
         Top             =   4320
         Width           =   975
      End
      Begin VB.CommandButton Cmsaved 
         Caption         =   "Save cost"
         Height          =   375
         Left            =   7200
         TabIndex        =   136
         ToolTipText     =   "Save every change"
         Top             =   3960
         Width           =   975
      End
      Begin VB.Frame Frame13 
         Caption         =   "CRYPTER"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   -69240
         TabIndex        =   164
         Top             =   2880
         Width           =   2055
         Begin VB.TextBox Text1 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   240
            Locked          =   -1  'True
            PasswordChar    =   "*"
            ScrollBars      =   2  'Vertical
            TabIndex        =   168
            Top             =   360
            Width           =   1575
         End
         Begin VB.CommandButton cmEncrypt 
            Caption         =   "Encrypt"
            Height          =   375
            Left            =   240
            TabIndex        =   167
            Top             =   1200
            Width           =   735
         End
         Begin VB.CommandButton cmDecrypt 
            Caption         =   "Decrypt"
            CausesValidation=   0   'False
            Height          =   375
            Left            =   1080
            TabIndex        =   166
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   240
            Locked          =   -1  'True
            PasswordChar    =   "*"
            ScrollBars      =   2  'Vertical
            TabIndex        =   165
            Top             =   720
            Width           =   1575
         End
      End
      Begin VB.TextBox Txpassordhide 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4320
         TabIndex        =   163
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Frame Frame14 
         Caption         =   "PASSWORD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   -71640
         TabIndex        =   156
         Top             =   2880
         Width           =   2175
         Begin VB.CommandButton Cmpassword 
            Caption         =   "Save"
            Height          =   375
            Left            =   720
            TabIndex        =   162
            Top             =   1200
            Width           =   1095
         End
         Begin VB.TextBox TXpassnew 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   720
            TabIndex        =   159
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox TXpassold 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   720
            TabIndex        =   158
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label25 
            Caption         =   "New :"
            Height          =   255
            Left            =   240
            TabIndex        =   161
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label24 
            Caption         =   "Old   :"
            Height          =   255
            Left            =   240
            TabIndex        =   160
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.CommandButton Cmcaptured 
         Caption         =   "Capture"
         Height          =   375
         Left            =   7200
         TabIndex        =   153
         ToolTipText     =   "Pictures viewer"
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton CmChat 
         Caption         =   "Chatting"
         Height          =   375
         Left            =   7200
         TabIndex        =   152
         ToolTipText     =   "Chatting room"
         Top             =   3240
         Width           =   975
      End
      Begin VB.TextBox TXTIME 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "H:mm:ss"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   4
         EndProperty
         Height          =   285
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   97
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox TXDATE 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "d/MM/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   96
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txsecurity 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   3600
         PasswordChar    =   "*"
         TabIndex        =   148
         Top             =   360
         Width           =   1095
      End
      Begin VB.Frame Frame11 
         Caption         =   "STATUS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   8160
         TabIndex        =   129
         Top             =   1440
         Width           =   1095
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            Caption         =   "On"
            Height          =   255
            Left            =   360
            TabIndex        =   132
            Top             =   960
            Width           =   375
         End
         Begin VB.Shape Shape5 
            FillColor       =   &H0000FF00&
            FillStyle       =   0  'Solid
            Height          =   135
            Left            =   120
            Top             =   960
            Width           =   135
         End
         Begin VB.Shape Shape4 
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   135
            Left            =   120
            Top             =   720
            Width           =   135
         End
         Begin VB.Shape Shape3 
            FillColor       =   &H0000FFFF&
            FillStyle       =   0  'Solid
            Height          =   135
            Left            =   120
            Top             =   480
            Width           =   135
         End
         Begin VB.Shape Shape2 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Left            =   120
            Top             =   240
            Width           =   135
         End
         Begin VB.Label Label28 
            Caption         =   "Off"
            Height          =   255
            Left            =   360
            TabIndex        =   133
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label30 
            Caption         =   "Ready"
            Height          =   255
            Left            =   360
            TabIndex        =   131
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label31 
            Caption         =   "Online"
            Height          =   255
            Left            =   360
            TabIndex        =   130
            Top             =   480
            Width           =   495
         End
      End
      Begin VB.CommandButton CmListCls 
         Caption         =   "C"
         Height          =   255
         Left            =   6800
         TabIndex        =   134
         ToolTipText     =   "Clear box"
         Top             =   3380
         Width           =   255
      End
      Begin VB.Frame Frame12 
         Caption         =   "USER INFO :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   -71640
         TabIndex        =   116
         Top             =   720
         Width           =   5895
         Begin VB.CommandButton Cminfoclear 
            Caption         =   "Clears"
            Height          =   375
            Left            =   4560
            TabIndex        =   128
            Top             =   1320
            Width           =   1095
         End
         Begin VB.CommandButton CmInfosave 
            Caption         =   "Save"
            Height          =   375
            Left            =   4560
            TabIndex        =   127
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox TIF4 
            Height          =   285
            Left            =   1560
            TabIndex        =   124
            Text            =   "-"
            Top             =   1440
            Width           =   2775
         End
         Begin VB.TextBox TIF3 
            Height          =   285
            Left            =   1560
            TabIndex        =   123
            Text            =   "-"
            Top             =   1080
            Width           =   2775
         End
         Begin VB.TextBox TIF2 
            Height          =   285
            Left            =   1560
            TabIndex        =   122
            Text            =   "-"
            Top             =   720
            Width           =   2775
         End
         Begin VB.TextBox TIF1 
            Height          =   285
            Left            =   1560
            TabIndex        =   121
            Text            =   "-"
            Top             =   360
            Width           =   2775
         End
         Begin VB.Label Label23 
            Caption         =   "E-mail                :"
            Height          =   255
            Left            =   240
            TabIndex        =   120
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label22 
            Caption         =   "Phone               :"
            Height          =   255
            Left            =   240
            TabIndex        =   119
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label21 
            Caption         =   "Address             :"
            Height          =   255
            Left            =   240
            TabIndex        =   118
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label16 
            Caption         =   "Company name :"
            Height          =   255
            Left            =   240
            TabIndex        =   117
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "REMOTE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   -68520
         TabIndex        =   110
         Top             =   720
         Width           =   2775
         Begin VB.OptionButton Option1 
            Caption         =   "None"
            Height          =   255
            Left            =   1560
            TabIndex        =   184
            Top             =   1080
            Width           =   855
         End
         Begin VB.OptionButton Oprmt7 
            Caption         =   "Chatroom"
            Height          =   255
            Left            =   240
            TabIndex        =   154
            Top             =   840
            Width           =   1095
         End
         Begin VB.OptionButton Oprmt6 
            Caption         =   "Stop Online"
            Height          =   255
            Left            =   240
            TabIndex        =   150
            Top             =   1080
            Width           =   1215
         End
         Begin VB.OptionButton Oprmt5 
            Caption         =   "None"
            Height          =   255
            Left            =   240
            TabIndex        =   149
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton Oprmt4 
            Caption         =   "Captured"
            Height          =   255
            Left            =   240
            TabIndex        =   135
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton Oprmt3 
            Caption         =   "Logoff"
            Height          =   255
            Left            =   1560
            TabIndex        =   115
            Top             =   600
            Width           =   855
         End
         Begin VB.OptionButton Oprmt2 
            Caption         =   "Reboot"
            Height          =   255
            Left            =   1560
            TabIndex        =   114
            Top             =   840
            Width           =   975
         End
         Begin VB.OptionButton Oprmt1 
            Caption         =   "Shutdown"
            Height          =   255
            Left            =   1560
            TabIndex        =   113
            Top             =   360
            Width           =   1095
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "server.frx":180FD
            Left            =   240
            List            =   "server.frx":18110
            TabIndex        =   112
            Text            =   "PC01"
            Top             =   1800
            Width           =   1095
         End
         Begin VB.CommandButton CmTurnoff 
            Caption         =   "Remote Proces"
            Height          =   615
            Left            =   1440
            TabIndex        =   111
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label Label19 
            Caption         =   "Clients number :"
            Height          =   255
            Left            =   240
            TabIndex        =   155
            Top             =   1560
            Width           =   1095
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "HISTORY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   3600
         TabIndex        =   94
         Top             =   3240
         Width           =   3495
         Begin VB.ListBox List2 
            Height          =   1035
            Left            =   120
            TabIndex        =   95
            Top             =   240
            Width           =   3255
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "COST"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   240
         TabIndex        =   60
         Top             =   3240
         Width           =   3255
         Begin VB.TextBox TXDISCOUNT 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   91
            Text            =   "0"
            Top             =   960
            Width           =   375
         End
         Begin VB.CommandButton CMDDiscmin 
            Caption         =   "<"
            Height          =   280
            Left            =   2520
            TabIndex        =   90
            Top             =   975
            Width           =   255
         End
         Begin VB.CommandButton CMDDiscmax 
            Caption         =   ">"
            Height          =   280
            Left            =   2760
            TabIndex        =   89
            Top             =   975
            Width           =   255
         End
         Begin VB.TextBox TXCOST 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   87
            Text            =   "0"
            Top             =   600
            Width           =   1335
         End
         Begin VB.CommandButton CMDratemax 
            Caption         =   ">"
            Height          =   280
            Left            =   2760
            TabIndex        =   86
            Top             =   600
            Width           =   255
         End
         Begin VB.CommandButton CMDratemin 
            Caption         =   "<"
            Height          =   280
            Left            =   2520
            TabIndex        =   85
            Top             =   600
            Width           =   255
         End
         Begin VB.TextBox Txststep 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   65
            Text            =   "0"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox Txstpmnt 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   63
            Text            =   "0"
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton CMDStpmax 
            Caption         =   ">"
            Height          =   280
            Left            =   2760
            TabIndex        =   62
            Top             =   240
            Width           =   255
         End
         Begin VB.CommandButton CMDStpmin 
            Caption         =   "<"
            Height          =   280
            Left            =   2520
            TabIndex        =   61
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label3 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Discount:"
            Height          =   255
            Left            =   120
            TabIndex        =   93
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label8 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "% / Hours"
            Height          =   255
            Left            =   1560
            TabIndex        =   92
            Top             =   960
            Width           =   795
         End
         Begin VB.Label Label1 
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "Rate/Hours :"
            Height          =   255
            Left            =   120
            TabIndex        =   88
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "Minute :"
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "NETSCAN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   -74760
         TabIndex        =   55
         Top             =   720
         Width           =   2895
         Begin VB.CommandButton CmPing 
            Caption         =   "Ping"
            Height          =   375
            Left            =   1800
            TabIndex        =   102
            Top             =   3480
            Width           =   855
         End
         Begin VB.TextBox TXPING 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   600
            MaxLength       =   24
            TabIndex        =   100
            Text            =   "192.168.1.1"
            Top             =   3480
            Width           =   1095
         End
         Begin VB.ListBox ListPing 
            Height          =   840
            Left            =   240
            TabIndex        =   99
            Top             =   2400
            Width           =   2415
         End
         Begin VB.CommandButton btnNetworscan 
            Caption         =   "Scan"
            Height          =   375
            Left            =   1800
            TabIndex        =   78
            ToolTipText     =   "Scan PC's Name on the Network"
            Top             =   1920
            Width           =   855
         End
         Begin VB.ListBox List1 
            Height          =   1185
            ItemData        =   "server.frx":18132
            Left            =   240
            List            =   "server.frx":18134
            Style           =   1  'Checkbox
            TabIndex        =   56
            Top             =   600
            Width           =   2415
         End
         Begin VB.Label LblPing 
            Caption         =   "I.P."
            Height          =   255
            Left            =   240
            TabIndex        =   101
            Top             =   3480
            Width           =   375
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            Caption         =   "View network IP and PC's name :"
            Height          =   255
            Left            =   240
            TabIndex        =   80
            Top             =   360
            Width           =   2415
         End
         Begin VB.Label Label7 
            Caption         =   "Presh scan button"
            Height          =   255
            Left            =   240
            TabIndex        =   57
            Top             =   1920
            Width           =   1575
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "CHANNEL LOG"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   -71640
         TabIndex        =   50
         Top             =   720
         Width           =   2895
         Begin VB.Timer Timer2 
            Interval        =   1000
            Left            =   360
            Top             =   960
         End
         Begin VB.TextBox CLtimer 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   137
            Text            =   "0"
            Top             =   3120
            Width           =   375
         End
         Begin VB.CommandButton bntListen 
            Caption         =   "Listen"
            Height          =   375
            Left            =   240
            TabIndex        =   79
            Tag             =   "Connect"
            Top             =   3480
            Width           =   1095
         End
         Begin MSWinsockLib.Winsock sock1 
            Index           =   0
            Left            =   360
            Top             =   480
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   393216
         End
         Begin VB.TextBox txtLog 
            Height          =   2655
            Left            =   240
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   53
            Top             =   360
            Width           =   2415
         End
         Begin VB.TextBox txtPort 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   720
            TabIndex        =   52
            Text            =   "123"
            Top             =   3120
            Width           =   615
         End
         Begin VB.CommandButton bntClear 
            Caption         =   "Clears"
            Height          =   375
            Left            =   1560
            TabIndex        =   51
            Top             =   3480
            Width           =   1095
         End
         Begin VB.Label Label27 
            Caption         =   "ClsTimer :"
            Height          =   255
            Left            =   1440
            TabIndex        =   138
            Top             =   3120
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "Port :"
            Height          =   255
            Left            =   240
            TabIndex        =   54
            Top             =   3120
            Width           =   495
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "DATABASE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4095
         Left            =   -74760
         TabIndex        =   48
         Top             =   600
         Width           =   9015
         Begin VB.CommandButton Cmprintreports 
            Caption         =   "Print Reports"
            Height          =   375
            Left            =   6480
            TabIndex        =   180
            ToolTipText     =   "Deleted selected data"
            Top             =   3120
            Width           =   1095
         End
         Begin VB.CommandButton Cmprintreport 
            Caption         =   "Print Bill"
            Height          =   375
            Left            =   6480
            TabIndex        =   177
            ToolTipText     =   "Deleted selected data"
            Top             =   3600
            Width           =   1095
         End
         Begin VB.TextBox hx8 
            DataField       =   "PAYMENT"
            DataSource      =   "Adodc1"
            Height          =   285
            Left            =   6960
            TabIndex        =   176
            Top             =   960
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox hx7 
            DataField       =   "DISCOUNT"
            DataSource      =   "Adodc1"
            Height          =   285
            Left            =   6000
            TabIndex        =   175
            Top             =   960
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox hx6 
            DataField       =   "COST"
            DataSource      =   "Adodc1"
            Height          =   285
            Left            =   5040
            TabIndex        =   174
            Top             =   960
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox hx5 
            DataField       =   "DURATION"
            DataSource      =   "Adodc1"
            Height          =   285
            Left            =   4080
            TabIndex        =   173
            Top             =   960
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox hx4 
            DataField       =   "TIMEOUT"
            DataSource      =   "Adodc1"
            Height          =   285
            Left            =   3240
            TabIndex        =   172
            Top             =   960
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox hx3 
            DataField       =   "TIMEIN"
            DataSource      =   "Adodc1"
            Height          =   285
            Left            =   2280
            TabIndex        =   171
            Top             =   960
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox hx2 
            DataField       =   "DATE"
            DataSource      =   "Adodc1"
            Height          =   285
            Left            =   1320
            TabIndex        =   170
            Top             =   960
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox hx1 
            DataField       =   "USER"
            DataSource      =   "Adodc1"
            Height          =   285
            Left            =   600
            TabIndex        =   169
            Top             =   960
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.CommandButton BTNReportDel 
            Caption         =   "Delete"
            Height          =   375
            Left            =   7680
            TabIndex        =   77
            ToolTipText     =   "Deleted selected data"
            Top             =   3600
            Width           =   1095
         End
         Begin VB.CommandButton BTNReportRefs 
            Caption         =   "View All"
            Height          =   375
            Left            =   7680
            TabIndex        =   76
            ToolTipText     =   "View all datas"
            Top             =   3120
            Width           =   1095
         End
         Begin VB.Frame Frame8 
            Caption         =   "SEARCH"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   2400
            TabIndex        =   71
            Top             =   3000
            Width           =   3855
            Begin VB.CheckBox Check1 
               Caption         =   "Today"
               Height          =   195
               Left            =   1800
               TabIndex        =   103
               Top             =   240
               Width           =   855
            End
            Begin VB.TextBox scrh2 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   600
               TabIndex        =   74
               Top             =   600
               Width           =   1095
            End
            Begin VB.TextBox scrh1 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   600
               TabIndex        =   72
               Top             =   240
               Width           =   1095
            End
            Begin VB.Image Image11 
               Height          =   480
               Left            =   3000
               Picture         =   "server.frx":18136
               Top             =   360
               Width           =   480
            End
            Begin VB.Label Label10 
               Caption         =   "User :"
               Height          =   255
               Left            =   120
               TabIndex        =   75
               Top             =   600
               Width           =   495
            End
            Begin VB.Label Label9 
               Caption         =   "Date :"
               Height          =   255
               Left            =   120
               TabIndex        =   73
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "INFORMATIONS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   240
            TabIndex        =   66
            Top             =   3000
            Width           =   2055
            Begin VB.TextBox Text3 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   68
               Top             =   480
               Width           =   855
            End
            Begin VB.TextBox Text4 
               Alignment       =   2  'Center
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#.##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   1
               EndProperty
               Height          =   285
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   67
               Top             =   480
               Width           =   855
            End
            Begin VB.Label Label12 
               Caption         =   "Data :"
               Height          =   255
               Left            =   120
               TabIndex        =   70
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label13 
               Caption         =   "Cost :"
               Height          =   255
               Left            =   1080
               TabIndex        =   69
               Top             =   240
               Width           =   615
            End
         End
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   330
            Left            =   480
            Top             =   2400
            Visible         =   0   'False
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   582
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   8
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data.mdb;Persist Security Info=False"
            OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data.mdb;Persist Security Info=False"
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "select * from TBL1"
            Caption         =   "Adodc1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "server.frx":18F78
            Height          =   2535
            Left            =   240
            TabIndex        =   49
            Top             =   360
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   4471
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
            AllowDelete     =   -1  'True
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
            ColumnCount     =   8
            BeginProperty Column00 
               DataField       =   "USER"
               Caption         =   "USER"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "DATE"
               Caption         =   "DATE"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "TIMEIN"
               Caption         =   "TIMEIN"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "TIMEOUT"
               Caption         =   "TIMEOUT"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "DURATION"
               Caption         =   "DURATION"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "COST"
               Caption         =   "COST"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "DISCOUNT"
               Caption         =   "DISCOUNT"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column07 
               DataField       =   "PAYMENT"
               Caption         =   "PAYMENT"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   780,095
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   959,811
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   900,284
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   870,236
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   989,858
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   900,284
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   1019,906
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   1035,213
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "COUNTER"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   7815
         Begin VB.TextBox ETX11 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6720
            Locked          =   -1  'True
            TabIndex        =   46
            Top             =   2040
            Width           =   495
         End
         Begin VB.TextBox DTX11 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6720
            Locked          =   -1  'True
            TabIndex        =   45
            Top             =   1680
            Width           =   495
         End
         Begin VB.TextBox CTX11 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6720
            Locked          =   -1  'True
            TabIndex        =   44
            Top             =   1320
            Width           =   495
         End
         Begin VB.TextBox BTX11 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6720
            Locked          =   -1  'True
            TabIndex        =   43
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox ATX11 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6720
            Locked          =   -1  'True
            TabIndex        =   42
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox ATX09 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   5880
            Locked          =   -1  'True
            TabIndex        =   41
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox ETX07 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   5160
            TabIndex        =   40
            Text            =   "Ready"
            Top             =   2040
            Width           =   735
         End
         Begin VB.TextBox DTX07 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   5160
            TabIndex        =   39
            Text            =   "Ready"
            Top             =   1680
            Width           =   735
         End
         Begin VB.TextBox CTX07 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   5160
            TabIndex        =   38
            Text            =   "Ready"
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox BTX07 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   5160
            TabIndex        =   37
            Text            =   "Ready"
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox ATX07 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   5160
            TabIndex        =   36
            Text            =   "Ready"
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox ETX06 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4320
            Locked          =   -1  'True
            TabIndex        =   35
            Text            =   "0"
            Top             =   2040
            Width           =   855
         End
         Begin VB.TextBox ETX05 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   34
            Text            =   "0"
            Top             =   2040
            Width           =   855
         End
         Begin VB.TextBox ETX04 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   33
            Text            =   "0"
            Top             =   2040
            Width           =   855
         End
         Begin VB.TextBox ETX03 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   32
            Text            =   "0"
            Top             =   2040
            Width           =   495
         End
         Begin VB.TextBox ETX02 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   2040
            Width           =   855
         End
         Begin VB.TextBox ETX01 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   30
            Text            =   "0"
            Top             =   2040
            Width           =   735
         End
         Begin VB.TextBox DTX06 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4320
            Locked          =   -1  'True
            TabIndex        =   29
            Text            =   "0"
            Top             =   1680
            Width           =   855
         End
         Begin VB.TextBox DTX05 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   28
            Text            =   "0"
            Top             =   1680
            Width           =   855
         End
         Begin VB.TextBox DTX04 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   27
            Text            =   "0"
            Top             =   1680
            Width           =   855
         End
         Begin VB.TextBox DTX03 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   26
            Text            =   "0"
            Top             =   1680
            Width           =   495
         End
         Begin VB.TextBox DTX02 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   1680
            Width           =   855
         End
         Begin VB.TextBox DTX01 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   24
            Text            =   "0"
            Top             =   1680
            Width           =   735
         End
         Begin VB.TextBox CTX06 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4320
            Locked          =   -1  'True
            TabIndex        =   23
            Text            =   "0"
            Top             =   1320
            Width           =   855
         End
         Begin VB.TextBox CTX05 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   22
            Text            =   "0"
            Top             =   1320
            Width           =   855
         End
         Begin VB.TextBox CTX04 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   21
            Text            =   "0"
            Top             =   1320
            Width           =   855
         End
         Begin VB.TextBox CTX03 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   20
            Text            =   "0"
            Top             =   1320
            Width           =   495
         End
         Begin VB.TextBox CTX02 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   1320
            Width           =   855
         End
         Begin VB.TextBox CTX01 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   18
            Text            =   "0"
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox BTX01 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   17
            Text            =   "0"
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox BTX02 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox BTX03 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   15
            Text            =   "0"
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox BTX04 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   14
            Text            =   "0"
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox BTX05 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   13
            Text            =   "0"
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox BTX06 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4320
            Locked          =   -1  'True
            TabIndex        =   12
            Text            =   "0"
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox ATX06 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4320
            Locked          =   -1  'True
            TabIndex        =   11
            Text            =   "0"
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox ATX05 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   10
            Text            =   "0"
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox ATX04 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   9
            Text            =   "0"
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox ATX03 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   8
            Text            =   "0"
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox ATX02 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox ATX01 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   6
            Text            =   "0"
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox BTX09 
            Height          =   285
            Left            =   5880
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox CTX09 
            Height          =   285
            Left            =   5880
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   1320
            Width           =   855
         End
         Begin VB.TextBox DTX09 
            Height          =   285
            Left            =   5880
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   1680
            Width           =   855
         End
         Begin VB.TextBox ETX09 
            Height          =   285
            Left            =   5880
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   2040
            Width           =   855
         End
         Begin VB.Shape shp5 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   255
            Left            =   7400
            Shape           =   3  'Circle
            Top             =   2040
            Width           =   255
         End
         Begin VB.Shape shp4 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   255
            Left            =   7400
            Shape           =   3  'Circle
            Top             =   1680
            Width           =   255
         End
         Begin VB.Shape shp3 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   255
            Left            =   7400
            Shape           =   3  'Circle
            Top             =   1320
            Width           =   255
         End
         Begin VB.Shape shp2 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   255
            Left            =   7400
            Shape           =   3  'Circle
            Top             =   960
            Width           =   255
         End
         Begin VB.Shape shp1 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   255
            Left            =   7400
            Shape           =   3  'Circle
            Top             =   600
            Width           =   255
         End
         Begin VB.Image Image9 
            Height          =   240
            Left            =   120
            Picture         =   "server.frx":18F8D
            Top             =   2040
            Width           =   135
         End
         Begin VB.Label ETX10 
            Caption         =   "05"
            Height          =   255
            Left            =   300
            TabIndex        =   84
            Top             =   2040
            Width           =   255
         End
         Begin VB.Image Image8 
            Height          =   240
            Left            =   120
            Picture         =   "server.frx":191A5
            Top             =   1680
            Width           =   135
         End
         Begin VB.Label DTX10 
            Caption         =   "04"
            Height          =   255
            Left            =   300
            TabIndex        =   83
            Top             =   1680
            Width           =   255
         End
         Begin VB.Image Image6 
            Height          =   240
            Left            =   120
            Picture         =   "server.frx":193BD
            Top             =   1320
            Width           =   135
         End
         Begin VB.Label CTX10 
            Caption         =   "03"
            Height          =   255
            Left            =   300
            TabIndex        =   82
            Top             =   1320
            Width           =   255
         End
         Begin VB.Image Image5 
            Height          =   240
            Left            =   120
            Picture         =   "server.frx":195D5
            Top             =   960
            Width           =   135
         End
         Begin VB.Label BTX10 
            Caption         =   "02"
            Height          =   255
            Left            =   300
            TabIndex        =   81
            Top             =   960
            Width           =   255
         End
         Begin VB.Image Image4 
            Height          =   240
            Left            =   120
            Picture         =   "server.frx":197ED
            Top             =   600
            Width           =   135
         End
         Begin VB.Label Label2 
            Caption         =   "  Counter    Duration    Step        Cost         Discount      Total         Type       Timein       Index  Status"
            Height          =   255
            Left            =   600
            TabIndex        =   58
            Top             =   240
            Width           =   7095
         End
         Begin VB.Label ATX10 
            Caption         =   "01"
            Height          =   255
            Left            =   300
            TabIndex        =   47
            Top             =   600
            Width           =   255
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H80000014&
         BorderStyle     =   0  'None
         Caption         =   "ABOUT IT"
         Height          =   4215
         Left            =   -74760
         TabIndex        =   104
         Top             =   480
         Width           =   2775
         Begin VB.Image Image12 
            Height          =   750
            Left            =   1800
            Picture         =   "server.frx":19A05
            Top             =   3000
            Width           =   750
         End
         Begin VB.Image Image16 
            Height          =   390
            Left            =   240
            Picture         =   "server.frx":1D029
            Top             =   2640
            Width           =   390
         End
         Begin VB.Label Label38 
            BackStyle       =   0  'Transparent
            Caption         =   "gungmanik@telkom.net"
            Height          =   255
            Left            =   720
            TabIndex        =   146
            Top             =   2640
            Width           =   1695
         End
         Begin VB.Image Image3 
            Height          =   4215
            Left            =   0
            Top             =   0
            Width           =   2775
         End
         Begin VB.Image Image2 
            Height          =   525
            Left            =   360
            Picture         =   "server.frx":1D1CA
            Top             =   3600
            Width           =   2085
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "MWARNET"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   109
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "Cyber Cafe Sollutions"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   108
            Top             =   600
            Width           =   2535
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "Version 2.6.2"
            Height          =   255
            Left            =   120
            TabIndex        =   107
            Top             =   840
            Width           =   2535
         End
         Begin VB.Label Label17 
            BackColor       =   &H80000014&
            BackStyle       =   0  'Transparent
            Caption         =   "Programed and Design by"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   106
            Top             =   2160
            Width           =   2295
         End
         Begin VB.Label Label18 
            BackColor       =   &H80000014&
            BackStyle       =   0  'Transparent
            Caption         =   "A.A.Ngr.Manik Artawan, ST,MT"
            Height          =   255
            Left            =   240
            TabIndex        =   105
            Top             =   2400
            Width           =   2415
         End
         Begin VB.Image Image7 
            Height          =   705
            Left            =   720
            Picture         =   "server.frx":1D875
            Top             =   1320
            Width           =   1380
         End
      End
      Begin VB.Label Label40 
         Caption         =   "Administrator :"
         Height          =   255
         Left            =   2400
         TabIndex        =   147
         Top             =   390
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "server.frx":1E1C1
         Top             =   120
         Width           =   480
      End
      Begin VB.Image Image14 
         Height          =   510
         Left            =   -72480
         Picture         =   "server.frx":1EC7B
         Top             =   100
         Width           =   510
      End
      Begin VB.Image Image13 
         Height          =   450
         Left            =   -67680
         Picture         =   "server.frx":1F1D8
         Top             =   120
         Width           =   465
      End
      Begin VB.Image Image10 
         Height          =   480
         Left            =   -70080
         Picture         =   "server.frx":1F884
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "TODAY :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8280
         TabIndex        =   98
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   8640
      Top             =   960
   End
   Begin VB.Label Label39 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      Caption         =   "Version 2.6.2"
      Height          =   255
      Left            =   840
      TabIndex        =   188
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Serial no: 001/MT/2007/5Cln/Sadewa"
      Height          =   255
      Left            =   840
      TabIndex        =   157
      Top             =   6840
      Width           =   3255
   End
   Begin VB.Label Label36 
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail    :"
      Height          =   255
      Left            =   5160
      TabIndex        =   145
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label35 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone   :"
      Height          =   255
      Left            =   5160
      TabIndex        =   144
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "Address :"
      Height          =   255
      Left            =   5160
      TabIndex        =   143
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   "Name    :"
      Height          =   255
      Left            =   5160
      TabIndex        =   142
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Labelwarnet3 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5880
      TabIndex        =   141
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Labelwarnet2 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Phone"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5880
      TabIndex        =   140
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Labelwarnet1 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5880
      TabIndex        =   139
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label Labelwarnet 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5880
      TabIndex        =   59
      Top             =   480
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MWARNET 2 - FREEWARE EDITION
'COPYRIGHT(C) 2007 MTechnologi Bali Indonesia
'Programed by A.A.Ngr.Manik Artawan
'e-mail : gungmanik@telkom.net
'---------------------------------------------
'THANK YOU FOR DOWNLOAD THIS SMALL APPLICATION
'---------------------------------------------

Dim SocketCounter As Long
Private Function SecondsToTime(ByVal dSeconds As Double) As String
    SecondsToTime = Format(DateAdd("s", dSeconds, "00:00:00"), "HH:mm:ss")
End Function

Private Sub bntClear_Click()
txtLog.Text = ""
End Sub

Private Sub bntListen_Click()
On Error Resume Next
For n = 1 To SocketCounter
    sock1(n).Close
    Unload sock1(n)
Next
On Error GoTo t
sock1(0).Close
sock1(0).LocalPort = txtPort
sock1(0).Listen
txtLog = "Listening on Port " & txtPort
Exit Sub
t:
MsgBox "Error : " & Err.Description, vbCritical
End Sub

Private Sub btnNetworscan_Click()
Screen.MousePointer = vbHourglass
List1.AddItem "Scanning... Please wait..."
Dim l As New LAN
Dim s() As String
s = Split(l.GetPCList, "||")
List1.Clear
For i = LBound(s) To UBound(s)
List1.AddItem s(i)
Next
Label7.Caption = "Found " & UBound(s) & " PCs"
 Screen.MousePointer = vbNormal
End Sub

Private Sub BTNReportDel_Click()
If txsecurity.Text = Txpassordhide.Text Then
With Adodc1.Recordset
    If .RecordCount = 0 Then Exit Sub
    .Delete
    .Update
    End With
End If
End Sub

Private Sub BTNReportRefs_Click()
scrh1.Text = ""
costtotal
End Sub

Private Sub costtotal()
On Error GoTo ER1
Text4.Text = 0
Adodc1.Refresh
With Adodc1.Recordset
    If .RecordCount = 0 Then Exit Sub
    If .RecordCount > 0 Then .MoveFirst
    Do While Not .EOF
        Text4.Text = Val(Text4.Text) + !PAYMENT
        .MoveNext
    Loop
End With
Text4.Text = Format(Text4, "#,###")
Text3.Text = Adodc1.Recordset.RecordCount
ER1:
Exit Sub
End Sub

Private Sub Check1_Click()
If Check1.Value = Checked Then
    scrh1.Text = Date
    Else
    BTNReportRefs_Click
    End If
End Sub


Private Sub CmCalculator_Click()
Form2.Show
End Sub

Private Sub Cmcaptured_Click()
Form3.Show
End Sub

Private Sub CmChat_Click()
Form4.Show
End Sub

Private Sub CMDDiscmin_Click()
TXDISCOUNT.Text = Val(TXDISCOUNT - 5)
If TXDISCOUNT.Text < 0 Then TXDISCOUNT.Text = 0
End Sub

Private Sub CMDDiscmax_Click()
TXDISCOUNT.Text = Val(TXDISCOUNT + 5)
If TXDISCOUNT.Text > 100 Then TXDISCOUNT.Text = 100
End Sub

Private Sub cmDecrypt_Click()
Text2.Text = decrypted(3, Text1.Text)
If Text2.Text <> "" Then
Text1.Text = Text2.Text
End If
End Sub

Private Sub CMDratemax_Click()
TXCOST.Text = Val(TXCOST + 100)
BAGI = Txstpmnt.Text / 60
Txststep.Text = Val(TXCOST.Text) * BAGI
End Sub

Private Sub CMDratemin_Click()
TXCOST.Text = Val(TXCOST - 100)
If TXCOST.Text < 0 Then TXCOST.Text = 0
BAGI = Txstpmnt.Text / 60
Txststep.Text = Val(TXCOST.Text) * BAGI
End Sub

Private Sub CMDStpmax_Click()
Txstpmnt.Text = Val(Txstpmnt + 1)
If Txstpmnt.Text > 60 Then Txstpmnt.Text = 60
BAGI = Txstpmnt.Text / 60
Txststep.Text = Val(TXCOST.Text) * BAGI
End Sub

Private Sub CMDStpmin_Click()
Txstpmnt.Text = Val(Txstpmnt - 1)
If Txstpmnt.Text < 0 Then Txstpmnt.Text = 0
BAGI = Txstpmnt.Text / 60
Txststep.Text = Val(TXCOST.Text) * BAGI
End Sub

Private Sub cmEncrypt_Click()
Text2.Text = encrypted(3, Text1.Text)
If Text2.Text <> "" Then
Text1.Text = Text2.Text
End If
End Sub

Private Sub Cminfoclear_Click()
TIF1.Text = "-": TIF2.Text = "-": TIF3.Text = "-": TIF4.Text = "-"
End Sub

Private Sub CmInfosave_Click()
If txsecurity.Text = Txpassordhide.Text Then
    USERDATASAVE
    Else
    USERDATA
    End If
End Sub

Private Sub CmListCls_Click()
List2.Clear
End Sub

Private Sub Cmpassword_Click()
If TXpassold.Text <> Txpassordhide.Text Then
    TXpassold.Text = "Bad password"
    Exit Sub
    End If
    
Dim intFileHandle As Integer
    intFileHandle = FreeFile
    Open App.Path + "\mwarnet.pwd" For Output As #intFileHandle
    Text1.Text = TXpassnew.Text
    cmEncrypt_Click
    Write #intFileHandle, Text1.Text
    Close #intFileHandle
    
Txpassordhide.Text = TXpassnew.Text
TXpassnew.Text = "": TXpassold.Text = "Successfully"
End Sub

Private Sub CmPing_Click()
ListPing.Clear
Dim ECHO As ICMP_ECHO_REPLY
Dim pos As Integer
Call Ping(TXPING.Text, ECHO)
With ListPing
.AddItem GetStatusCode(ECHO.status)
.AddItem ECHO.Address
.AddItem ECHO.RoundTripTime & " ms"
.AddItem ECHO.DataSize & " bytes"
If Left$(ECHO.data, 1) <> Chr$(0) Then
    pos = InStr(ECHO.data, Chr$(0))
    .AddItem Left$(ECHO.data, pos - 1)
    End If
.AddItem ECHO.DataPointer
End With
End Sub



Private Sub Cmprintbill_Click()
SSTab1.Tab = 1
End Sub

Private Sub Cmprintreport_Click()
If hx1.Text = "" Then MsgBox "Please select database to print!", vbExclamation: Exit Sub
Form6.Show
End Sub


Private Sub Cmsaved_Click()
If txsecurity.Text = Txpassordhide.Text Then
    SAVERATE
    Else
    SERVERDATA
    End If
End Sub

Private Sub CmTimeronline_Click()
Form5.Show
End Sub

Private Sub CmTurnoff_Click()
    If Combo1.Text = "PC01" Then indexsock = ATX11.Text: ATX11.Text = ""
    If Combo1.Text = "PC02" Then indexsock = BTX11.Text: BTX11.Text = ""
    If Combo1.Text = "PC03" Then indexsock = CTX11.Text: CTX11.Text = ""
    If Combo1.Text = "PC04" Then indexsock = DTX11.Text: DTX11.Text = ""
    If Combo1.Text = "PC05" Then indexsock = ETX11.Text: ETX11.Text = ""
    If indexsock = "" Then
        Else
        If Oprmt1.Value = True Then sock1(indexsock).SendData "SHUTDOWN"
        If Oprmt2.Value = True Then sock1(indexsock).SendData "REBOOT"
        If Oprmt3.Value = True Then sock1(indexsock).SendData "LOGOFF"
        If Oprmt4.Value = True Then sock1(indexsock).SendData "CAPTURED"
        If Oprmt5.Value = True Then Exit Sub
        If Oprmt6.Value = True Then sock1(indexsock).SendData "STOP"
        If Oprmt7.Value = True Then Form4.Show: sock1(indexsock).SendData "CHAT"
        End If
End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
Form7.Show
End Sub

Private Sub Command3_Click()
If Check2.Value = Checked Then
    If Combo2.Text = "PC01" Then ATX07.Text = "STOP"
    If Combo2.Text = "PC02" Then BTX07.Text = "STOP"
    If Combo2.Text = "PC03" Then CTX07.Text = "STOP"
    If Combo2.Text = "PC04" Then DTX07.Text = "STOP"
    If Combo2.Text = "PC05" Then ETX07.Text = "STOP"
    Check2.Value = Unchecked
    End If
End Sub



Private Sub sock1_Close(Index As Integer)
sock1(Index).Close
Unload sock1(Index)
txtLog = txtLog & "Client" & Index & " -> *** Disconnected" & vbCrLf
SAVERATE
End Sub

Private Sub sock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
SocketCounter = SocketCounter + 1
Load sock1(SocketCounter)
sock1(SocketCounter).Accept requestID
txtLog = "Client Connected. IP : " & sock1(0).RemoteHostIP & " , Client Nick : Client" & sockcounter & vbCrLf
sock1(SocketCounter).SendData "Your Nick is ""Client" & SocketCounter & """"
End Sub

Private Sub sock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim dat As String
sock1(Index).GetData dat, vbString
txtLog = txtLog & "Client" & Index & " : " & dat & vbCrLf
On Error Resume Next

If dat = "PC01-ON" Then shp1.FillColor = vbGreen: ATX11.Text = Index
If dat = "PC02-ON" Then shp2.FillColor = vbGreen: BTX11.Text = Index
If dat = "PC03-ON" Then shp3.FillColor = vbGreen: CTX11.Text = Index
If dat = "PC04-ON" Then shp4.FillColor = vbGreen: DTX11.Text = Index
If dat = "PC05-ON" Then shp5.FillColor = vbGreen: ETX11.Text = Index

If dat = "PC01" Then
    sock1(Index).SendData "Duration: " & ATX02.Text & "  Cost: " & ATX06.Text
    ATX07.Text = "ONLINE"
    ATX11.Text = Index
    shp1.FillColor = vbYellow
    End If
If dat = "PC01-STOP" Then
    sock1(Index).SendData "STOP"
    ATX07.Text = "STOP"
    shp1.FillColor = vbWhite
    End If
If dat = "PC02" Then
    sock1(Index).SendData "Duration: " & BTX02.Text & "  Cost: " & BTX06.Text
    BTX07.Text = "ONLINE"
    BTX11.Text = Index
    shp2.FillColor = vbYellow
    End If
If dat = "PC02-STOP" Then
    sock1(Index).SendData "STOP"
    BTX07.Text = "STOP"
    shp2.FillColor = vbWhite
    End If
 If dat = "PC03" Then
    sock1(Index).SendData "Duration: " & CTX02.Text & "  Cost: " & CTX06.Text
    CTX07.Text = "ONLINE"
    CTX11.Text = Index
    shp3.FillColor = vbYellow
    End If
If dat = "PC03-STOP" Then
    sock1(Index).SendData "STOP"
    CTX07.Text = "STOP"
    shp3.FillColor = vbWhite
    End If
If dat = "PC04" Then
    sock1(Index).SendData "Duration: " & DTX02.Text & "  Cost: " & DTX06.Text
    DTX07.Text = "ONLINE"
    DTX11.Text = Index
    shp4.FillColor = vbYellow
    End If
If dat = "PC04-STOP" Then
    sock1(Index).SendData "STOP"
    DTX07.Text = "STOP"
    shp4.FillColor = vbWhite
    End If
If dat = "PC05" Then
    sock1(Index).SendData "Duration: " & ETX02.Text & "  Cost: " & ETX06.Text
    ETX07.Text = "ONLINE"
    ETX11.Text = Index
    shp5.FillColor = vbYellow
    End If
If dat = "PC05-STOP" Then
    sock1(Index).SendData "STOP"
    ETX07.Text = "STOP"
    shp5.FillColor = vbWhite
    End If
    
End Sub

Private Sub sock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
txtLog = txtLog & "*** Error ( Client" & Index & ") : " & Description & vbCrLf
sock1_Close Index
End Sub

Private Sub Form_Load()
If App.PrevInstance Then Unload Me: End
DisableCloseWindowButton Me
SERVERDATA
RATECHANGE
READVALUES
USERDATA
PASSVALUES
bntListen_Click
scrh1.Text = Date
Form4.Show
End Sub

Public Sub WRITEVALUES()
 Open App.Path + "\tempdata.txt" For Output As #1
 Write #1, ATX01, ATX02, ATX03, ATX04, ATX05, ATX06, ATX07, ATX08, ATX09, ATX11
 Write #1, BTX01, BTX02, BTX03, BTX04, BTX05, BTX06, BTX07, BTX08, BTX09, BTX11
 Write #1, CTX01, CTX02, CTX03, CTX04, CTX05, CTX06, CTX07, CTX08, CTX09, CTX11
 Write #1, DTX01, DTX02, DTX03, DTX04, DTX05, DTX06, DTX07, DTX08, DTX09, DTX11
 Write #1, ETX01, ETX02, ETX03, ETX04, ETX05, ETX06, ETX07, ETX08, ETX09, ETX11
 Close #1
End Sub

Public Sub READVALUES()
On Error GoTo t:
Dim intFileHandle As Integer
intFileHandle = FreeFile
Open App.Path + "\tempdata.txt" For Input As #1
Input #intFileHandle, a, B, c, d, E, F, G, H, i, J
ATX01 = a: ATX02 = B: ATX03 = c: ATX04 = d: ATX05 = E: ATX06 = F: ATX07 = G: ATX08 = H: ATX09 = i: ATX11 = J
Input #intFileHandle, a, B, c, d, E, F, G, H, i, J
BTX01 = a: BTX02 = B: BTX03 = c: BTX04 = d: BTX05 = E: BTX06 = F: BTX07 = G: BTX08 = H: BTX09 = i: BTX11 = J
Input #intFileHandle, a, B, c, d, E, F, G, H, i, J
CTX01 = a: CTX02 = B: CTX03 = c: CTX04 = d: CTX05 = E: CTX06 = F: CTX07 = G: CTX08 = H: CTX09 = i: CTX11 = J
Input #intFileHandle, a, B, c, d, E, F, G, H, i, J
DTX01 = a: DTX02 = B: DTX03 = c: DTX04 = d: DTX05 = E: DTX06 = F: DTX07 = G: DTX08 = H: DTX09 = i: DTX11 = J
Input #intFileHandle, a, B, c, d, E, F, G, H, i, J
ETX01 = a: ETX02 = B: ETX03 = c: ETX04 = d: ETX05 = E: ETX06 = F: ETX07 = G: ETX08 = H: ETX09 = i: ETX11 = J
Close #intFileHandle
Exit Sub
t:
Close #intFileHandle
Open App.Path + "\tempdata.txt" For Output As #1
Write #1, "0", "", "0", "1500", "0", "1500", "Ready", , "", ""
Write #1, "0", "", "0", "1500", "0", "1500", "Ready", , "", ""
Write #1, "0", "", "0", "1500", "0", "1500", "Ready", , "", ""
Write #1, "0", "", "0", "1500", "0", "1500", "Ready", , "", ""
Write #1, "0", "", "0", "1500", "0", "1500", "Ready", , "", ""
Close #1
End Sub

Public Sub PASSVALUES()
Dim intFileHandle As Integer
intFileHandle = FreeFile
Open App.Path + "\mwarnet.pwd" For Input As #1
Input #intFileHandle, pssd
    Text1.Text = pssd
    cmDecrypt_Click
    Txpassordhide.Text = Text1.Text
Close #intFileHandle
End Sub

Private Sub RATECHANGE()
ATX04.Text = Txststep.Text: ATX06.Text = ATX04.Text
BTX04.Text = Txststep.Text: BTX06.Text = BTX04.Text
CTX04.Text = Txststep.Text: CTX06.Text = CTX04.Text
DTX04.Text = Txststep.Text: DTX06.Text = DTX04.Text
ETX04.Text = Txststep.Text: ETX06.Text = ETX04.Text
End Sub

Private Sub SAVERATE()
    Dim intFileHandle As Integer
    intFileHandle = FreeFile
    Open App.Path + "\serversdata.txt" For Output As #intFileHandle
    Print #intFileHandle, TXCOST.Text
    Print #intFileHandle, TXDISCOUNT.Text
    Print #intFileHandle, Txstpmnt.Text
    Print #intFileHandle, Txststep.Text
    Print #intFileHandle, Labelwarnet.Caption
    Close #intFileHandle
End Sub

Private Sub SERVERDATA()
Dim intFileHandle As Integer
Dim strRETP As String
intFileHandle = FreeFile
Open App.Path + "\serversdata.txt" For Input As #intFileHandle
Line Input #intFileHandle, strRETP: TXCOST.Text = strRETP
Line Input #intFileHandle, strRETP: TXDISCOUNT.Text = strRETP
Line Input #intFileHandle, strRETP: Txstpmnt.Text = strRETP
Line Input #intFileHandle, strRETP: Txststep.Text = strRETP
Close #intFileHandle
End Sub

Private Sub USERDATA()
Dim intFileHandle As Integer
Dim strRETP As String
intFileHandle = FreeFile
Open App.Path + "\user.txt" For Input As #intFileHandle
Line Input #intFileHandle, strRETP: TIF1.Text = strRETP
Line Input #intFileHandle, strRETP: TIF2.Text = strRETP
Line Input #intFileHandle, strRETP: TIF3.Text = strRETP
Line Input #intFileHandle, strRETP: TIF4.Text = strRETP
Close #intFileHandle
Labelwarnet.Caption = TIF1.Text
Labelwarnet1.Caption = TIF2.Text
Labelwarnet2.Caption = TIF3.Text
Labelwarnet3.Caption = TIF4.Text
End Sub

Private Sub USERDATASAVE()
    Dim intFileHandle As Integer
    intFileHandle = FreeFile
    Open App.Path + "\user.txt" For Output As #intFileHandle
    Print #intFileHandle, TIF1.Text
    Print #intFileHandle, TIF2.Text
    Print #intFileHandle, TIF3.Text
    Print #intFileHandle, TIF4.Text
    Close #intFileHandle
End Sub


Private Sub scrh1_Change()
Select Case Len(scrh1)
Case 2
scrh1.SelText = "/"
End Select
With Adodc1
.RecordSource = "select * from TBL1 where date like '%" & _
scrh1.Text & "%'"
.Refresh
End With
costtotal
End Sub

Private Sub scrh2_Change()
With Adodc1
.RecordSource = "select * from TBL1 where USER like '%" & _
scrh2.Text & "%'"
.Refresh
End With
costtotal
End Sub


Private Sub TIF1_Change()
Labelwarnet.Caption = TIF1.Text
End Sub

Private Sub Timer1_Timer()
TXTIME.Text = Time: TXDATE.Text = Date
COUNTER01
COLORTEX
End Sub

Private Sub COLORTEX()
If ATX07.Text = "ONLINE" Then ATX07.BackColor = vbYellow Else ATX07.BackColor = vbWhite
If BTX07.Text = "ONLINE" Then BTX07.BackColor = vbYellow Else BTX07.BackColor = vbWhite
If CTX07.Text = "ONLINE" Then CTX07.BackColor = vbYellow Else CTX07.BackColor = vbWhite
If DTX07.Text = "ONLINE" Then DTX07.BackColor = vbYellow Else DTX07.BackColor = vbWhite
If ETX07.Text = "ONLINE" Then ETX07.BackColor = vbYellow Else ETX07.BackColor = vbWhite
End Sub

Private Sub COUNTER01()
LISTC = List2.ListCount
If ATX01.Text = "1" Then ATX09.Text = Time
If ATX07.Text = "MOVE" Then
    ATX01.Text = "0": ATX02.Text = "": ATX03.Text = "0": ATX04.Text = Txststep.Text
    ATX05.Text = "0": ATX06.Text = Val(ATX04.Text) - Val(ATX05.Text)
    ATX07.Text = "Ready"
    ATX09.Text = ""
    ATX11.Text = ""
    End If
If ATX07.Text = "STOP" Then
    List2.AddItem "PC-" & ATX10.Caption & " Cost: " & ATX04.Text & " Duration: " & ATX02.Text
    List2.Selected(LISTC) = True
    With Adodc1.Recordset
        .AddNew
            !USER = ATX10.Caption
            !Date = TXDATE.Text
            !TIMEIN = ATX09.Text
            !Timeout = TXTIME.Text
            !DURATION = ATX02.Text
            !COST = ATX04.Text
            !DISCOUNT = ATX05.Text
            !PAYMENT = ATX06.Text
        .Update
        .MoveLast
    End With
    ATX01.Text = "0": ATX02.Text = "": ATX03.Text = "0": ATX04.Text = Txststep.Text
    ATX05.Text = "0": ATX06.Text = Val(ATX04.Text) - Val(ATX05.Text)
    ATX07.Text = "Ready"
    ATX09.Text = ""
    ATX11.Text = ""
    WRITEVALUES
End If
If ATX07.Text = "ONLINE" Then
    WRITEVALUES
    ATX01.Text = Val(ATX01.Text + 1)
    ATX02.Text = SecondsToTime(ATX01.Text)
    ATX03.Text = Val(ATX03 + 1)
    If ATX03.Text = Val(Txstpmnt.Text) * 60 Then
        ATX03.Text = 0
        ATX04.Text = Val(ATX04.Text) + Val(Txststep.Text)
        ATX05.Text = Val(ATX04.Text) * Val(TXDISCOUNT.Text) / 100
        ATX06.Text = Val(ATX04.Text) - Val(ATX05.Text)
        End If
End If

If BTX01.Text = "1" Then BTX09.Text = Time
If BTX07.Text = "MOVE" Then
    BTX01.Text = "0": BTX02.Text = "": BTX03.Text = "0": BTX04.Text = Txststep.Text
    BTX05.Text = "0": BTX06.Text = Val(BTX04.Text) - Val(BTX05.Text)
    BTX07.Text = "Ready"
    BTX09.Text = ""
    BTX11.Text = ""
    End If
If BTX07.Text = "STOP" Then
    List2.AddItem "PC-" & BTX10.Caption & " Cost: " & BTX04.Text & " Duration: " & BTX02.Text
    List2.Selected(LISTC) = True
    With Adodc1.Recordset
        .AddNew
            !USER = BTX10.Caption
            !Date = TXDATE.Text
            !TIMEIN = BTX09.Text
            !Timeout = TXTIME.Text
            !DURATION = BTX02.Text
            !COST = BTX04.Text
            !DISCOUNT = BTX05.Text
            !PAYMENT = BTX06.Text
        .Update
        .MoveLast
    End With
    BTX01.Text = "0": BTX02.Text = "": BTX03.Text = "0": BTX04.Text = Txststep.Text
    BTX05.Text = "0": BTX06.Text = Val(BTX04.Text) - Val(BTX05.Text)
    BTX07.Text = "Ready"
    BTX09.Text = ""
    BTX11.Text = ""
    WRITEVALUES
End If
If BTX07.Text = "ONLINE" Then
    WRITEVALUES
    BTX01.Text = Val(BTX01.Text + 1)
    BTX02.Text = SecondsToTime(BTX01.Text)
    BTX03.Text = Val(BTX03 + 1)
    If BTX03.Text = Val(Txstpmnt.Text) * 60 Then
        BTX03.Text = 0
        BTX04.Text = Val(BTX04.Text) + Val(Txststep.Text)
        BTX05.Text = Val(BTX04.Text) * Val(TXDISCOUNT.Text) / 100
        BTX06.Text = Val(BTX04.Text) - Val(BTX05.Text)
        End If
End If

If CTX01.Text = "1" Then CTX09.Text = Time
If CTX07.Text = "MOVE" Then
    CTX01.Text = "0": CTX02.Text = "": CTX03.Text = "0": CTX04.Text = Txststep.Text
    CTX05.Text = "0": CTX06.Text = Val(CTX04.Text) - Val(CTX05.Text)
    CTX07.Text = "Ready"
    CTX09.Text = ""
    ATX11.Text = ""
    End If
If CTX07.Text = "STOP" Then
    List2.AddItem "PC-" & CTX10.Caption & " Cost: " & CTX04.Text & " Duration: " & CTX02.Text
    List2.Selected(LISTC) = True
    With Adodc1.Recordset
        .AddNew
            !USER = CTX10.Caption
            !Date = TXDATE.Text
            !TIMEIN = CTX09.Text
            !Timeout = TXTIME.Text
            !DURATION = CTX02.Text
            !COST = CTX04.Text
            !DISCOUNT = CTX05.Text
            !PAYMENT = CTX06.Text
        .Update
        .MoveLast
    End With
    CTX01.Text = "0": CTX02.Text = "": CTX03.Text = "0": CTX04.Text = Txststep.Text
    CTX05.Text = "0": CTX06.Text = Val(CTX04.Text) - Val(CTX05.Text)
    CTX07.Text = "Ready"
    CTX09.Text = ""
    CTX11.Text = ""
    WRITEVALUES
End If
If CTX07.Text = "ONLINE" Then
    WRITEVALUES
    CTX01.Text = Val(CTX01.Text + 1)
    CTX02.Text = SecondsToTime(CTX01.Text)
    CTX03.Text = Val(CTX03 + 1)
    If CTX03.Text = Val(Txstpmnt.Text) * 60 Then
        CTX03.Text = 0
        CTX04.Text = Val(CTX04.Text) + Val(Txststep.Text)
        CTX05.Text = Val(CTX04.Text) * Val(TXDISCOUNT.Text) / 100
        CTX06.Text = Val(CTX04.Text) - Val(CTX05.Text)
        End If
End If

If DTX01.Text = "1" Then DTX09.Text = Time
If DTX07.Text = "MOVE" Then
    DTX01.Text = "0": DTX02.Text = "": DTX03.Text = "0": DTX04.Text = Txststep.Text
    DTX05.Text = "0": DTX06.Text = Val(DTX04.Text) - Val(DTX05.Text)
    DTX07.Text = "Ready"
    DTX09.Text = ""
    ATX11.Text = ""
    End If
If DTX07.Text = "STOP" Then
    List2.AddItem "PC-" & DTX10.Caption & " Cost: " & DTX04.Text & " Duration: " & DTX02.Text
    List2.Selected(LISTC) = True
    With Adodc1.Recordset
        .AddNew
            !USER = DTX10.Caption
            !Date = TXDATE.Text
            !TIMEIN = DTX09.Text
            !Timeout = TXTIME.Text
            !DURATION = DTX02.Text
            !COST = DTX04.Text
            !DISCOUNT = DTX05.Text
            !PAYMENT = DTX06.Text
        .Update
        .MoveLast
    End With
    DTX01.Text = "0": DTX02.Text = "": DTX03.Text = "0": DTX04.Text = Txststep.Text
    DTX05.Text = "0": DTX06.Text = Val(DTX04.Text) - Val(DTX05.Text)
    DTX07.Text = "Ready"
    DTX09.Text = ""
    DTX11.Text = ""
    WRITEVALUES
End If
If DTX07.Text = "ONLINE" Then
    WRITEVALUES
    DTX01.Text = Val(DTX01.Text + 1)
    DTX02.Text = SecondsToTime(DTX01.Text)
    DTX03.Text = Val(DTX03 + 1)
    If DTX03.Text = Val(Txstpmnt.Text) * 60 Then
        DTX03.Text = 0
        DTX04.Text = Val(DTX04.Text) + Val(Txststep.Text)
        DTX05.Text = Val(DTX04.Text) * Val(TXDISCOUNT.Text) / 100
        DTX06.Text = Val(DTX04.Text) - Val(DTX05.Text)
        End If
End If

If ETX01.Text = "1" Then ETX09.Text = Time
If ETX07.Text = "MOVE" Then
    ETX01.Text = "0": ETX02.Text = "": ETX03.Text = "0": ETX04.Text = Txststep.Text
    ETX05.Text = "0": ETX06.Text = Val(ETX04.Text) - Val(ETX05.Text)
    ETX07.Text = "Ready"
    ETX09.Text = ""
    ATX11.Text = ""
    End If
If ETX07.Text = "STOP" Then
    List2.AddItem "PC-" & ETX10.Caption & " Cost: " & ETX04.Text & " Duration: " & ETX02.Text
    List2.Selected(LISTC) = True
    With Adodc1.Recordset
        .AddNew
            !USER = ETX10.Caption
            !Date = TXDATE.Text
            !TIMEIN = ETX09.Text
            !Timeout = TXTIME.Text
            !DURATION = ETX02.Text
            !COST = ETX04.Text
            !DISCOUNT = ETX05.Text
            !PAYMENT = ETX06.Text
        .Update
        .MoveLast
    End With
    ETX01.Text = "0": ETX02.Text = "": ETX03.Text = "0": ETX04.Text = Txststep.Text
    ETX05.Text = "0": ETX06.Text = Val(ETX04.Text) - Val(ETX05.Text)
    ETX07.Text = "Ready"
    ETX09.Text = ""
    ETX11.Text = ""
    WRITEVALUES
End If
If ETX07.Text = "ONLINE" Then
    WRITEVALUES
    ETX01.Text = Val(ETX01.Text + 1)
    ETX02.Text = SecondsToTime(ETX01.Text)
    ETX03.Text = Val(ETX03 + 1)
    If ETX03.Text = Val(Txstpmnt.Text) * 60 Then
        ETX03.Text = 0
        ETX04.Text = Val(ETX04.Text) + Val(Txststep.Text)
        ETX05.Text = Val(ETX04.Text) * Val(TXDISCOUNT.Text) / 100
        ETX06.Text = Val(ETX04.Text) - Val(ETX05.Text)
        End If
End If
End Sub

Private Sub Timer2_Timer()
CLtimer.Text = Val(CLtimer + 1)
If CLtimer.Text > 20 Then CLtimer.Text = 0: bntClear_Click
End Sub
