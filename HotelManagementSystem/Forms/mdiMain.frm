VERSION 5.00
Object = "{9EDDC69F-10E8-4DE7-9648-EC8A45F005C0}#1.0#0"; "b8Controls4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "3JNet Hotel Management System"
   ClientHeight    =   8085
   ClientLeft      =   225
   ClientTop       =   825
   ClientWidth     =   13485
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox bgHeader 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1080
      Left            =   0
      ScaleHeight     =   72
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   899
      TabIndex        =   18
      Top             =   0
      Width           =   13485
      Begin VB.PictureBox bgRecOpt 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5F5F5&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   765
         Left            =   3510
         ScaleHeight     =   51
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   950
         TabIndex        =   29
         Top             =   330
         Width           =   14250
         Begin b8Controls4.b8Line b8Line2 
            Height          =   30
            Left            =   0
            TabIndex        =   30
            Top             =   720
            Width           =   15720
            _ExtentX        =   27728
            _ExtentY        =   53
            BorderColor1    =   14737632
            BorderColor2    =   16777215
         End
         Begin MSComctlLib.Toolbar tbMenu 
            Height          =   810
            Left            =   0
            TabIndex        =   31
            Top             =   0
            Width           =   11580
            _ExtentX        =   20426
            _ExtentY        =   1429
            ButtonWidth     =   1217
            ButtonHeight    =   1429
            AllowCustomize  =   0   'False
            Wrappable       =   0   'False
            Style           =   1
            ImageList       =   "itb32x32"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   8
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "New"
                  Key             =   "New"
                  Object.ToolTipText     =   "Ctrl+F2"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Edit"
                  Key             =   "Edit"
                  Object.ToolTipText     =   "Ctrl+F3"
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Search"
                  Key             =   "Search"
                  Object.ToolTipText     =   "Ctrl+F4"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Delete"
                  Key             =   "Delete"
                  Object.ToolTipText     =   "Ctrl+F5"
                  ImageIndex      =   4
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Refresh"
                  Key             =   "Refresh"
                  Object.ToolTipText     =   "Ctrl+F6"
                  ImageIndex      =   5
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Print"
                  Key             =   "Print"
                  Object.ToolTipText     =   "Ctrl+F7"
                  ImageIndex      =   6
               EndProperty
               BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Close"
                  Key             =   "Close"
                  Object.ToolTipText     =   "Ctrl+F8"
                  ImageIndex      =   7
               EndProperty
            EndProperty
         End
      End
      Begin VB.PictureBox bgHeaderMenu 
         BackColor       =   &H00EDEBE9&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   0
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   1024
         TabIndex        =   19
         Top             =   0
         Width           =   15360
         Begin b8Controls4.b8Menu b8Menus 
            Height          =   285
            Index           =   0
            Left            =   60
            TabIndex        =   20
            Top             =   15
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   503
            BackColor       =   15592425
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   0
            Caption         =   "&System"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "&System"
            BorderColorNormal=   15592425
            BorderColorDown =   -2147483632
            BackColorNormal =   15592425
         End
         Begin b8Controls4.b8Menu b8Menus 
            Height          =   285
            Index           =   1
            Left            =   840
            TabIndex        =   21
            Top             =   15
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            BackColor       =   15592425
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   0
            Caption         =   "&Records"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "&Records"
            BorderColorNormal=   15592425
            BorderColorDown =   -2147483632
            BackColorNormal =   15592425
         End
         Begin b8Controls4.b8Menu b8Menus 
            Height          =   285
            Index           =   2
            Left            =   1740
            TabIndex        =   22
            Top             =   15
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   503
            BackColor       =   15592425
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   0
            Caption         =   "&Monitoring"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "&Monitoring"
            BorderColorNormal=   15592425
            BorderColorDown =   -2147483632
            BackColorNormal =   15592425
         End
         Begin b8Controls4.b8Menu b8Menus 
            Height          =   285
            Index           =   3
            Left            =   2730
            TabIndex        =   23
            Top             =   15
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   503
            BackColor       =   15592425
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   0
            Caption         =   "&Tools"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "&Tools"
            BorderColorNormal=   15592425
            BorderColorDown =   -2147483632
            BackColorNormal =   15592425
         End
         Begin b8Controls4.b8Menu b8Menus 
            Height          =   285
            Index           =   6
            Left            =   4680
            TabIndex        =   24
            Top             =   15
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   503
            BackColor       =   15592425
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   0
            Caption         =   "&Help"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "&Help"
            BorderColorNormal=   15592425
            BorderColorDown =   -2147483632
            BackColorNormal =   15592425
         End
         Begin b8Controls4.b8Line b8Line1 
            Height          =   30
            Left            =   0
            TabIndex        =   25
            Top             =   300
            Width           =   15720
            _ExtentX        =   27728
            _ExtentY        =   53
            BorderColor1    =   16119285
            BorderColor2    =   14737632
         End
         Begin b8Controls4.b8Menu b8Menus 
            Height          =   285
            Index           =   5
            Left            =   4050
            TabIndex        =   26
            Top             =   15
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   503
            BackColor       =   15592425
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   0
            Caption         =   "Re&ports"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Re&ports"
            BorderColorNormal=   15592425
            BorderColorDown =   -2147483632
            BackColorNormal =   15592425
         End
         Begin b8Controls4.b8Menu b8Menus 
            Height          =   285
            Index           =   4
            Left            =   3360
            TabIndex        =   32
            Top             =   15
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   503
            BackColor       =   15592425
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   0
            Caption         =   "&Actions"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "&Actions"
            BorderColorNormal=   15592425
            BorderColorDown =   -2147483632
            BackColorNormal =   15592425
         End
      End
      Begin b8Controls4.b8Line b8LLogoB 
         Height          =   30
         Left            =   0
         TabIndex        =   27
         Top             =   1050
         Visible         =   0   'False
         Width           =   15720
         _ExtentX        =   27728
         _ExtentY        =   53
         BorderColor1    =   14737632
         BorderColor2    =   16777215
      End
      Begin b8Controls4.b8SBtop b8SBT 
         Height          =   945
         Left            =   0
         TabIndex        =   28
         Top             =   330
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   1667
         MinWidth        =   180
         Begin VB.Image Image2 
            Height          =   540
            Left            =   690
            Picture         =   "mdiMain.frx":0000
            Top             =   150
            Width           =   1710
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   150
            Picture         =   "mdiMain.frx":34AB
            Top             =   150
            Width           =   480
         End
      End
   End
   Begin VB.Timer timeUpdateDate 
      Interval        =   1000
      Left            =   5040
      Top             =   1860
   End
   Begin b8Controls4.b8ClientWin b8CW 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7710
      Width           =   13485
      _ExtentX        =   23786
      _ExtentY        =   661
      Begin VB.PictureBox bgSystemBot 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   0
         ScaleHeight     =   405
         ScaleWidth      =   3495
         TabIndex        =   1
         Top             =   0
         Width           =   3495
         Begin VB.Image Image3 
            Height          =   360
            Left            =   0
            Picture         =   "mdiMain.frx":3D75
            Stretch         =   -1  'True
            Top             =   0
            Width           =   19995
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Programmed by: jaYPee"
            Height          =   285
            Left            =   120
            TabIndex        =   2
            Top             =   60
            Width           =   2385
         End
      End
   End
   Begin b8Controls4.b8SBCenter b8SBC 
      Align           =   3  'Align Left
      Height          =   6630
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   11695
      MinWidth        =   180
      Begin b8Controls4.b8SideTab b8ST 
         Height          =   345
         Index           =   0
         Left            =   60
         TabIndex        =   4
         Top             =   510
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   609
         Caption         =   "Quick Launch         [ Ctrl + Q ]"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Tahoma"
         FontSize        =   8.25
         ForeColor       =   8421504
         MaxHeight       =   3285
         BorderColor     =   12957347
         ContractedForeColor=   8421504
         ExpandedForeColor=   9594695
         Begin MSComctlLib.ListView listQL 
            Height          =   3075
            Left            =   30
            TabIndex        =   5
            Top             =   360
            Width           =   3345
            _ExtentX        =   5900
            _ExtentY        =   5424
            View            =   2
            Arrange         =   1
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            OLEDropMode     =   1
            _Version        =   393217
            Icons           =   "ilQL"
            SmallIcons      =   "ilQL"
            ForeColor       =   -2147483640
            BackColor       =   16777215
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            OLEDropMode     =   1
            NumItems        =   0
         End
         Begin MSComctlLib.ImageList ilQL 
            Left            =   1080
            Top             =   660
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   1
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdiMain.frx":3E4B
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
      Begin b8Controls4.b8SideTab b8ST 
         Height          =   345
         Index           =   1
         Left            =   60
         TabIndex        =   6
         Top             =   840
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   609
         Caption         =   "Search Item          [ Ctrl + S ]"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Tahoma"
         FontSize        =   8.25
         ForeColor       =   8421504
         MaxHeight       =   2245
         BorderColor     =   12957347
         ContractedForeColor=   8421504
         ExpandedForeColor=   9594695
         AutoContract    =   0   'False
         Begin VB.CommandButton cmdSearch 
            Caption         =   "Search"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2040
            TabIndex        =   9
            Top             =   1950
            Width           =   1215
         End
         Begin VB.ComboBox cmbLookIn 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1470
            Width           =   3135
         End
         Begin VB.TextBox txtSearchWhat 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   7
            Top             =   690
            Width           =   3165
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Look In:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   11
            Top             =   1260
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Search What:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   450
            Width           =   990
         End
      End
      Begin b8Controls4.b8SideTab b8ST 
         Height          =   345
         Index           =   2
         Left            =   60
         TabIndex        =   12
         Top             =   1170
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   609
         Caption         =   "Filter By Date        [ Ctrl + D ]"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Tahoma"
         FontSize        =   8.25
         ForeColor       =   8421504
         MaxHeight       =   2865
         BorderColor     =   12957347
         ContractedForeColor=   8421504
         ExpandedForeColor=   9594695
         AutoContract    =   0   'False
         Begin b8Controls4.b8DatePicker b8DateP 
            Height          =   2415
            Left            =   120
            TabIndex        =   13
            Top             =   420
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   4260
            BackColor       =   16777215
            MinDate         =   38968
            MaxDate         =   38968
         End
      End
      Begin VB.Label lblCurrentUser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   840
         TabIndex        =   17
         Top             =   45
         Width           =   180
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome,"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   165
         Left            =   150
         TabIndex        =   16
         Top             =   60
         Width           =   600
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   780
         TabIndex        =   15
         Top             =   255
         Width           =   180
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Today is "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   165
         Left            =   150
         TabIndex        =   14
         Top             =   270
         Width           =   585
      End
   End
   Begin MSComctlLib.ImageList i16x16g 
      Left            =   4260
      Top             =   2310
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
            Picture         =   "mdiMain.frx":3F5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":44F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":4A92
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":4E2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":51C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":5560
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ig24x24 
      Left            =   5745
      Top             =   2010
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   10
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":58FA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   4005
      Top             =   1290
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":5B27
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":6539
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":6F4B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":72E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":767F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":7A19
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":7DB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":87C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":91D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":9BE9
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":A5FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":B00D
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":BA1F
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":C431
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":C9CD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList itb32x32 
      Left            =   6480
      Top             =   1290
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":CF69
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":E8FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1028D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":11C1F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":135B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":14F43
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":168D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":18267
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":19BF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1B58D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1C269
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1CB49
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1D825
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1E501
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1F1DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1FEB9
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":20B95
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1140
      Top             =   330
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
            Picture         =   "mdiMain.frx":21471
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":21A0B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":21FA5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":2233F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":226D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":22A73
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   540
      Top             =   1530
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   10
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":22E0D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   540
      Top             =   330
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":2303A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":23A4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":2445E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":247F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":24B92
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":24F2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":252C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":25CD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":266EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":270FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":27B0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":28520
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":28F32
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":29944
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":29EE0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList4 
      Left            =   3015
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":2A47C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":2BE0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":2D7A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":2F132
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":30AC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":32456
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":33DE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":3577A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":3710C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":38AA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":3977C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":3A05C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":3AD38
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":3BA14
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":3C6F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":3D3CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":3E0A8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuSystem 
      Caption         =   "&System"
      Begin VB.Menu mnuManageUser 
         Caption         =   "&Manage Users"
      End
      Begin VB.Menu mnuChangeDeskClerk 
         Caption         =   "Change Desk Clerk"
      End
      Begin VB.Menu mnuS01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogOff 
         Caption         =   "&Log Off"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuRecords 
      Caption         =   "&Records"
      Begin VB.Menu mnuNewReservation 
         Caption         =   "&New Reservation"
      End
      Begin VB.Menu mnuCheckInHistory 
         Caption         =   "Check &In History"
      End
      Begin VB.Menu mnuS02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCustomers 
         Caption         =   "&Customers"
      End
      Begin VB.Menu mnuCompany 
         Caption         =   "Com&pany"
      End
      Begin VB.Menu mnuAccountReceivables 
         Caption         =   "Account Receivables"
      End
   End
   Begin VB.Menu mnuMonitoring 
      Caption         =   "&Monitoring"
      Begin VB.Menu mnuRoomS 
         Caption         =   "Rooms"
      End
      Begin VB.Menu mnuReservations 
         Caption         =   "Reservations"
      End
      Begin VB.Menu mnuInventoryView 
         Caption         =   "Inventory View"
      End
      Begin VB.Menu mnuBusinessSource 
         Caption         =   "Business Source"
      End
      Begin VB.Menu mnuChargeType 
         Caption         =   "Charge Type"
      End
      Begin VB.Menu mnuCountries 
         Caption         =   "Countries"
      End
      Begin VB.Menu mnuIDType 
         Caption         =   "ID Type"
      End
      Begin VB.Menu mnuPaymentType 
         Caption         =   "Payment Type"
      End
      Begin VB.Menu mnuRateType 
         Caption         =   "Rate Type"
      End
      Begin VB.Menu mnuRoomStatus 
         Caption         =   "Room Status"
      End
      Begin VB.Menu mnuRoomType 
         Caption         =   "Room Type"
      End
      Begin VB.Menu mnuVehicles 
         Caption         =   "Vehicles"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuPreferences 
         Caption         =   "&Preferences"
      End
      Begin VB.Menu mnuDatabaseUtilities 
         Caption         =   "&Database Utilities"
         Begin VB.Menu mnuBackupDatabase 
            Caption         =   "&Backup Database"
         End
         Begin VB.Menu mnuDatabaseRestore 
            Caption         =   "Database &Restore"
         End
      End
   End
   Begin VB.Menu mnuRecA 
      Caption         =   "&Action"
      Begin VB.Menu mnuRACN 
         Caption         =   "Create &New"
      End
      Begin VB.Menu mnuRAES 
         Caption         =   "&Edit Selected"
      End
      Begin VB.Menu mnuRAS 
         Caption         =   "&Search"
      End
      Begin VB.Menu mnuRADS 
         Caption         =   "&Delete Selected"
      End
      Begin VB.Menu mnuRARR 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuRAP 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuRASep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRAC 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnuRPTAccRec 
         Caption         =   "Accounts Receivable"
      End
      Begin VB.Menu mnuRPTOtherCharges 
         Caption         =   "Other Charges"
      End
      Begin VB.Menu mnuRPTCheckedInGuest 
         Caption         =   "Checked In Guest"
      End
      Begin VB.Menu mnuRPTCheckOut 
         Caption         =   "Check Out Report"
      End
      Begin VB.Menu mnuDueReservation 
         Caption         =   "Due Reservation"
      End
      Begin VB.Menu mnuRPTGuestList 
         Caption         =   "Guest List Report"
      End
      Begin VB.Menu mnuRPTRoomHistory 
         Caption         =   "Room History Report"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const m_TabShowQuickLaunch = 0
Private Const m_TabSearch = 1
Private Const m_TabFilterDate = 2

'Flag for User log
Public bUserLoggedOn As Boolean

'Control Procedures
'-----------------------------------------------------------
Private Sub b8CW_FormTabClick(ByVal sFormName As String, ByVal Index As Integer)
    modFuncChild.ActivateMDIChildForm sFormName
End Sub

Private Sub b8DateP_Change()
    Call Form_DateChange
End Sub

Private Sub b8SBC_BeforeResize(ByVal NewWidth As Integer)
    ResizeFb8SBC NewWidth
End Sub

Private Sub ResizeFb8SBC(ByVal NewWidth As Integer)
    
    'resize top side bar
    b8SBT.Width = NewWidth / Screen.TwipsPerPixelX
    bgSystemBot.Width = NewWidth
    'resize quick tabs
    Dim i As Integer
    For i = 0 To b8ST.UBound
        b8ST(i).Left = 60
        b8ST(i).Width = NewWidth - 120
    Next
    
    'resize window tab
    If b8SBC.Visible = True Then
        b8CW.SBWidth = NewWidth / Screen.TwipsPerPixelX
    Else
        b8CW.SBWidth = 0
    End If
    
    'call mdi resize to resize all opened mdi childs
    MDIForm_Resize
    
End Sub

Private Sub b8SBC_Resize()
    ResizeFb8SBC b8SBC.Width
End Sub

Private Sub b8SBT_Resize()
    b8SBC.Width = b8SBT.Width * Screen.TwipsPerPixelX
End Sub

Private Sub b8SBT_SizeChange(ByVal newSizeState As b8Controls4.eSizeState)
    
    If newSizeState = ssContracted Then
        b8CW.SBWidth = b8SBC.Width / Screen.TwipsPerPixelX
        b8SBC.Visible = True
        bgSystemBot.Visible = True
        b8LLogoB.Visible = False
    Else
        b8CW.SBWidth = 0
        b8SBC.Visible = False
        bgSystemBot.Visible = False
        b8LLogoB.Visible = True
    End If
    
    'call mdi resize to resize all opened child forms
    Call MDIForm_Resize
    
End Sub

Private Sub b8ST_BeforeExpand(Index As Integer)

    'resize contained controlsbeofre expanding
    Select Case Index
        Case m_TabSearch 'search
            'resize
            txtSearchWhat.Move 150, txtSearchWhat.Top, b8ST(Index).Width - 300
            cmbLookIn.Move 150, cmbLookIn.Top, txtSearchWhat.Width
            cmdSearch.Move b8ST(Index).Width - cmdSearch.Width - 150
        Case m_TabFilterDate 'filter date
            b8DateP.Move 150, b8DateP.Top, b8ST(Index).Width - 300
        
        Case m_TabShowQuickLaunch
            listQL.Move 150, listQL.Top, b8ST(Index).Width - 300

    End Select

End Sub

Private Sub b8ST_CompleteExpand(Index As Integer)
    Dim i As Integer
    
    For i = 0 To b8ST.UBound
        If Index <> i Then
            If b8ST(i).AutoContract = True Then
                b8ST(i).Expanded = False
            End If
        End If
    Next
End Sub

Private Sub b8ST_Resize(Index As Integer)
    
    Dim i As Integer
    
    For i = 1 To b8ST.UBound
        b8ST(i).Move b8ST(i).Left, (b8ST(i - 1).Top + b8ST(i - 1).Height) - 15
    Next
    
    If b8ST(Index).Expanded = True Then
        Select Case Index
            Case m_TabSearch 'search
                'resize
                txtSearchWhat.Move 150, txtSearchWhat.Top, b8ST(Index).Width - 300
                cmbLookIn.Move 150, cmbLookIn.Top, txtSearchWhat.Width
                cmdSearch.Move b8ST(Index).Width - cmdSearch.Width - 150
            
            Case m_TabFilterDate 'filter date
                b8DateP.Move 150, b8DateP.Top, b8ST(Index).Width - 300
            
            Case m_TabShowQuickLaunch
                listQL.Move 150, listQL.Top, b8ST(Index).Width - 300
                
        End Select
    End If

End Sub

Private Sub listQL_DblClick()

    Dim selItem As ListItem
    
    On Error GoTo RAE
    
    Set selItem = listQL.SelectedItem
    
    Select Case selItem.Key
        Case "reservation" 'Manage Products"
            If allowOpen("frmReservation", CurrUser.USER_NAME) = False Then
                MsgBox "You're not allowed to open this form." & vbCrLf & vbCrLf & _
                    "Please ask permission from your administrator.", vbInformation
                
                Exit Sub
            End If
            
            With frmReservation
                .Shortcut = True
                
                .Show vbModal
            End With
        Case "rooms_windows" 'Manage Supliers"
            If allowOpen("frmRoomsWindow", CurrUser.USER_NAME) = False Then
                MsgBox "You're not allowed to open this form." & vbCrLf & vbCrLf & _
                    "Please ask permission from your administrator.", vbInformation
                
                Exit Sub
            End If

            LoadForm frmRoomsWindow
        Case "cust" 'Manage Customers"
            If allowOpen("frmAllCustomer", CurrUser.USER_NAME) = False Then
                MsgBox "You're not allowed to open this form." & vbCrLf & vbCrLf & _
                    "Please ask permission from your administrator.", vbInformation
                
                Exit Sub
            End If

'            frmAllCustomer.ShowForm
        Case "poad" 'New P.O."
            If allowOpen("frmPOEntry", CurrUser.USER_NAME) = False Then
                MsgBox "You're not allowed to open this form." & vbCrLf & vbCrLf & _
                    "Please ask permission from your administrator.", vbInformation
                
                Exit Sub
            End If

'            frmPOEntry.ShowAdd
        Case "sale" 'New Sales Entry"
            If allowOpen("frmSIEntry", CurrUser.USER_NAME) = False Then
                MsgBox "You're not allowed to open this form." & vbCrLf & vbCrLf & _
                    "Please ask permission from your administrator.", vbInformation
                
                Exit Sub
            End If

'            frmSIEntry.ShowAdd
        Case "ppm" 'Purchases/Payments Mon."
            If allowOpen("frmAllPPM", CurrUser.USER_NAME) = False Then
                MsgBox "You're not allowed to open this form." & vbCrLf & vbCrLf & _
                    "Please ask permission from your administrator.", vbInformation
                
                Exit Sub
            End If

'            frmAllPPM.ShowForm
        Case "sicpm" 'Sales/Cust.Payments Mon."
            If allowOpen("frmAllSICPM", CurrUser.USER_NAME) = False Then
                MsgBox "You're not allowed to open this form." & vbCrLf & vbCrLf & _
                    "Please ask permission from your administrator.", vbInformation
                
                Exit Sub
            End If
            
'            frmAllSICPM.ShowForm
        Case "void" 'Void Products Mon."
            If allowOpen("frmAllVoid", CurrUser.USER_NAME) = False Then
                MsgBox "You're not allowed to open this form." & vbCrLf & vbCrLf & _
                    "Please ask permission from your administrator.", vbInformation
                
                Exit Sub
            End If
            
'            frmAllVoid.ShowForm
        Case "stock" 'Stock Inventory"
            If allowOpen("frmAllStockInv", CurrUser.USER_NAME) = False Then
                MsgBox "You're not allowed to open this form." & vbCrLf & vbCrLf & _
                    "Please ask permission from your administrator.", vbInformation
                
                Exit Sub
            End If
            
'            frmAllStockInv.ShowForm
        Case "checkcust" 'Manage Due Checks (Cust.)"
            If allowOpen("frmAllCustPayDueCheck", CurrUser.USER_NAME) = False Then
                MsgBox "You're not allowed to open this form." & vbCrLf & vbCrLf & _
                    "Please ask permission from your administrator.", vbInformation
                
                Exit Sub
            End If
            
'            frmAllCustPayDueCheck.ShowForm
        Case "checksupp" 'Manage Due Checks (Supp.)"
            If allowOpen("frmAllPTSDueCheck", CurrUser.USER_NAME) = False Then
                MsgBox "You're not allowed to open this form." & vbCrLf & vbCrLf & _
                    "Please ask permission from your administrator.", vbInformation
                
                Exit Sub
            End If
            
'            frmAllPTSDueCheck.ShowForm
    End Select

RAE:
    Set selItem = Nothing
End Sub

Private Sub MDIForm_Load()
    
    'show weclome
    frmWelcome.ShowForm
    
    'set menus
'    Set b8Menus(0).Menu = Me.mnuSystem
'    Set b8Menus(1).Menu = Me.mnuRecords
'    Set b8Menus(2).Menu = Me.mnuMonitoring
'    Set b8Menus(3).Menu = Me.mnuTools
'    Set b8Menus(4).Menu = Me.mnuRecA
'    Set b8Menus(5).Menu = Me.mnuReports
'    Set b8Menus(6).Menu = Me.mnuHelp
    
    'add quick launch items
    listQL.ListItems.Add , "reservation", "New Reservation", 1, 1
    listQL.ListItems.Add , "rooms_windows", "Rooms", 1, 1

    DisplayUserInfo

    HideTBButton "", True
    frmWelcome.Active
End Sub

Private Sub mnuAccountReceivables_Click()
    If allowOpen("frmAccountReceivableList", CurrUser.USER_NAME) = False Then
        MsgBox "You're not allowed to open this form." & vbCrLf & vbCrLf & _
            "Please ask permission from your administrator.", vbInformation
        
        Exit Sub
    End If
    
    LoadForm frmAccountReceivableList
End Sub

Private Sub mnuBackupDatabase_Click()
    If allowOpen("frmDBBackup", CurrUser.USER_NAME) = False Then
        MsgBox "You're not allowed to open this form." & vbCrLf & vbCrLf & _
            "Please ask permission from your administrator.", vbInformation
        
        Exit Sub
    End If
    
'    frmDBBackup.ShowForm
End Sub

Private Sub mnuBusinessSource_Click()
    If allowOpen("frmBusinessSourceList", CurrUser.USER_NAME) = False Then
        MsgBox "You're not allowed to open this form." & vbCrLf & vbCrLf & _
            "Please ask permission from your administrator.", vbInformation
        
        Exit Sub
    End If
    
    LoadForm frmBusinessSourceList
End Sub

Private Sub mnuChangeDeskClerk_Click()
    frmCloseShift.Show 1
End Sub

Private Sub mnuChargeType_Click()
    If allowOpen("frmChargeTypeList", CurrUser.USER_NAME) = False Then
        MsgBox "You're not allowed to open this form." & vbCrLf & vbCrLf & _
            "Please ask permission from your administrator.", vbInformation
        
        Exit Sub
    End If
    
    LoadForm frmChargeTypeList
End Sub

Private Sub mnuCheckInHistory_Click()
    If allowOpen("frmCheckInList", CurrUser.USER_NAME) = False Then
        MsgBox "You're not allowed to open this form." & vbCrLf & vbCrLf & _
            "Please ask permission from your administrator.", vbInformation
        
        Exit Sub
    End If
    
    LoadForm frmCheckInList
End Sub

Private Sub mnuCompany_Click()
    If allowOpen("frmCompanyList", CurrUser.USER_NAME) = False Then
        MsgBox "You're not allowed to open this form." & vbCrLf & vbCrLf & _
            "Please ask permission from your administrator.", vbInformation
        
        Exit Sub
    End If
    
    LoadForm frmCompanyList
End Sub

Private Sub mnuCountries_Click()
    If allowOpen("frmCountriesList", CurrUser.USER_NAME) = False Then
        MsgBox "You're not allowed to open this form." & vbCrLf & vbCrLf & _
            "Please ask permission from your administrator.", vbInformation
        
        Exit Sub
    End If
    
    LoadForm frmCountriesList
End Sub

Private Sub mnuCustomers_Click()
    If allowOpen("frmCustomersList", CurrUser.USER_NAME) = False Then
        MsgBox "You're not allowed to open this form." & vbCrLf & vbCrLf & _
            "Please ask permission from your administrator.", vbInformation
        
        Exit Sub
    End If
    
    LoadForm frmCustomersList
End Sub

Private Sub mnuDatabaseRestore_Click()
'    If allowOpen("frmRestore", CurrUser.USER_NAME) = False Then
'        MsgBox "You're not allowed to open this form." & vbCrLf & vbCrLf & _
'            "Please ask permission from your administrator.", vbInformation
'
'        Exit Sub
'    End If
'
'    frmRestore.ShowForm
End Sub

Private Sub mnuDueReservation_Click()
    If allowOpen("frmRPTDueReservation", CurrUser.USER_NAME) = False Then
        MsgBox "You're not allowed to open this form." & vbCrLf & vbCrLf & _
            "Please ask permission from your administrator.", vbInformation
        
        Exit Sub
    End If
    
    frmRPTDueReservation.Show vbModal
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuIDType_Click()
    If allowOpen("frmIDTypeList", CurrUser.USER_NAME) = False Then
        MsgBox "You're not allowed to open this form." & vbCrLf & vbCrLf & _
            "Please ask permission from your administrator.", vbInformation
        
        Exit Sub
    End If
    
    LoadForm frmIDTypeList
End Sub

Private Sub mnuLogOff_Click()
    If MsgBox("Are you sure you want to log out?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

    
    'SendMessage frmShortcuts.hwnd, WM_CLOSE, 0, 0
    UnloadChilds
'    SendMessage frmShortcuts.hWnd, WM_ACTIVATE, 0, 0
    
    'ClearInfoMsg
'    StatusBar1.Panels(3).Text = ""
'    StatusBar1.Panels(4).Text = ""
    
    CurrUser.USER_NAME = ""
    CurrUser.USER_PK = 0
    
    frmLogin.Show vbModal: If CloseMe = True Then Unload Me: Exit Sub: Exit Sub
'    DisplayUserInfo
End Sub

Private Sub mnuManageUser_Click()
    frmUsersList.Show vbModal
End Sub

Private Sub mnuNewReservation_Click()
    If allowOpen("frmReservation", CurrUser.USER_NAME) = False Then
        MsgBox "You're not allowed to open this form." & vbCrLf & vbCrLf & _
            "Please ask permission from your administrator.", vbInformation
        
        Exit Sub
    End If
    
    frmReservation.Show vbModal
End Sub

Private Sub mnuPaymentType_Click()
    If allowOpen("frmPaymentTypeList", CurrUser.USER_NAME) = False Then
        MsgBox "You're not allowed to open this form." & vbCrLf & vbCrLf & _
            "Please ask permission from your administrator.", vbInformation
        
        Exit Sub
    End If
    
    LoadForm frmPaymentTypeList
End Sub

Private Sub mnuPreferences_Click()
'    If allowOpen("frmPref", CurrUser.USER_NAME) = False Then
'        MsgBox "You're not allowed to open this form." & vbCrLf & vbCrLf & _
'            "Please ask permission from your administrator.", vbInformation
'
'        Exit Sub
'    End If
'
'    frmPref.ShowForm 0
End Sub

Private Sub mnuRateType_Click()
    If allowOpen("frmRateTypeList", CurrUser.USER_NAME) = False Then
        MsgBox "You're not allowed to open this form." & vbCrLf & vbCrLf & _
            "Please ask permission from your administrator.", vbInformation
        
        Exit Sub
    End If
    
    LoadForm frmRateTypeList
End Sub

Private Sub mnuReservations_Click()
    If allowOpen("frmReservationList", CurrUser.USER_NAME) = False Then
        MsgBox "You're not allowed to open this form." & vbCrLf & vbCrLf & _
            "Please ask permission from your administrator.", vbInformation
        
        Exit Sub
    End If
    
    LoadForm frmReservationList
End Sub

Private Sub mnuRooms_Click()
    If allowOpen("frmRoomsList", CurrUser.USER_NAME) = False Then
        MsgBox "You're not allowed to open this form." & vbCrLf & vbCrLf & _
            "Please ask permission from your administrator.", vbInformation
        
        Exit Sub
    End If

    LoadForm frmRoomsList
End Sub

Private Sub mnuRoomStatus_Click()
    If allowOpen("frmRoomStatusList", CurrUser.USER_NAME) = False Then
        MsgBox "You're not allowed to open this form." & vbCrLf & vbCrLf & _
            "Please ask permission from your administrator.", vbInformation
        
        Exit Sub
    End If
    
    LoadForm frmRoomStatusList
End Sub

Private Sub mnuRoomType_Click()
    If allowOpen("frmRoomTypeList", CurrUser.USER_NAME) = False Then
        MsgBox "You're not allowed to open this form." & vbCrLf & vbCrLf & _
            "Please ask permission from your administrator.", vbInformation
        
        Exit Sub
    End If
    
    LoadForm frmRoomTypeList
End Sub

Private Sub mnuRPTAccRec_Click()
    If allowOpen("frmRPTAccRec", CurrUser.USER_NAME) = False Then
        MsgBox "You're not allowed to open this form." & vbCrLf & vbCrLf & _
            "Please ask permission from your administrator.", vbInformation
        
        Exit Sub
    End If
    
    frmRPTAccRec.Show vbModal
End Sub

Private Sub mnuRPTCheckedInGuest_Click()
    If allowOpen("rpt_CheckIn_Guest", CurrUser.USER_NAME) = False Then
        MsgBox "You're not allowed to open this form." & vbCrLf & vbCrLf & _
            "Please ask permission from your administrator.", vbInformation
        
        Exit Sub
    End If
    
    With frmReports
        .strReport = "CheckIn Guest"

        frmReports.Show vbModal
    End With
End Sub

Private Sub mnuRPTCheckOut_Click()
    If allowOpen("frmRPTCheckOut", CurrUser.USER_NAME) = False Then
        MsgBox "You're not allowed to open this form." & vbCrLf & vbCrLf & _
            "Please ask permission from your administrator.", vbInformation
        
        Exit Sub
    End If
    
    frmRPTCheckOut.Show vbModal
End Sub

Private Sub mnuRPTGuestList_Click()
    If allowOpen("rpt_Guest_List", CurrUser.USER_NAME) = False Then
        MsgBox "You're not allowed to open this form." & vbCrLf & vbCrLf & _
            "Please ask permission from your administrator.", vbInformation
        
        Exit Sub
    End If
    
    With frmReports
        .strReport = "Guest List"

        frmReports.Show vbModal
    End With
End Sub

Private Sub mnuRPTOtherCharges_Click()
    If allowOpen("frmRPTOtherCharges", CurrUser.USER_NAME) = False Then
        MsgBox "You're not allowed to open this form." & vbCrLf & vbCrLf & _
            "Please ask permission from your administrator.", vbInformation
        
        Exit Sub
    End If
    
    frmRPTOtherCharges.Show vbModal
End Sub

Private Sub mnuInventoryView_Click()
    If allowOpen("frmInventoryView", CurrUser.USER_NAME) = False Then
        MsgBox "You're not allowed to open this form." & vbCrLf & vbCrLf & _
            "Please ask permission from your administrator.", vbInformation
        
        Exit Sub
    End If
    
    LoadForm frmInventoryView
End Sub

Private Sub mnuRPTRoomHistory_Click()
    If allowOpen("frmRPTRoomHistory", CurrUser.USER_NAME) = False Then
        MsgBox "You're not allowed to open this form." & vbCrLf & vbCrLf & _
            "Please ask permission from your administrator.", vbInformation
        
        Exit Sub
    End If
    
    frmRPTRoomHistory.Show 1
End Sub

Private Sub mnuVehicles_Click()
    If allowOpen("frmVehiclesList", CurrUser.USER_NAME) = False Then
        MsgBox "You're not allowed to open this form." & vbCrLf & vbCrLf & _
            "Please ask permission from your administrator.", vbInformation
        
        Exit Sub
    End If
    
    LoadForm frmVehiclesList
End Sub

Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    ActiveForm.CommandPass Button.Key
End Sub

Private Sub timeUpdateDate_Timer()
    lblDate.Caption = FormatDateTime(Now, vbGeneralDate)
End Sub

'Private Sub txtSearchWhat_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        Call cmdSearch_Click
'    End If
'End Sub

'Private Sub cmdSearch_Click()
'    Form_Search
'End Sub

'-----------------------------------------------------------
' end Control Procedures


' MDI Form procedures
'-----------------------------------------------------------
Private Sub MDIForm_Resize()
    Dim frm As Form
    
    On Error Resume Next
    
    'resize header menus bg
    'bgHeaderMenu.Left = b8SBC.Width / Screen.TwipsPerPixelX
    
    'resize bg Record Opt
    bgRecOpt.Move b8SBC.Width / Screen.TwipsPerPixelX
    
    'resize childs
    If GetActiveChildCount > 0 Then
        For Each frm In Forms
        If frm.Name <> Me.Name Then
            If frm.MDIChild = True Then
                If frm.Name = Me.ActiveForm.Name Then
                    ResizeMdiChildForm frm
                Else
                    frm.Visible = False
                End If
            End If
        End If
        
        Next
        
    End If
    
    Set frm = Nothing
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    End
End Sub
'Get Opened MDI Child Forms Count
Public Function GetActiveChildCount() As Integer
    
    Dim frm As Form
    Dim iCount As Integer
    
    iCount = 0
    
    For Each frm In Forms
        If frm.Name <> Me.Name Then
            If frm.MDIChild = True Then
                iCount = iCount + 1
            End If
        End If
    Next
    
    GetActiveChildCount = iCount
    Set frm = Nothing
    
End Function

'-----------------------------------------------------------
' >> End MDI Form procedures
'------------------------------------------------------------



'------------------------------------------------------------
' Parent To Child procedures
'------------------------------------------------------------

Public Sub AddChild(ByRef CFrm As Form)

    'load form
    modFuncChild.LoadForm CFrm
    
End Sub

Public Sub ActivateChild(ByRef CFrm As Form)

    'activate form
    Me.b8CW.SetActiveWindow CFrm.Name
    
    'refresh record operation buttons
'    Form_CanAdd
'    Form_CanEdit
'    Form_CanDelete
'    Form_CanRefresh
'    Form_CanPrint
'    Form_CanSearch
'    Form_SetSearch

End Sub

Public Sub RemoveChild(ByVal sFormName As String)
    
    'remove form
    Me.b8CW.RemoveChildWindow sFormName
    
End Sub

'Record Operation

'Public Function Form_CanAdd() As Boolean
'
'    Dim bReturn As Boolean
'
'    On Error Resume Next
'
'    bReturn = False
'    bReturn = Me.ActiveForm.Form_CanAdd
'
'    b8RecOpt(0).Enabled = bReturn
'
'    Form_CanAdd = bReturn
'
'    err.Clear
'
'End Function

'Public Function Form_CanEdit() As Boolean
'
'    Dim bReturn As Boolean
'
'    On Error Resume Next
'
'    bReturn = False
'    bReturn = Me.ActiveForm.Form_CanEdit
'
'    b8RecOpt(1).Enabled = bReturn
'
'    Form_CanEdit = bReturn
'
'    err.Clear
'
'End Function

'Public Function Form_Edit()
'
'    If Form_CanEdit Then
'        Me.ActiveForm.Form_Edit
'    End If
'
'End Function


'Public Function Form_CanDelete() As Boolean
'
'    Dim bReturn As Boolean
'
'    On Error Resume Next
'
'    bReturn = False
'    bReturn = Me.ActiveForm.Form_CanDelete
'
'    b8RecOpt(2).Enabled = bReturn
'
'    Form_CanDelete = bReturn
'
'    err.Clear
'
'End Function


'Public Function Form_Delete()
'
'    If Form_CanDelete Then
'        Me.ActiveForm.Form_Delete
'    End If
'
'End Function


'Public Function Form_CanRefresh() As Boolean
'
'    Dim bReturn As Boolean
'
'    On Error Resume Next
'
'    bReturn = False
'    bReturn = Me.ActiveForm.Form_CanRefresh
'
'    b8RecOpt(3).Enabled = bReturn
'
'    Form_CanRefresh = bReturn
'
'    err.Clear
'
'End Function


'Public Function Form_Refresh()
'
'    If Form_CanRefresh Then
'        Me.ActiveForm.Form_Refresh
'    End If
'
'End Function



'Public Function Form_CanPrint() As Boolean
'
'    Dim bReturn As Boolean
'
'    On Error Resume Next
'
'    bReturn = False
'    bReturn = Me.ActiveForm.Form_CanPrint
'
'    b8RecOpt(4).Enabled = bReturn
'
'    Form_CanPrint = bReturn
'
'    err.Clear
'
'End Function

'Public Function Form_Print()
'    If Form_CanPrint Then
'        Me.ActiveForm.Form_Print
'    End If
'End Function


'Public Function Form_CanSearch() As Boolean
'
'    Dim bReturn As Boolean
'
'    On Error Resume Next
'
'    bReturn = False
'    bReturn = Me.ActiveForm.Form_CanSearch
'
'    Form_CanSearch = bReturn
'
'    err.Clear
'
'End Function



Public Function Form_ShowQuickLaunch()

    'expand side bar
    If b8SBT.SizeState <> ssContracted Then
        b8SBT.SizeState = ssContracted
    End If

    'expand search tab
    If b8ST(m_TabShowQuickLaunch).Expanded = False Then
        b8ST(m_TabShowQuickLaunch).Expanded = True
    End If
    
    On Error Resume Next
    b8ST(m_TabShowQuickLaunch).SetFocus
    'HLTxt txtSearchWhat
    err.Clear
    
End Function

'Public Function Form_ShowSearch()
'
'    'expand side bar
'    If b8SBT.SizeState <> ssContracted Then
'        b8SBT.SizeState = ssContracted
'    End If
'
'    'expand search tab
'    If b8ST(m_TabSearch).Expanded = False Then
'        b8ST(m_TabSearch).Expanded = True
'    End If
'
'    On Error Resume Next
'    b8ST(m_TabSearch).SetFocus
'    HLTxt txtSearchWhat
'    err.Clear
'
'End Function


Public Function Form_ShowDateFilter()

    'expand side bar
    If b8SBT.SizeState <> ssContracted Then
        b8SBT.SizeState = ssContracted
    End If

    'expand search tab
    If b8ST(m_TabFilterDate).Expanded = False Then
        b8ST(m_TabFilterDate).Expanded = True
    End If
    
    On Error Resume Next
    b8ST(m_TabFilterDate).SetFocus
    b8DateP.SetFocus
    err.Clear
    
End Function


Public Function Form_SetSearch()
    Dim bReturn As Boolean
    Dim sFields() As String
    Dim i  As Integer
    
    'clear
    txtSearchWhat.Text = ""
    cmbLookIn.Clear
    
    On Error Resume Next
    
    bReturn = False
    bReturn = Me.ActiveForm.Form_SetSearch(sFields)

    txtSearchWhat.Enabled = bReturn
    cmbLookIn.Enabled = bReturn
    cmdSearch.Enabled = bReturn
    
    If bReturn = True Then
        cmbLookIn.AddItem "All Fields"
        cmbLookIn.ListIndex = 0
        If UBound(sFields) >= 0 Then
            For i = 0 To UBound(sFields)
                cmbLookIn.AddItem sFields(i)
            Next
        End If
    Else
        'contract search tab if it was expanded
        If b8ST(m_TabSearch).Expanded = True Then
            b8ST(m_TabSearch).Expanded = False
        End If
        
    End If
    
    Form_SetSearch = bReturn
    
    err.Clear
End Function


'Public Function Form_Search()
'
'    Dim bResult As Boolean
'
'    'default
'    bResult = False
'
'
'    On Error GoTo errh
'
'    If txtSearchWhat.Text = "Enter text here" Then
'        txtSearchWhat.Text = ""
'    End If
'
'    If Len(Trim(txtSearchWhat.Text)) <= 0 Then
'        MsgBox "Please enter text to search.", vbExclamation
'        txtSearchWhat.Text = "Enter text here"
'        HLTxt txtSearchWhat
'        GoTo errh
'    End If
'
'    If Len(Trim(cmbLookIn.Text)) <= 0 Then
'        MsgBox "Please enter valid field.", vbExclamation
'        cmbLookIn.SetFocus
'        GoTo errh
'    End If
'
'
'    bResult = Me.ActiveForm.Form_Search(Trim(txtSearchWhat.Text), Trim(cmbLookIn.Text))
'
'    If bResult = False Then
'        MsgBox "Cannot find '" & txtSearchWhat.Text & "'", vbExclamation
'        HLTxt txtSearchWhat
'    End If
'
'errh:
'    err.Clear
'
'End Function

Public Sub Form_DateChange()

    On Error GoTo errh
    Me.ActiveForm.Form_DateChange
errh:
End Sub

Public Function Form_StartBussy()
    Me.MousePointer = vbHourglass
End Function

Public Function Form_EndBussy()
    Me.MousePointer = vbDefault
End Function

Public Sub AFForm_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 83 And Shift = 4 Then
        b8Menus(0).ShowPopUp
    ElseIf KeyCode = 82 And Shift = 4 Then
        b8Menus(1).ShowPopUp
    ElseIf KeyCode = 77 And Shift = 4 Then
        b8Menus(2).ShowPopUp
    ElseIf KeyCode = 84 And Shift = 4 Then
        b8Menus(3).ShowPopUp
    ElseIf KeyCode = 72 And Shift = 4 Then
        b8Menus(4).ShowPopUp
        
    ElseIf KeyCode = 81 And Shift = 2 Then
        'Ctrl + Q
        Me.Form_ShowQuickLaunch
    ElseIf KeyCode = 68 And Shift = 2 Then
        'Ctrl + D
        Me.Form_ShowDateFilter
    End If
    
    'MsgBox KeyCode & " - " & Shift
End Sub
'------------------------------------------------------------
' >>> Parent To Child procedures


'Member variables property
Public Property Get TabSearchIndex() As Integer
    TabSearchIndex = m_TabSearch
End Property

Public Sub HideTBButton(ByVal srcPatern As String, Optional srcAllButton As Boolean)
    If srcAllButton = True Then srcPatern = "ttttttt"
    If Mid$(srcPatern, 1, 1) = "t" Then tbMenu.Buttons(2).Visible = False: mnuRACN.Visible = False
    If Mid$(srcPatern, 2, 1) = "t" Then tbMenu.Buttons(3).Visible = False: mnuRAES.Visible = False
    If Mid$(srcPatern, 3, 1) = "t" Then tbMenu.Buttons(4).Visible = False: mnuRAS.Visible = False
    If Mid$(srcPatern, 4, 1) = "t" Then tbMenu.Buttons(5).Visible = False: mnuRADS.Visible = False
    If Mid$(srcPatern, 5, 1) = "t" Then tbMenu.Buttons(6).Visible = False: mnuRARR.Visible = False
    If Mid$(srcPatern, 6, 1) = "t" Then tbMenu.Buttons(7).Visible = False: mnuRAP.Visible = False
    If Mid$(srcPatern, 7, 1) = "t" Then tbMenu.Buttons(8).Visible = False: mnuRAC.Visible = False
'    If mnuRAC.Visible = False Then mnuRASep2.Visible = False
End Sub

Public Sub ShowTBButton(ByVal srcPatern As String, Optional srcAllButton As Boolean)
    'Highligh active form in opened form list
    If srcAllButton = True Then srcPatern = "ttttttt"
    If Mid$(srcPatern, 1, 1) = "t" Then tbMenu.Buttons(2).Visible = True: mnuRACN.Visible = True
    If Mid$(srcPatern, 2, 1) = "t" Then tbMenu.Buttons(3).Visible = True: mnuRAES.Visible = True
    If Mid$(srcPatern, 3, 1) = "t" Then tbMenu.Buttons(4).Visible = True: mnuRAS.Visible = True
    If Mid$(srcPatern, 4, 1) = "t" Then tbMenu.Buttons(5).Visible = True: mnuRADS.Visible = True
    If Mid$(srcPatern, 5, 1) = "t" Then tbMenu.Buttons(6).Visible = True: mnuRARR.Visible = True
    If Mid$(srcPatern, 6, 1) = "t" Then tbMenu.Buttons(7).Visible = True: mnuRAP.Visible = True
    If Mid$(srcPatern, 7, 1) = "t" Then tbMenu.Buttons(8).Visible = True: mnuRAC.Visible = True
'    If mnuRAC.Visible = True Then mnuRASep2.Visible = True
End Sub

Public Sub UnloadChilds()
''Unload all active forms
    Dim Form As Form
    
    For Each Form In Forms
       ''Unload all active childs
       If Form.Name <> Me.Name And Form.Name <> "frmShortcuts" Then Unload Form
    Next Form
   
    Set Form = Nothing
End Sub

Private Sub DisplayUserInfo()
    'Display the current user info
'    If CurrUser.USER_ISADMIN = True Then
'        StatusBar1.Panels(4).Text = "Admin"
'    Else
'        StatusBar1.Panels(4).Text = "Operator"
'    End If
'    StatusBar1.Panels(3).Text = CurrUser.USER_NAME
    
    Dim RS As New Recordset
    
    RS.Open "SELECT * FROM [Business Info]", CN, adOpenStatic, adLockReadOnly
    
    CurrBiz.BUSINESS_NAME = RS.Fields(0)
    CurrBiz.BUSINESS_ADDRESS = RS.Fields(1)
    CurrBiz.BUSINESS_CONTACT_INFO = RS.Fields(2)
    
    Set RS = Nothing
End Sub

