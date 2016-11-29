VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmdaftarsiswa 
   Caption         =   "Daftar Siswa"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7185
   ScaleWidth      =   9840
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFC0C0&
      Height          =   7215
      Left            =   0
      ScaleHeight     =   7155
      ScaleWidth      =   9795
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      Begin TabDlg.SSTab SSTab1 
         Height          =   6855
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   12091
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "Daftar Siswa"
         TabPicture(0)   =   "frmdaftarbk.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "DataGrid1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Konseling"
         TabPicture(1)   =   "frmdaftarbk.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "CRViewer1"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Potensi Siswa"
         TabPicture(2)   =   "frmdaftarbk.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "CRViewer2"
         Tab(2).ControlCount=   1
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   6135
            Left            =   120
            TabIndex        =   2
            Top             =   480
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   10821
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   15
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
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin CRVIEWERLibCtl.CRViewer CRViewer2 
            Height          =   6255
            Left            =   -74880
            TabIndex        =   4
            Top             =   480
            Width           =   9255
            DisplayGroupTree=   -1  'True
            DisplayToolbar  =   -1  'True
            EnableGroupTree =   -1  'True
            EnableNavigationControls=   -1  'True
            EnableStopButton=   -1  'True
            EnablePrintButton=   -1  'True
            EnableZoomControl=   -1  'True
            EnableCloseButton=   -1  'True
            EnableProgressControl=   -1  'True
            EnableSearchControl=   -1  'True
            EnableRefreshButton=   -1  'True
            EnableDrillDown =   -1  'True
            EnableAnimationControl=   -1  'True
            EnableSelectExpertButton=   0   'False
            EnableToolbar   =   -1  'True
            DisplayBorder   =   -1  'True
            DisplayTabs     =   -1  'True
            DisplayBackgroundEdge=   -1  'True
            SelectionFormula=   ""
            EnablePopupMenu =   -1  'True
            EnableExportButton=   0   'False
            EnableSearchExpertButton=   0   'False
            EnableHelpButton=   0   'False
         End
         Begin CRVIEWERLibCtl.CRViewer CRViewer1 
            Height          =   6255
            Left            =   -74880
            TabIndex        =   3
            Top             =   480
            Width           =   9255
            DisplayGroupTree=   -1  'True
            DisplayToolbar  =   -1  'True
            EnableGroupTree =   -1  'True
            EnableNavigationControls=   -1  'True
            EnableStopButton=   -1  'True
            EnablePrintButton=   -1  'True
            EnableZoomControl=   -1  'True
            EnableCloseButton=   -1  'True
            EnableProgressControl=   -1  'True
            EnableSearchControl=   -1  'True
            EnableRefreshButton=   -1  'True
            EnableDrillDown =   -1  'True
            EnableAnimationControl=   -1  'True
            EnableSelectExpertButton=   0   'False
            EnableToolbar   =   -1  'True
            DisplayBorder   =   -1  'True
            DisplayTabs     =   -1  'True
            DisplayBackgroundEdge=   -1  'True
            SelectionFormula=   ""
            EnablePopupMenu =   -1  'True
            EnableExportButton=   0   'False
            EnableSearchExpertButton=   0   'False
            EnableHelpButton=   0   'False
         End
      End
   End
End
Attribute VB_Name = "frmdaftarsiswa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    Picture1.Top = 0
    Picture1.Left = 0
    Picture1.Height = ScaleHeight
    Picture1.Width = ScaleWidth
    SSTab1.Left = 120
    SSTab1.Top = 120
    SSTab1.Height = Picture1.Height - 240
    SSTab1.Width = Picture1.Width - 240
    CRViewer1.Top = 480
    CRViewer1.Left = 120
    CRViewer1.Height = SSTab1.Height - 600
    CRViewer1.Width = SSTab1.Width - 240
    DataGrid1.Left = 120
    DataGrid1.Top = 480
    DataGrid1.Height = SSTab1.Height - 600
    DataGrid1.Width = SSTab1.Width - 240
End Sub
