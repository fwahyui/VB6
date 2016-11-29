VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmReport 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Patient Report"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form2"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FramePatHis 
      BackColor       =   &H00C0C0C0&
      Caption         =   "FramePatHis"
      Height          =   495
      Left            =   720
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   855
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Patient History"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2640
         TabIndex        =   12
         Top             =   360
         Width           =   2325
      End
   End
   Begin VB.CommandButton cmdPatHis 
      BackColor       =   &H00808080&
      Caption         =   "Patient history"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Frame frameAppApp 
      BackColor       =   &H00C0C0C0&
      Caption         =   "frameAppApp"
      Height          =   6975
      Left            =   1800
      TabIndex        =   1
      Top             =   840
      Width           =   9615
      Begin MSAdodcLib.Adodc Adocount 
         Height          =   375
         Left            =   3240
         Top             =   1920
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=clinic.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=clinic.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select count(*) from patient"
         Caption         =   "Adocount"
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
      Begin VB.TextBox txtcount 
         DataField       =   "Expr1000"
         DataSource      =   "Adocount"
         Height          =   285
         Left            =   7800
         TabIndex        =   7
         Top             =   2040
         Width           =   1575
      End
      Begin MSAdodcLib.Adodc Adoappapp 
         Height          =   375
         Left            =   3240
         Top             =   1560
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=clinic.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=clinic.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select p.id,p.name,p.address,p.age,p.sex,p.referto,d.reappointment from patient p,disease d where p.id=4"
         Caption         =   "Adoappapp"
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
      Begin MSDataGridLib.DataGrid Gridappapp 
         Bindings        =   "frmReport.frx":0000
         Height          =   4215
         Left            =   240
         TabIndex        =   6
         Top             =   2520
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   7435
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   14737632
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
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
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "id"
            Caption         =   "ID"
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
            DataField       =   "name"
            Caption         =   "Name"
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
         BeginProperty Column02 
            DataField       =   "address"
            Caption         =   "Address"
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
         BeginProperty Column03 
            DataField       =   "age"
            Caption         =   "Age"
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
         BeginProperty Column04 
            DataField       =   "sex"
            Caption         =   "Sex"
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
         BeginProperty Column05 
            DataField       =   "referto"
            Caption         =   "Checked By"
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
         BeginProperty Column06 
            DataField       =   "reappointment"
            Caption         =   "Re-appointment Date"
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
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1604.976
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1604.976
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   615.118
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1500.095
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1695.118
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdshowAppApp 
         BackColor       =   &H00808080&
         Caption         =   "Show"
         Height          =   375
         Left            =   2040
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1800
         Width           =   735
      End
      Begin VB.OptionButton OpAllFromToday 
         BackColor       =   &H00C0C0C0&
         Caption         =   "All From Today"
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Top             =   1440
         Width           =   1455
      End
      Begin VB.OptionButton opNextWeak 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Next Weak"
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   1080
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Approximate Appointments"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2760
         TabIndex        =   9
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "No Of Appointments:"
         Height          =   255
         Left            =   6000
         TabIndex        =   8
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "All Filtered ID's:"
         Height          =   195
         Left            =   600
         TabIndex        =   5
         Top             =   2160
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdAppApp 
      BackColor       =   &H00808080&
      Caption         =   "Appr. Appointment"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAppApp_Click()
Call FrameVisibility
 frameAppApp.Visible = True
End Sub
Function FrameVisibility()
 frameAppApp.Visible = False
 FramePatHis.Visible = False
End Function

Private Sub cmdPatHis_Click()
Call FrameVisibility
 FramePatHis.Visible = True
End Sub

Private Sub cmdshowAppApp_Click()
If opNextWeak Then
 Adoappapp.RecordSource = "select p.id,p.name,p.address,p.age,p.sex,p.referto,d.reappointment from patient p,disease d where p.id=d.id " & _
 "and reappointment between #" & Date & "# and #" & (Date + 7) & "# order by reappointment"
 Adocount.RecordSource = "select count(*) from patient p,disease d where p.id=d.id and reappointment between #" & Date & "# and #" & (Date + 7) & "#"
 
Else
 Adoappapp.RecordSource = "select p.id,p.name,p.address,p.age,p.sex,p.referto,d.reappointment from patient p,disease d where p.id=d.id " & _
 "and reappointment > #" & (Date - 1) & "# order by reappointment"
 Adocount.RecordSource = "select count(*) from patient p,disease d where p.id=d.id " & _
 "and reappointment > #" & (Date - 1) & "#"
End If
Adoappapp.Refresh
Gridappapp.Refresh
Adocount.Refresh
End Sub

Private Sub Form_Load()
 Call SetBasicFrameSettings
 Call cmdshowAppApp_Click
 frameAppApp.Visible = True
End Sub
Function SetBasicFrameSettings()
frameAppApp.Left = 1920
frameAppApp.Top = 960
frameAppApp.Width = 9615
frameAppApp.Height = 6975
frameAppApp.BorderStyle = 0

FramePatHis.Left = 1920
FramePatHis.Top = 960
FramePatHis.Width = 9615
FramePatHis.Height = 6975
FramePatHis.BorderStyle = 0
End Function
