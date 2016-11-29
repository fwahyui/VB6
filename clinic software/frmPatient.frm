VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPatient 
   BorderStyle     =   0  'None
   Caption         =   "-------------------------------------------------------------------------------------------"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   Picture         =   "frmPatient.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "About Us"
      Height          =   375
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   8040
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Reports"
      Height          =   375
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   8040
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Change Password"
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   8040
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Exit"
      Height          =   375
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   8040
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   15
      Left            =   4560
      TabIndex        =   40
      Top             =   600
      Width           =   15
   End
   Begin VB.TextBox Atxtdose 
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   9720
      TabIndex        =   36
      Text            =   " "
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton Acmdadd 
      BackColor       =   &H00808080&
      Caption         =   "Add"
      Height          =   255
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   5400
      Width           =   855
   End
   Begin VB.TextBox Atxtmedicine 
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   7560
      TabIndex        =   33
      Text            =   " "
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox Atxtdesease 
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   7560
      TabIndex        =   17
      Text            =   " "
      Top             =   3480
      Width           =   3735
   End
   Begin VB.TextBox Atxtreappointment 
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   7560
      TabIndex        =   16
      Text            =   " "
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton AcmdSavePrint 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Save && Print"
      Height          =   375
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7320
      Width           =   1215
   End
   Begin VB.TextBox AtxtFee 
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   9720
      TabIndex        =   14
      Text            =   " "
      Top             =   3960
      Width           =   1455
   End
   Begin VB.TextBox BtxtID 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   1680
      TabIndex        =   7
      Text            =   " "
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox Btxtname 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   1680
      TabIndex        =   6
      Text            =   " "
      Top             =   3240
      Width           =   1935
   End
   Begin VB.TextBox Btxtage 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   1680
      TabIndex        =   5
      Text            =   " "
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox Btxtaddress 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   1680
      TabIndex        =   4
      Text            =   " "
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CommandButton BcmdSave 
      BackColor       =   &H00808080&
      Caption         =   "Save Only"
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6360
      Width           =   1095
   End
   Begin VB.TextBox Btxtreferto 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   1680
      TabIndex        =   2
      Text            =   " "
      Top             =   5640
      Width           =   1935
   End
   Begin VB.ComboBox Btxtsex 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   330
      ItemData        =   "frmPatient.frx":B452
      Left            =   1680
      List            =   "frmPatient.frx":B45C
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdSaveShift 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Save and/or  >> "
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7440
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid MedicineFlex 
      Height          =   1935
      Left            =   6600
      TabIndex        =   34
      Top             =   5760
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   3413
      _Version        =   393216
      Rows            =   12
      Cols            =   3
      FixedCols       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Control Panel:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1440
      TabIndex        =   41
      Top             =   8120
      Width           =   1515
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   3
      Left            =   7440
      Shape           =   4  'Rounded Rectangle
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   2
      Left            =   9600
      Shape           =   4  'Rounded Rectangle
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   1
      Left            =   9600
      Shape           =   4  'Rounded Rectangle
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   5
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   4
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   3
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   4920
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   2
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   1
      Left            =   7440
      Shape           =   4  'Rounded Rectangle
      Top             =   3360
      Width           =   3975
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   0
      Left            =   7440
      Shape           =   4  'Rounded Rectangle
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   0
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter List of medicine:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   7800
      TabIndex        =   39
      Top             =   4560
      Width           =   2565
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Medicine:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   210
      Left            =   6360
      TabIndex        =   38
      Top             =   5040
      Width           =   795
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dose:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   210
      Left            =   9120
      TabIndex        =   37
      Top             =   5040
      Width           =   465
   End
   Begin VB.Label Atxtname 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "----------"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   10320
      TabIndex        =   32
      Top             =   1800
      Width           =   600
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Patient ID:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   210
      Left            =   6240
      TabIndex        =   31
      Top             =   1800
      Width           =   810
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Age:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   210
      Left            =   6240
      TabIndex        =   30
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Atxtage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "----------"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   7560
      TabIndex        =   29
      Top             =   2160
      Width           =   600
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   210
      Left            =   9120
      TabIndex        =   28
      Top             =   1800
      Width           =   510
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sex:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   210
      Left            =   9120
      TabIndex        =   27
      Top             =   2160
      Width           =   345
   End
   Begin VB.Label Atxtaddress 
      BackStyle       =   0  'Transparent
      Caption         =   "----------"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   7560
      TabIndex        =   26
      Top             =   2520
      Width           =   1440
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Disease:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   210
      Left            =   6360
      TabIndex        =   25
      Top             =   3480
      Width           =   705
   End
   Begin VB.Label AtxtID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "----------"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   7560
      TabIndex        =   24
      Top             =   1800
      Width           =   600
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   210
      Left            =   6240
      TabIndex        =   23
      Top             =   2520
      Width           =   765
   End
   Begin VB.Label Atxtsex 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "----------"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   10320
      TabIndex        =   22
      Top             =   2160
      Width           =   600
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reappointment:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   210
      Left            =   6120
      TabIndex        =   21
      Top             =   3960
      Width           =   1305
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Checked by:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   210
      Left            =   9120
      TabIndex        =   20
      Top             =   2520
      Width           =   1020
   End
   Begin VB.Label AtxtCheckedBy 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "----------"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   10320
      TabIndex        =   19
      Top             =   2520
      Width           =   600
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fee:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   210
      Left            =   9120
      TabIndex        =   18
      Top             =   3960
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Patient ID:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   210
      Left            =   600
      TabIndex        =   13
      Top             =   2640
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   210
      Left            =   600
      TabIndex        =   12
      Top             =   3240
      Width           =   510
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Age:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   210
      Left            =   600
      TabIndex        =   11
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sex:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   210
      Left            =   600
      TabIndex        =   10
      Top             =   4440
      Width           =   345
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   210
      Left            =   600
      TabIndex        =   9
      Top             =   5040
      Width           =   765
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Refer To:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   210
      Left            =   600
      TabIndex        =   8
      Top             =   5760
      Width           =   750
   End
End
Attribute VB_Name = "frmpatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim clinic As Connection
Dim disease As Recordset
Dim patient As Recordset
Dim NoOfMed As Integer
Const MedFlexSize As Integer = 10
Public Sub HighLight()
With Screen.ActiveForm
 If (TypeOf .ActiveControl Is TextBox) Then
  .ActiveControl.SelStart = 0
  .ActiveControl.SelLength = Len(.ActiveControl)
 End If
End With
End Sub
Function IsValidDate(text As String) As Boolean
On Error GoTo Invalid
Dim d As Date
d = Format(text)
IsValidDate = True
Exit Function
Invalid:
IsValidDate = False
End Function

Private Sub Acmdadd_Click()
If NoOfMed >= MedFlexSize Then
 MsgBox "You can't add medicine more than " & MedFlexSize
 Exit Sub
End If
If Trim(Atxtmedicine.text) = "" Then
 MsgBox "Please,enter a medicine"
 Atxtmedicine.SetFocus
 Exit Sub
End If
If Trim(Atxtdose.text) = "" Then
 MsgBox "Please,enter a dose"
 Atxtdose.SetFocus
 Exit Sub
End If
NoOfMed = NoOfMed + 1
MedicineFlex.TextMatrix(NoOfMed, 0) = NoOfMed
MedicineFlex.TextMatrix(NoOfMed, 1) = Atxtmedicine.text
MedicineFlex.TextMatrix(NoOfMed, 2) = Atxtdose.text
Atxtmedicine.text = ""
Atxtdose.text = ""
Atxtmedicine.SetFocus
End Sub

Private Sub AcmdSavePrint_Click()
 Dim i As Integer
 If Not IsValidWholePatient Then
  HighLight
  Exit Sub
 End If
 If NoOfMed > 0 Then
  For i = 1 To NoOfMed
   disease.AddNew
    disease("id") = Trim(AtxtID.Caption)
    disease("checkeddate") = Date
    If Trim(Trim(Atxtreappointment.text)) <> "" Then
     disease("reappointment") = Trim(Atxtreappointment.text)
    End If
    disease("disease") = Trim(Atxtdesease.text)
    disease("madicine") = MedicineFlex.TextMatrix(i, 1)
    disease("dose") = MedicineFlex.TextMatrix(i, 2)
    disease("fee") = Trim(AtxtFee.text)
   disease.Update
  Next
 Else
   disease.AddNew
    disease("id") = Trim(AtxtID.Caption)
    disease("checkeddate") = Date
    If Trim(Trim(Atxtreappointment.text)) <> "" Then
     disease("reappointment") = Trim(Atxtreappointment.text)
    End If
    disease("disease") = Trim(Atxtdesease.text)
    disease("fee") = Trim(AtxtFee.text)
   disease.Update
 End If
 MsgBox "Record is successfully saved"
 'frmmain.Show
 Call showReport
 'Unload Me
End Sub
Function showReport()
MedicalReport.Sections.Item(1).Controls("txtdate").Caption _
 = Format(Date, "medium date")
MedicalReport.Sections.Item(1).Controls("txtid").Caption _
 = AtxtID.Caption
MedicalReport.Sections.Item(1).Controls("txtname").Caption _
 = Atxtname.Caption
MedicalReport.Sections.Item(1).Controls("txtage").Caption _
 = Atxtage.Caption & " year"
MedicalReport.Sections.Item(1).Controls("txtaddress").Caption _
 = Atxtaddress.Caption
MedicalReport.Sections.Item(1).Controls("txtfee").Caption _
 = AtxtFee.text & " rupees"
If Trim(Atxtreappointment.text) <> "" Then
 MedicalReport.Sections.Item(1).Controls("txtreappointment").Caption _
  = Format(Atxtreappointment.text, "medium date")
End If
MedicalReport.Sections.Item(1).Controls("txtsex").Caption _
 = Atxtsex.Caption
MedicalReport.Sections.Item(1).Controls("txtdisease").Caption _
 = Atxtdesease.text
MedicalReport.Sections.Item(1).Controls("txtcheckedby").Caption _
 = AtxtCheckedBy.Caption

'row0
MedicalReport.Sections.Item(1).Controls("a00").Caption _
 = MedicineFlex.TextMatrix(0, 0)
MedicalReport.Sections.Item(1).Controls("a01").Caption _
 = MedicineFlex.TextMatrix(0, 1)
MedicalReport.Sections.Item(1).Controls("a02").Caption _
 = MedicineFlex.TextMatrix(0, 2)
'row1
MedicalReport.Sections.Item(1).Controls("a10").Caption _
 = MedicineFlex.TextMatrix(1, 0)
MedicalReport.Sections.Item(1).Controls("a11").Caption _
 = MedicineFlex.TextMatrix(1, 1)
MedicalReport.Sections.Item(1).Controls("a12").Caption _
 = MedicineFlex.TextMatrix(1, 2)
'row2
MedicalReport.Sections.Item(1).Controls("a20").Caption _
 = MedicineFlex.TextMatrix(2, 0)
MedicalReport.Sections.Item(1).Controls("a21").Caption _
 = MedicineFlex.TextMatrix(2, 1)
MedicalReport.Sections.Item(1).Controls("a22").Caption _
 = MedicineFlex.TextMatrix(2, 2)
'row3
MedicalReport.Sections.Item(1).Controls("a30").Caption _
 = MedicineFlex.TextMatrix(3, 0)
MedicalReport.Sections.Item(1).Controls("a31").Caption _
 = MedicineFlex.TextMatrix(3, 1)
MedicalReport.Sections.Item(1).Controls("a32").Caption _
 = MedicineFlex.TextMatrix(3, 2)
'row4
MedicalReport.Sections.Item(1).Controls("a40").Caption _
 = MedicineFlex.TextMatrix(4, 0)
MedicalReport.Sections.Item(1).Controls("a41").Caption _
 = MedicineFlex.TextMatrix(4, 1)
MedicalReport.Sections.Item(1).Controls("a42").Caption _
 = MedicineFlex.TextMatrix(4, 2)
'row5
MedicalReport.Sections.Item(1).Controls("a50").Caption _
 = MedicineFlex.TextMatrix(5, 0)
MedicalReport.Sections.Item(1).Controls("a51").Caption _
 = MedicineFlex.TextMatrix(5, 1)
MedicalReport.Sections.Item(1).Controls("a52").Caption _
 = MedicineFlex.TextMatrix(5, 2)
'row6
MedicalReport.Sections.Item(1).Controls("a60").Caption _
 = MedicineFlex.TextMatrix(6, 0)
MedicalReport.Sections.Item(1).Controls("a61").Caption _
 = MedicineFlex.TextMatrix(6, 1)
MedicalReport.Sections.Item(1).Controls("a62").Caption _
 = MedicineFlex.TextMatrix(6, 2)
'row7
MedicalReport.Sections.Item(1).Controls("a70").Caption _
 = MedicineFlex.TextMatrix(7, 0)
MedicalReport.Sections.Item(1).Controls("a71").Caption _
 = MedicineFlex.TextMatrix(7, 1)
MedicalReport.Sections.Item(1).Controls("a72").Caption _
 = MedicineFlex.TextMatrix(7, 2)
'row8
MedicalReport.Sections.Item(1).Controls("a80").Caption _
 = MedicineFlex.TextMatrix(8, 0)
MedicalReport.Sections.Item(1).Controls("a81").Caption _
 = MedicineFlex.TextMatrix(8, 1)
MedicalReport.Sections.Item(1).Controls("a82").Caption _
 = MedicineFlex.TextMatrix(8, 2)
'row9
MedicalReport.Sections.Item(1).Controls("a90").Caption _
 = MedicineFlex.TextMatrix(9, 0)
MedicalReport.Sections.Item(1).Controls("a91").Caption _
 = MedicineFlex.TextMatrix(9, 1)
MedicalReport.Sections.Item(1).Controls("a92").Caption _
 = MedicineFlex.TextMatrix(9, 2)
'row10
MedicalReport.Sections.Item(1).Controls("a100").Caption _
 = MedicineFlex.TextMatrix(10, 0)
MedicalReport.Sections.Item(1).Controls("a101").Caption _
 = MedicineFlex.TextMatrix(10, 1)
MedicalReport.Sections.Item(1).Controls("a102").Caption _
 = MedicineFlex.TextMatrix(10, 2)
 
 MedicalReport.Refresh
 MedicalReport.Show
End Function
Function IsValidWholePatient() As Boolean
If Trim(Atxtdesease.text) = "" Then
 MsgBox "Please,enter an disease of patient"
 Atxtdesease.SetFocus
 Exit Function
End If
If Trim(Atxtreappointment.text) <> "" Then
  If IsNumeric(Trim(Atxtreappointment.text)) Or Not IsValidDate(Trim(Atxtreappointment.text)) Then
  Atxtreappointment.SetFocus
  MsgBox "Not valid Date,Try e.g: " & Format(Now, "Medium Date")
   Exit Function
 End If
End If
If Trim(AtxtFee.text) = "" Then
 MsgBox "Please,enter fee taken by patient"
 AtxtFee.SetFocus
 Exit Function
End If
If Not IsNumeric(Trim(AtxtFee.text)) Then
 MsgBox "Please,enter numeric value for fee"
 AtxtFee.SetFocus
 Exit Function
End If
If Trim(AtxtFee.text) < 0 Then
 MsgBox "Negative fee not allowed"
 AtxtFee.SetFocus
 Exit Function
End If
If NoOfMed < 1 Then
 If MsgBox("You don't like to add any medicine?", vbYesNo, "No medicine added") = vbNo Then
  Exit Function
 End If
End If
 IsValidWholePatient = True
End Function
Private Sub cmdSaveShift_Click()
If Not IsExistPatient Then
 If Not SaveNewPatient Then
  Exit Sub
 End If
End If
'Frame2.Enabled = True
Call BeforeShiftAfter
Call ClearPatientRecord
Call LoadPatientId
Call EnableAfterCheck
Atxtdesease.SetFocus
End Sub
Function IsExistPatient() As Boolean
On Error GoTo err
If patient.EOF And patient.BOF Then
 Exit Function
Else
 patient.MoveFirst
End If

While Not patient.EOF
  Dim id As String
  id = patient("id")
  If Trim(patient("id")) = Trim(BtxtID.text) Then
  'If id = Trim(BtxtID.Text) Then
   Call LoadExistPatientRecord
   IsExistPatient = True
   Exit Function
  End If
 patient.MoveNext
Wend
'Call ClearPatientRecord
Exit Function
err:
'MsgBox err.Description

End Function
Function BeforeShiftAfter()
 AtxtID = BtxtID
 Atxtname = Btxtname
 Atxtage = Btxtage
 Atxtsex = Btxtsex
 Atxtaddress = Btxtaddress
 AtxtID = BtxtID
 AtxtCheckedBy = Btxtreferto
Atxtdesease.text = ""
AtxtFee.text = ""
Atxtreappointment.text = ""
Atxtmedicine.text = ""
Atxtdose.text = ""
Call ClearMedicineFlex
End Function
Function ClearMedicineFlex()
Dim i As Integer
For i = 1 To NoOfMed
 MedicineFlex.TextMatrix(i, 0) = ""
 MedicineFlex.TextMatrix(i, 1) = ""
 MedicineFlex.TextMatrix(i, 2) = ""
Next
NoOfMed = 0
End Function

Private Sub Command1_Click()
MsgBox "Heaven Soft Company Mirpur (A.K) Pakistan,Email:nishat_kazmi@ yahoo.com"
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command4_Click()
frmReport.Show
End Sub

Private Sub Form_Load()
  Set clinic = New Connection
  clinic.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=clinic.mdb;Persist Security Info=False"
  Set disease = New Recordset
  disease.Open "select * from disease", clinic, adOpenStatic, adLockOptimistic
  Set patient = New Recordset
  patient.Open "select * from patient order by id", clinic, adOpenStatic, adLockOptimistic
 Call LoadPatientId
 Btxtsex.text = "Male"
 Call FlexClear
 Call DisableAfterCheck
End Sub
Function DisableAfterCheck()
 Atxtdesease.Enabled = False
 Atxtreappointment.Enabled = False
 AtxtFee.Enabled = False
 Atxtmedicine.Enabled = False
 Atxtdose.Enabled = False
 Acmdadd.Enabled = False
 AcmdSavePrint.Enabled = False
End Function
Function EnableAfterCheck()
 Atxtdesease.Enabled = True
 Atxtreappointment.Enabled = True
 AtxtFee.Enabled = True
 Atxtmedicine.Enabled = True
 Atxtdose.Enabled = True
 Acmdadd.Enabled = True
 AcmdSavePrint.Enabled = True
End Function
Function FlexClear()
 MedicineFlex.TextMatrix(0, 0) = "Sno"
 MedicineFlex.TextMatrix(0, 1) = "Medicine"
 MedicineFlex.TextMatrix(0, 2) = "Dose"

Dim i As Integer
For i = 1 To MedFlexSize
 MedicineFlex.TextMatrix(i, 0) = ""
 MedicineFlex.TextMatrix(i, 1) = ""
 MedicineFlex.TextMatrix(i, 2) = ""
Next
End Function

Private Sub BcmdSave_Click()
 If Not SaveNewPatient Then
  Exit Sub
 End If
End Sub
Function IsIDexist() As Boolean
On Error GoTo err
If patient.EOF And patient.BOF Then
 Exit Function
Else
 patient.MoveFirst
End If

While Not patient.EOF
  'Dim id As String
  'id = patient("id")
  If Trim(patient("id")) = Trim(BtxtID.text) Then
  'If id = Trim(BtxtID.Text) Then
   'Call LoadExistPatientRecord
   IsIDexist = True
   Exit Function
  End If
 patient.MoveNext
Wend
'Call ClearPatientRecord
Exit Function
err:
'MsgBox err.Description

End Function
Function SaveNewPatient() As Boolean
On Error GoTo err
If Not IsValidPatient Then
 Call HighLight
 Exit Function
End If
If IsIDexist Then
 MsgBox "This ID is already exist try other one"
 BtxtID.SetFocus
 Call HighLight
 Exit Function
End If
patient.AddNew
 patient("id") = Trim(BtxtID.text)
 patient("name") = Trim(Btxtname.text)
 patient("age") = Trim(Btxtage.text)
 patient("sex") = Trim(Btxtsex.text)
 patient("address") = Trim(Btxtaddress.text)
 patient("referto") = Trim(Btxtreferto.text)
patient.Update
MsgBox "New Patient's Record Is Successfully Saved"
SaveNewPatient = True
Exit Function
err:
 MsgBox err.Description
End Function
Function IsValidPatient() As Boolean
If Trim(BtxtID.text) = "" Then
 MsgBox "Please,enter an ID for patient"
 BtxtID.SetFocus
 Exit Function
End If
If Not IsNumeric(Trim(BtxtID.text)) Then
 MsgBox "Please,enter numeric value for ID"
 BtxtID.SetFocus
 Exit Function
End If
If Trim(BtxtID.text) < 0 Then
 MsgBox "Negative ID not allowed"
 BtxtID.SetFocus
 Exit Function
End If
If Trim(Btxtname.text) = "" Then
 MsgBox "Please,enter name of patient"
 Btxtname.SetFocus
 Exit Function
End If
If Trim(Btxtage.text) = "" Then
 MsgBox "Please,enter age of patient"
 Btxtage.SetFocus
 Exit Function
End If
If Not IsNumeric(Trim(Btxtage.text)) Then
 MsgBox "Please,enter numeric value for age"
 Btxtage.SetFocus
 Exit Function
End If
If Trim(Btxtage.text) < 1 Then
 MsgBox "Negative or 0 age not allowed"
 Btxtage.SetFocus
 Exit Function
End If
If Trim(Btxtreferto.text) = "" Then
 MsgBox "Please,enter name of doctor for refer"
 Btxtreferto.SetFocus
 Exit Function
End If

IsValidPatient = True
End Function
Private Sub BtxtID_Change()
On Error GoTo err
If patient.EOF And patient.BOF Then
 Exit Sub
Else
 patient.MoveFirst
End If

While Not patient.EOF
  Dim id As String
  id = patient("id")
  If Trim(patient("id")) = Trim(BtxtID.text) Then
  'If id = Trim(BtxtID.Text) Then
   Call LoadExistPatientRecord
   Exit Sub
  End If
  'MsgBox patient("id") & "  " & Trim(BtxtID.Text)
 patient.MoveNext
Wend
Call ClearPatientRecord
Exit Sub
err:
MsgBox err.Description
End Sub
Function LoadExistPatientRecord()
On Error GoTo err
 Btxtname.text = patient("name")
 Btxtage.text = patient("age")
 Btxtsex.text = patient("sex")
 Btxtaddress.text = patient("address")
 Btxtreferto.text = patient("referto")
Exit Function
err:
End Function
Function ClearPatientRecord()
 Btxtname.text = ""
 Btxtage.text = ""
 Btxtsex.text = "Male"
 Btxtaddress.text = ""
 Btxtreferto.text = ""
End Function
Function LoadPatientId()
 If patient.BOF And patient.EOF Then
  BtxtID.text = 1
  Exit Function
 End If
 patient.MoveLast
 BtxtID.text = patient("id") + 1
End Function



