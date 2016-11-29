VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmOtherServices 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6555
   ClientLeft      =   3105
   ClientTop       =   1575
   ClientWidth     =   9135
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   9135
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frEnPrint 
      Caption         =   "Encode And Print"
      Height          =   2055
      Left            =   360
      TabIndex        =   23
      Top             =   3240
      Width           =   6255
   End
   Begin VB.Frame frScan 
      Caption         =   "Pic Scan"
      Height          =   2175
      Left            =   7200
      TabIndex        =   24
      Top             =   0
      Width           =   6015
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         TabIndex        =   60
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "Scan Sets:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   59
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2520
         TabIndex        =   58
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame frPrint 
      Caption         =   "Print"
      Height          =   2175
      Left            =   7200
      TabIndex        =   20
      Top             =   2880
      Width           =   6615
      Begin VB.Timer tmrPrint 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   120
         Top             =   1680
      End
      Begin VB.TextBox txtPrintAdtlAmt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5640
         MaxLength       =   4
         TabIndex        =   41
         Top             =   1320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox chkPrintGraphics 
         BackColor       =   &H80000004&
         Caption         =   "w/ Graphics"
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
         Left            =   600
         TabIndex        =   39
         Top             =   1370
         Width           =   1695
      End
      Begin VB.TextBox txtPrintQtyL 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3000
         MaxLength       =   4
         TabIndex        =   32
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtPrintQtyS 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3000
         MaxLength       =   4
         TabIndex        =   31
         Top             =   960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox chkPrintShort 
         BackColor       =   &H80000004&
         Caption         =   "Short 8 1/2 x 11"
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
         Left            =   600
         TabIndex        =   26
         Top             =   1000
         Width           =   1695
      End
      Begin VB.CheckBox chkPrintLong 
         BackColor       =   &H80000004&
         Caption         =   "Long 8 1/2 x 13"
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
         Left            =   600
         TabIndex        =   25
         Top             =   650
         Width           =   1695
      End
      Begin VB.Label lblPrintRateS 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   43
         Top             =   960
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblPrintRateL 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   42
         Top             =   600
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "Addt'l:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   4440
         TabIndex        =   40
         Top             =   1350
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "Rate/page:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   3960
         TabIndex        =   38
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblPrintTAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5280
         TabIndex        =   37
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4680
         TabIndex        =   36
         Top             =   1800
         Width           =   495
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   5280
         X2              =   6360
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label lblPrintAmtS 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   35
         Top             =   960
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblPrintAmtL 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   34
         Top             =   600
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5520
         TabIndex        =   33
         Top             =   240
         Width           =   855
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   2520
         X2              =   2640
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   2520
         X2              =   2640
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   2760
         TabIndex        =   28
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "Size/Addt'l:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2655
      Left            =   240
      TabIndex        =   21
      Top             =   2880
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   4683
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "PRINT"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ENCODE"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ENCODE && PRINT"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "PIC SCAN"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cl&ear"
      Height          =   375
      Left            =   1560
      TabIndex        =   30
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "A&dd"
      Height          =   375
      Left            =   240
      TabIndex        =   29
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H0080C0FF&
      Caption         =   "ENCODE && PRINT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   720
      TabIndex        =   18
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox txtPrnQnty 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5640
      TabIndex        =   13
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      TabIndex        =   12
      Top             =   1440
      Width           =   615
   End
   Begin VB.CheckBox chkPrint 
      BackColor       =   &H0080C0FF&
      Caption         =   "PRINT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   720
      TabIndex        =   11
      Top             =   960
      Width           =   1695
   End
   Begin VB.CheckBox chkScan 
      BackColor       =   &H0080C0FF&
      Caption         =   "PIC SCAN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   720
      TabIndex        =   10
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CheckBox chkEncode 
      BackColor       =   &H0080C0FF&
      Caption         =   "ENCODE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   720
      TabIndex        =   9
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      TabIndex        =   8
      Top             =   1800
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Short 8 1/2 x 11"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4800
      TabIndex        =   7
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Long 8 1/2 x 13"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4800
      TabIndex        =   6
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2520
      TabIndex        =   5
      Text            =   "2000.00"
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "C&ancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Appearance      =   0  'Flat
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "X"
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   375
   End
   Begin VB.Frame frEncode 
      Caption         =   "Encode"
      Height          =   1815
      Left            =   6960
      TabIndex        =   22
      Top             =   1080
      Width           =   6735
      Begin VB.Timer tmrEncode 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   120
         Top             =   1320
      End
      Begin VB.CheckBox chkEncodeL 
         BackColor       =   &H80000004&
         Caption         =   "Long 8 1/2 x 13"
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
         Left            =   600
         TabIndex        =   47
         Top             =   645
         Width           =   1695
      End
      Begin VB.CheckBox chkEncodeS 
         BackColor       =   &H80000004&
         Caption         =   "Short 8 1/2 x 11"
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
         Left            =   600
         TabIndex        =   46
         Top             =   1005
         Width           =   1695
      End
      Begin VB.TextBox txtEncQtyS 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3000
         MaxLength       =   4
         TabIndex        =   45
         Top             =   960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtEncQtyL 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3000
         MaxLength       =   4
         TabIndex        =   44
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label18 
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "Size:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   57
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   2760
         TabIndex        =   56
         Top             =   240
         Width           =   855
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   2520
         X2              =   2640
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line5 
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   2520
         X2              =   2640
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label17 
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5520
         TabIndex        =   55
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblEncAmtL 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   54
         Top             =   600
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblEncAmtS 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   53
         Top             =   960
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         X1              =   5280
         X2              =   6360
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label14 
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4680
         TabIndex        =   52
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblEncTAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5280
         TabIndex        =   51
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "Rate/page:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   3960
         TabIndex        =   50
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblEncRateL 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   49
         Top             =   600
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblEncRateS 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   48
         Top             =   960
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000000&
      BackStyle       =   1  'Opaque
      Height          =   1935
      Left            =   8040
      Top             =   2280
      Width           =   5655
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Service Type:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Qnty:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5160
      TabIndex        =   17
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Qnty:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2880
      TabIndex        =   16
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Qnty:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2880
      TabIndex        =   15
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Size:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4200
      TabIndex        =   14
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblType 
      BackColor       =   &H00000000&
      Caption         =   " Other Services"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5655
   End
End
Attribute VB_Name = "frmOtherServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkEncodeL_Click()
  If chkEncodeL.Value = vbChecked Then
    txtEncQtyL.Visible = True
    lblEncRateL.Visible = True
    lblEncAmtL.Visible = True
  Else
    txtEncQtyL.Visible = False
    lblEncRateL.Visible = False
    lblEncAmtL.Visible = False
  End If
End Sub

Private Sub chkEncodeS_Click()
  If chkEncodeS.Value = vbChecked Then
    txtEncQtyS.Visible = True
    lblEncRateS.Visible = True
    lblEncAmtS.Visible = True
  Else
    txtEncQtyS.Visible = False
    lblEncRateS.Visible = False
    lblEncAmtS.Visible = False
  End If
End Sub

Private Sub chkPrintGraphics_Click()
  If chkPrintGraphics.Value = vbChecked Then
    Label8(2).Visible = True
    txtPrintAdtlAmt.Visible = True
  Else
    Label8(2).Visible = False
    txtPrintAdtlAmt.Visible = False
  End If
End Sub

Private Sub chkPrintLong_Click()
  If chkPrintLong.Value = vbChecked Then
    Line1.Visible = True
    txtPrintQtyL.Visible = True
    lblPrintAmtL.Visible = True
    lblPrintRateL.Visible = True
  Else
    Line1.Visible = False
    txtPrintQtyL.Visible = False
    lblPrintAmtL.Visible = False
    lblPrintRateL.Visible = False
  End If
End Sub

Private Sub chkPrintShort_Click()
  If chkPrintShort.Value = vbChecked Then
    Line2.Visible = True
    txtPrintQtyS.Visible = True
    lblPrintAmtS.Visible = True
    lblPrintRateS.Visible = True
  Else
    Line2.Visible = False
    txtPrintQtyS.Visible = False
    lblPrintAmtS.Visible = False
    lblPrintRateS.Visible = False
  End If
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub lblType_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  MoveForm Me
End Sub

Private Sub tab1_Click()
Select Case True
  Case tab1.Tabs(1).Selected = True
    MsgBox "1"
  Case tab1.Tabs(2).Selected = True
    MsgBox "2"
End Select
End Sub

Private Sub tmrEncode_Timer()
Dim AmtL As Double
Dim AmtS As Double
  AmtL = Val(Trim(txtEncQtyL.Text)) * Val(Trim(lblEncRateL.Caption))
  Me.lblEncAmtL.Caption = FormatNumber(AmtL, 2)
  AmtS = Val(Trim(txtEncQtyS.Text)) * Val(Trim(lblEncRateS.Caption))
  Me.lblEncAmtS.Caption = FormatNumber(AmtS, 2)
  Me.lblEncTAmt.Caption = FormatNumber(AmtL + AmtS, 2)
End Sub

Private Sub tmrPrint_Timer()
Dim AmtL As Double
Dim AmtS As Double
  AmtL = Val(Trim(txtPrintQtyL.Text)) * Val(Trim(lblPrintRateL.Caption))
  Me.lblPrintAmtL.Caption = FormatNumber(AmtL, 2)
  AmtS = Val(Trim(txtPrintQtyS.Text)) * Val(Trim(lblPrintRateS.Caption))
  Me.lblPrintAmtS.Caption = FormatNumber(AmtS, 2)
  Me.lblPrintTAmt.Caption = FormatNumber(AmtL + AmtS + Val(Trim(txtPrintAdtlAmt.Text)), 2)
End Sub
