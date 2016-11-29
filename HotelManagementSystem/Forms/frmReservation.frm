VERSION 5.00
Object = "{9EDDC69F-10E8-4DE7-9648-EC8A45F005C0}#1.0#0"; "b8Controls4.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmReservation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reservation"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11265
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   11265
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox bgMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E3F9FB&
      Height          =   7005
      Left            =   0
      ScaleHeight     =   467
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   775
      TabIndex        =   4
      Top             =   540
      Width           =   11625
      Begin VB.CommandButton CmdCheckIn 
         Caption         =   "Check In"
         Height          =   315
         Left            =   6780
         TabIndex        =   79
         Top             =   6150
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "Cancel Reservation"
         Height          =   315
         Left            =   6780
         TabIndex        =   78
         Top             =   6540
         Visible         =   0   'False
         Width           =   2685
      End
      Begin VB.HScrollBar hsAdults 
         Height          =   315
         Left            =   6330
         Min             =   1
         TabIndex        =   33
         Top             =   2340
         Value           =   1
         Width           =   495
      End
      Begin VB.HScrollBar hsChildrens 
         Height          =   315
         Left            =   6330
         TabIndex        =   32
         Top             =   2760
         Width           =   495
      End
      Begin VB.HScrollBar hsDays 
         Height          =   315
         Left            =   6330
         Min             =   1
         TabIndex        =   31
         Top             =   1920
         Value           =   1
         Width           =   495
      End
      Begin VB.TextBox txtBalance 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   9090
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "0.00"
         Top             =   4440
         Width           =   1815
      End
      Begin VB.TextBox txtAmountPaid 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   9090
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "0.00"
         Top             =   4050
         Width           =   1815
      End
      Begin VB.TextBox txtTotalCharges 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   9090
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "0.00"
         Top             =   1920
         Width           =   1815
      End
      Begin VB.CommandButton cmdUsrHistory 
         Caption         =   "Modification History"
         Height          =   315
         Left            =   420
         TabIndex        =   27
         Top             =   6180
         Width           =   1680
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   315
         Left            =   9480
         TabIndex        =   26
         Top             =   6540
         Width           =   1335
      End
      Begin VB.TextBox txtChildrens 
         Alignment       =   2  'Center
         Height          =   345
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "0"
         Top             =   2760
         Width           =   465
      End
      Begin VB.TextBox txtAdults 
         Alignment       =   2  'Center
         Height          =   345
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "1"
         Top             =   2340
         Width           =   465
      End
      Begin VB.TextBox txtDays 
         Alignment       =   2  'Center
         Height          =   345
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "1"
         Top             =   1920
         Width           =   465
      End
      Begin VB.TextBox txtDiscount 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   9090
         TabIndex        =   22
         Text            =   "0.00"
         Top             =   3180
         Width           =   1575
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   9090
         TabIndex        =   21
         Text            =   "0.00"
         Top             =   1500
         Width           =   1815
      End
      Begin VB.TextBox txtIDNumber 
         Height          =   345
         Left            =   1680
         TabIndex        =   20
         Top             =   4200
         Width           =   1815
      End
      Begin VB.TextBox txtCompany 
         Height          =   345
         Left            =   1680
         TabIndex        =   19
         Top             =   2820
         Width           =   1815
      End
      Begin VB.TextBox txtAddress 
         Height          =   345
         Left            =   1680
         TabIndex        =   18
         Top             =   1950
         Width           =   1815
      End
      Begin VB.TextBox txtFirstName 
         Height          =   345
         Left            =   1680
         TabIndex        =   17
         Top             =   1500
         Width           =   1815
      End
      Begin VB.TextBox txtLastName 
         Height          =   345
         Left            =   1680
         TabIndex        =   16
         Top             =   1050
         Width           =   1815
      End
      Begin VB.TextBox txtReservationNo 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   1680
         TabIndex        =   15
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   315
         Left            =   9480
         TabIndex        =   14
         Top             =   6150
         Width           =   1335
      End
      Begin VB.TextBox txtOtherCharges 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   9090
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "0.00"
         Top             =   2340
         Width           =   1815
      End
      Begin VB.TextBox txtSubTotal 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   9090
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "0.00"
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   9090
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "0.00"
         Top             =   3600
         Width           =   1815
      End
      Begin VB.TextBox txtVehicleModel 
         Height          =   345
         Left            =   1650
         TabIndex        =   10
         Top             =   5340
         Width           =   1815
      End
      Begin VB.TextBox txtPlateNo 
         Height          =   345
         Left            =   1650
         TabIndex        =   9
         Top             =   5760
         Width           =   1815
      End
      Begin VB.TextBox txtNotes 
         Height          =   1485
         Left            =   4530
         TabIndex        =   8
         Top             =   4500
         Width           =   3045
      End
      Begin VB.CommandButton cmdLookupCust 
         Caption         =   ">>"
         Height          =   315
         Left            =   3540
         TabIndex        =   7
         Top             =   1050
         Width           =   375
      End
      Begin VB.CommandButton cmdLookupComp 
         Caption         =   ">>"
         Height          =   315
         Left            =   3540
         TabIndex        =   6
         Top             =   2820
         Width           =   375
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
         Height          =   315
         Left            =   8130
         TabIndex        =   5
         Top             =   6150
         Visible         =   0   'False
         Width           =   1335
      End
      Begin b8Controls4.b8Line b8Line2 
         Height          =   30
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   53
         BorderColor1    =   15592425
         BorderColor2    =   16777215
      End
      Begin b8Controls4.b8Line b8Line5 
         Height          =   30
         Left            =   240
         TabIndex        =   35
         Top             =   6060
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   53
         BorderColor1    =   15592425
         BorderColor2    =   16777215
      End
      Begin MSDataListLib.DataCombo dcCountry 
         Height          =   315
         Left            =   1680
         TabIndex        =   36
         Top             =   2400
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker dtpDateIn 
         Height          =   345
         Left            =   5760
         TabIndex        =   37
         Top             =   1080
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   609
         _Version        =   393216
         Format          =   77266945
         CurrentDate     =   39536
      End
      Begin b8Controls4.b8GradLine b8GradLine1 
         Height          =   240
         Left            =   180
         TabIndex        =   38
         Top             =   240
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   423
         Color1          =   14737632
         Color2          =   16119285
         Caption         =   "Guest Information"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   4210752
      End
      Begin b8Controls4.b8GradLine b8GradLine3 
         Height          =   240
         Left            =   210
         TabIndex        =   39
         Top             =   3330
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   423
         Color1          =   14737632
         Color2          =   16119285
         Caption         =   "Identification Information"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   4210752
      End
      Begin b8Controls4.b8GradLine b8GradLine4 
         Height          =   240
         Left            =   4050
         TabIndex        =   40
         Top             =   240
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   423
         Color1          =   14737632
         Color2          =   16119285
         Caption         =   "Rate Information"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   4210752
      End
      Begin MSComCtl2.DTPicker dtpDateOut 
         Height          =   345
         Left            =   9090
         TabIndex        =   41
         Top             =   1050
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   609
         _Version        =   393216
         Format          =   77266945
         CurrentDate     =   39536
      End
      Begin MSDataListLib.DataCombo dcIDType 
         Height          =   315
         Left            =   1680
         TabIndex        =   42
         Top             =   3780
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcRateType 
         Height          =   315
         Left            =   5760
         TabIndex        =   43
         Top             =   1530
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcBusSource 
         Height          =   315
         Left            =   5760
         TabIndex        =   44
         Top             =   3210
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin b8Controls4.b8GradLine b8GradLine2 
         Height          =   240
         Left            =   240
         TabIndex        =   45
         Top             =   4650
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   423
         Color1          =   14737632
         Color2          =   16119285
         Caption         =   "Vehicle Information"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   4210752
      End
      Begin MSDataListLib.DataCombo dcVehicle 
         Height          =   315
         Left            =   1650
         TabIndex        =   46
         Top             =   4920
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcRoomNumber 
         Height          =   315
         Left            =   5760
         TabIndex        =   77
         Top             =   660
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label lblRM 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   9450
         TabIndex        =   76
         Top             =   3030
         Width           =   45
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Charges"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7650
         TabIndex        =   75
         Top             =   1950
         Width           =   1395
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Business Source"
         Height          =   300
         Left            =   4320
         TabIndex        =   74
         Top             =   3210
         Width           =   1395
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. of Childrens"
         Height          =   300
         Left            =   4320
         TabIndex        =   73
         Top             =   2790
         Width           =   1395
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. of Adults"
         Height          =   300
         Left            =   4320
         TabIndex        =   72
         Top             =   2370
         Width           =   1395
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. of Days"
         Height          =   300
         Left            =   4320
         TabIndex        =   71
         Top             =   1950
         Width           =   1395
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
         Height          =   300
         Left            =   7650
         TabIndex        =   70
         Top             =   3210
         Width           =   1395
      End
      Begin VB.Label lblRatePerPeriod 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Rate/Period"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   7650
         TabIndex        =   69
         Top             =   1500
         Width           =   1395
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Rate Type"
         Height          =   300
         Left            =   4320
         TabIndex        =   68
         Top             =   1530
         Width           =   1395
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Date Out"
         Height          =   300
         Left            =   7650
         TabIndex        =   67
         Top             =   1080
         Width           =   1395
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Date In"
         Height          =   300
         Left            =   4320
         TabIndex        =   66
         Top             =   1110
         Width           =   1395
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Room Number"
         Height          =   300
         Left            =   4320
         TabIndex        =   65
         Top             =   660
         Width           =   1395
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ID Type"
         Height          =   300
         Left            =   240
         TabIndex        =   64
         Top             =   3780
         Width           =   1395
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ID Number"
         Height          =   300
         Left            =   240
         TabIndex        =   63
         Top             =   4200
         Width           =   1395
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Country"
         Height          =   300
         Left            =   240
         TabIndex        =   62
         Top             =   2400
         Width           =   1395
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Company"
         Height          =   300
         Left            =   240
         TabIndex        =   61
         Top             =   2850
         Width           =   1395
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   300
         Left            =   240
         TabIndex        =   60
         Top             =   1980
         Width           =   1395
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "First name"
         Height          =   300
         Left            =   240
         TabIndex        =   59
         Top             =   1530
         Width           =   1395
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Last name"
         Height          =   300
         Left            =   240
         TabIndex        =   58
         Top             =   1080
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Reservation No"
         Height          =   300
         Left            =   240
         TabIndex        =   57
         Top             =   630
         Width           =   1395
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   315
         Left            =   10680
         TabIndex        =   56
         Top             =   3180
         Width           =   165
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  'Dot
         X1              =   516
         X2              =   516
         Y1              =   44
         Y2              =   274
      End
      Begin VB.Label lblOtherCharges 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Other Charges"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   7650
         TabIndex        =   55
         Top             =   2370
         Width           =   1395
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7650
         TabIndex        =   54
         Top             =   2790
         Width           =   1395
      End
      Begin VB.Label lblAmountPaid 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Paid"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   7650
         TabIndex        =   53
         Top             =   4080
         Width           =   1395
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Balance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7650
         TabIndex        =   52
         Top             =   4470
         Width           =   1395
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7650
         TabIndex        =   51
         Top             =   3630
         Width           =   1395
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         Height          =   300
         Left            =   210
         TabIndex        =   50
         Top             =   5340
         Width           =   1395
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicle Make"
         Height          =   300
         Left            =   210
         TabIndex        =   49
         Top             =   4920
         Width           =   1395
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Plate No."
         Height          =   300
         Left            =   210
         TabIndex        =   48
         Top             =   5760
         Width           =   1395
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Notes"
         Height          =   285
         Left            =   4530
         TabIndex        =   47
         Top             =   4230
         Width           =   585
      End
   End
   Begin VB.PictureBox bgHeader 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   30
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   687
      TabIndex        =   1
      Top             =   0
      Width           =   10305
      Begin VB.Image Image1 
         Height          =   480
         Left            =   30
         Picture         =   "frmReservation.frx":0000
         Top             =   30
         Width           =   480
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reservation"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00926747&
         Height          =   345
         Left            =   600
         TabIndex        =   3
         Top             =   30
         Width           =   1710
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fill all fields or fields with '*' then click 'Save' button to update."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00926747&
         Height          =   180
         Left            =   600
         TabIndex        =   2
         Top             =   360
         Width           =   3900
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7470
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6588
            Text            =   "Reserved By:"
            TextSave        =   "Reserved By:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6588
            Text            =   "Cancelled By:"
            TextSave        =   "Cancelled By:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6588
            Text            =   "Business Source:"
            TextSave        =   "Business Source:"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmReservation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public State                As FormState 'Variable used to determine on how the form used
Public PK                   As String 'Variable used to get what record is going to edit
Public Shortcut             As Boolean 'Determine if this form is open from the shortcut link or in the list

Dim HaveAction              As Boolean 'Variable used to detect if the user perform some action
Dim RS                      As New Recordset

Private Sub DisplayForEditing()
On Error GoTo err
    
    With RS
        txtReservationNo.Text = .Fields("ReservationNo")
        txtLastName.Text = getValueAt("SELECT LastName FROM Customers WHERE CustomerID = " & RS.Fields("CustomerID"), "LastName")
        txtFirstName.Text = getValueAt("SELECT FirstName FROM Customers WHERE CustomerID = " & RS.Fields("CustomerID"), "FirstName")
        txtAddress.Text = .Fields("Address")
        dcCountry.BoundText = .Fields("CountryID")
        If RS.Fields("CompanyID") <> "" Then _
            txtCompany.Text = getValueAt("SELECT Company FROM Company WHERE CompanyID = " & RS.Fields("CompanyID"), "Company")
        dcIDType.BoundText = .Fields("IDTypeID")
        txtIDNumber.Text = .Fields("IDNumber")
        dcRoomNumber.BoundText = .Fields("RoomNumber")
        dtpDateIn.Value = .Fields("DateIn")
        If .Fields("DateOut") >= Date Then
            dtpDateOut.Value = .Fields("DateOut")
        ElseIf .Fields("DateIn") = Date Then
            dtpDateOut.Value = dtpDateIn.Value + 1
        Else
            dtpDateOut.Value = Date
        End If
        dcRateType.BoundText = .Fields("RateType")
        txtRate.Text = toMoney(.Fields("Rate"))
        txtOtherCharges.Text = toMoney(.Fields("OtherCharges"))
        txtDiscount.Text = .Fields("Discount")
        txtAmountPaid.Text = toMoney(.Fields("AmountPaid"))
        txtDays.Text = dtpDateOut.Value - dtpDateIn.Value '.Fields("Days")
        txtAdults.Text = .Fields("Adults")
        txtChildrens.Text = .Fields("Childrens")
        dcBusSource.BoundText = .Fields("BusinessSourceID")
        dcVehicle.BoundText = .Fields("VehicleID")
        txtVehicleModel.Text = .Fields("VehicleModel")
        txtPlateNo.Text = .Fields("PlateNo")
        txtNotes.Text = .Fields("Notes")
    End With
    
    hsDays.Value = txtDays.Text
    hsAdults.Value = txtAdults.Text
    hsChildrens.Value = txtChildrens.Text
    
    StatusBar1.Panels(1).Text = "Reserved By: " & getValueAt("SELECT UserID FROM Users WHERE PK = " & RS.Fields("ReservedBy"), "UserID")
    StatusBar1.Panels(2).Text = "Cancelled By: " & getValueAt("SELECT UserID FROM Users WHERE PK = " & RS.Fields("CancelledBy"), "UserID")
    StatusBar1.Panels(3).Text = "Business Source: " & dcBusSource.Text
    
    Exit Sub
    
err:
    If err.Number = 94 Then Resume Next
    
    prompt_err err, Name, "DisplayForEditing"
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdCancel_Click()
    If MsgBox("Are you sure you want to cancel this reservation?", vbYesNo + vbExclamation) = vbYes Then
        ChangeValue CN, "Reservation", "CancelledBy", CurrUser.USER_PK, True, "WHERE ReservationNo = '" & txtReservationNo.Text & "'"
    
        Unload Me
    End If
End Sub

Private Sub CmdCheckIn_Click()
On Error GoTo err
    Dim intCount As Integer
    
    If dcRoomNumber.BoundText = "" Then
        MsgBox "Please select Room number.", vbCritical
        
        Exit Sub
    End If
    
    Call Save_Record
    
    If HaveAction = False Then Exit Sub
    
    If MsgBox("Are you sure?", vbYesNo + vbInformation) = vbYes Then
        Dim FolioNumber As String
        
        FolioNumber = getIndex("Transactions")
        FolioNumber = GenerateID(FolioNumber, Format$(Date, "yy") & "-", "00000")

        CN.BeginTrans
        
        CN.Execute "INSERT INTO Transactions ( FolioNumber, CustomerID, Address, CountryID, CompanyID, IDTypeID, IDNumber, RoomNumber, DateIn, DateOut, RateType, Rate, Days, Adults, Childrens, BusinessSourceID, VehicleID, VehicleModel, PlateNo, Status, Notes, ReservedBy, CheckInBy, AddedByFK ) " & _
                    "SELECT '" & FolioNumber & "' AS Folio, CustomerID, Address, CountryID, CompanyID, IDTypeID, IDNumber, " & dcRoomNumber.BoundText & ", DateIn, DateOut, RateType, Rate, Days, Adults, Childrens, BusinessSourceID, VehicleID, VehicleModel, PlateNo, 'Check In' AS CheckIn, Notes, ReservedBy, " & CurrUser.USER_PK & ", " & CurrUser.USER_PK & " " & _
                    "FROM Reservation " & _
                    "WHERE ReservationNo='" & txtReservationNo.Text & "'"
        
        Call AddRatePerPeriod(FolioNumber)
        
        ChangeValue CN, "Rooms", "RoomStatusID", 2, True, "WHERE RoomNumber = " & dcRoomNumber.BoundText
        ChangeValue CN, "Reservation", "Status", "Check In", False, "WHERE ReservationNo = '" & txtReservationNo.Text & "'"
        
        CN.Execute "UPDATE [Inventory] SET Status = 'Check In', ID = '" & FolioNumber & "' " & _
                    "WHERE ID= '" & txtReservationNo.Text & "' AND RoomNumber= " & dcRoomNumber.BoundText & " AND Status='Reservation'"

        CN.CommitTrans
        Unload Me
    End If
    
    Exit Sub
    
err:
    CN.RollbackTrans
    prompt_err err, Name, "CmdCheckIn_Click"
    Screen.MousePointer = vbDefault
End Sub

Private Sub AddRatePerPeriod(FolioNumber As String)
On Error GoTo err
    
    Dim rsRatePerPeriod As New ADODB.Recordset
    Dim tmpDate As Date

    tmpDate = dtpDateIn.Value
    
    If txtAdults.Tag = "" Then Exit Sub
    
    Dim intAdults As Integer
    
    If txtAdults.Text = hsAdults.Min Then
        intAdults = 0
    Else
        intAdults = CInt(txtAdults.Text) - hsAdults.Min
    End If

    With rsRatePerPeriod
        .Open "SELECT * FROM [Rate Per Period] WHERE FolioNumber = '" & FolioNumber & "' ORDER BY [Date]", CN, adOpenStatic, adLockOptimistic

        Do Until tmpDate > dtpDateOut.Value - 1
'            .Filter = "[Date] = #" & tmpDate & "#"
            
'            If .RecordCount = 0 Then
                .AddNew
                
                .Fields("FolioNumber") = FolioNumber
                .Fields("Date") = tmpDate
                .Fields("RoomNumber") = dcRoomNumber.BoundText
                .Fields("RateTypeID") = dcRateType.BoundText
                .Fields("Rate") = txtRate.Text
                .Fields("Adults") = txtAdults.Tag * intAdults
                .Fields("Childrens") = toMoney(txtChildrens.Tag) * toNumber(txtChildrens.Text)
                
                .Update
'            End If
            tmpDate = tmpDate + 1
        Loop
    End With
    
    rsRatePerPeriod.Close

    Exit Sub

err:
    prompt_err err, Name, "AddRatePerPeriod"
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub ResetFields()
'  clearText Me
'
'  txtEntry(15).Text = "0.00"
'  txtEntry(1).SetFocus
End Sub

Private Sub cmdLookupComp_Click()
    With frmCompanyLookup
        Set .RefForm = Me
        
        .Show vbModal
    End With
End Sub

Private Sub cmdLookupCust_Click()
    With frmCustomerLookup
        Set .RefForm = Me
        
        .Show vbModal
    End With
End Sub

Private Sub CmdPrint_Click()
    With frmReports
        .strReport = "Reservation"
        
        .strWhere = "{qry_rpt_Reservation.ReservationNo} = '" & txtReservationNo.Text & "'"
        
        frmReports.Show vbModal
    End With
End Sub

Private Sub Save_Record()
On Error GoTo err
    
    Dim rsCustomers As New Recordset
    Dim CustomerID As Integer
    Dim CompanyID As Integer

    If Trim(txtLastName.Text) = "" Or Trim(txtFirstName.Text) = "" Then
        MsgBox "Please complete the name of a guest.", vbInformation
        
        Exit Sub
    End If
        
    CN.BeginTrans

    With rsCustomers
        .Open "SELECT * FROM Customers WHERE LastName = '" & txtLastName.Text & "' AND FirstName = '" & txtFirstName.Text & "'", CN, adOpenStatic, adLockOptimistic
    
        If .RecordCount > 0 Then
            txtLastName.Tag = .Fields("CustomerID")
        Else
            .AddNew
            
            CustomerID = getIndex("Customers")
            txtLastName.Tag = CustomerID
            
            .Fields("CustomerID") = CustomerID
            .Fields("LastName") = txtLastName.Text
            .Fields("FirstName") = txtFirstName.Text
            
            .Update
        End If
        
        .Close
        
        If txtCompany.Text = "" Then GoSub ContinueSave
        
        .Open "SELECT * FROM Company WHERE Company = '" & txtCompany.Text & "'", CN, adOpenStatic, adLockOptimistic
        
        If .RecordCount > 0 Then
            txtCompany.Tag = .Fields("CompanyID")
        Else
            .AddNew

            CompanyID = getIndex("Company")
            txtCompany.Tag = CompanyID
            
            .Fields("CompanyID") = CompanyID
            .Fields("Company") = txtCompany.Text
            
            .Update
        End If
        
        .Close
    End With
    
ContinueSave:

    If State = adStateAddMode Then
        RS.AddNew
        
        RS.Fields("ReservationNo") = txtReservationNo.Text
        RS.Fields("ReservedBy") = CurrUser.USER_PK
        RS.Fields("AddedByFK") = CurrUser.USER_PK
    Else
        RS.Fields("DateModified") = Now
        RS.Fields("LastUserFK") = CurrUser.USER_PK
    End If
    
    With RS
        .Fields("CustomerID") = txtLastName.Tag
        .Fields("Address") = txtAddress.Text
        .Fields("CountryID") = dcCountry.BoundText
        .Fields("CompanyID") = IIf(txtCompany.Tag = "", Null, txtCompany.Tag)
        .Fields("IDTypeID") = dcIDType.BoundText
        .Fields("IDNumber") = txtIDNumber.Text
        .Fields("RoomNumber") = dcRoomNumber.BoundText
        .Fields("DateIn") = dtpDateIn.Value
        .Fields("DateOut") = dtpDateOut.Value
        .Fields("RateType") = dcRateType.BoundText
        .Fields("Rate") = txtRate.Text
        .Fields("OtherCharges") = txtOtherCharges.Text
        .Fields("Discount") = txtDiscount.Text
        .Fields("AmountPaid") = txtAmountPaid.Text
        .Fields("Days") = txtDays.Text
        .Fields("Adults") = txtAdults.Text
        .Fields("Childrens") = txtChildrens.Text
        .Fields("Total") = txtTotal.Text
        .Fields("BusinessSourceID") = dcBusSource.BoundText
        .Fields("VehicleID") = IIf(dcVehicle.BoundText = "", Null, dcVehicle.BoundText)
        .Fields("VehicleModel") = txtVehicleModel.Text
        .Fields("PlateNo") = txtPlateNo.Text
        .Fields("Notes") = txtNotes.Text

        .Update
    End With

    '----------------------------
    'Delete record from Inventory and add a new check in/out date
    CN.Execute "DELETE ID, Status " & _
                "From [Inventory] " & _
                "WHERE ID='" & txtReservationNo.Text & "' AND Status='Reservation'"
                
    Dim dtpStartDate As Date
    
    dtpStartDate = dtpDateIn.Value
    
    Do Until dtpStartDate = dtpDateOut.Value
        CN.Execute "INSERT INTO [Inventory] ( ID, RoomNumber, [Date], CustomerID, Status ) " & _
                "VALUES ('" & txtReservationNo.Text & "', " & dcRoomNumber.BoundText & ", #" & dtpStartDate & "#, " & txtLastName.Tag & ", 'Reservation')"

        dtpStartDate = dtpStartDate + 1
    Loop
    '----------------------------

    CN.CommitTrans

    HaveAction = True
    
    Exit Sub

err:
    CN.RollbackTrans
    prompt_err err, Name, "cmdSave_Click"
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSave_Click()
    Call Save_Record
    
    If State = adStateAddMode And HaveAction = True Then
        MsgBox "New record has been successfully saved.", vbInformation

'        Unload frmReservation
    ElseIf State = adStateEditMode And HaveAction = True Then
        MsgBox "Changes in  record has been successfully saved.", vbInformation
        
'        Unload frmReservation
    End If
End Sub

Private Sub cmdUsrHistory_Click()
    On Error Resume Next
    Dim tDate1 As String
    Dim tDate2 As String
    Dim tUser1 As String
    Dim tUser2 As String
    
    tDate1 = Format$(RS.Fields("DateAdded"), "MMM-dd-yyyy HH:MM AMPM")
    tDate2 = Format$(RS.Fields("DateModified"), "MMM-dd-yyyy HH:MM AMPM")
    
    tUser1 = getValueAt("SELECT PK,CompleteName FROM Users WHERE PK = " & RS.Fields("AddedByFK"), "CompleteName")
    tUser2 = getValueAt("SELECT PK,CompleteName FROM Users WHERE PK = " & RS.Fields("LastUserFK"), "CompleteName")
    
    MsgBox "Date Added: " & tDate1 & vbCrLf & _
           "Added By: " & tUser1 & vbCrLf & _
           "" & vbCrLf & _
           "Last Modified: " & tDate2 & vbCrLf & _
           "Modified By: " & tUser2, vbInformation, "Modification History"
           
    tDate1 = vbNullString
    tDate2 = vbNullString
    tUser1 = vbNullString
    tUser2 = vbNullString
End Sub

Private Sub dcRateType_Click(Area As Integer)
On Error GoTo err

    Dim rsRoomRates As New ADODB.Recordset
    
    With rsRoomRates
        .Open "SELECT * FROM [Room Rates] WHERE RoomNumber = " & dcRoomNumber.BoundText & " AND RateTypeID = " & dcRateType.BoundText, CN, adOpenStatic, adLockOptimistic
    
        If .RecordCount > 0 Then
            txtRate.Text = toMoney(!RoomRate)
            txtAdults.Text = !NoofPerson
            hsAdults.Min = !NoofPerson
            hsAdults.Value = !NoofPerson
            txtAdults.Tag = !ExtraAdultRate
            txtChildrens.Tag = !ExtraChildRate
        End If
    End With
    
    rsRoomRates.Close
    
    Call ComputeRate
    
    Exit Sub
    
err:
    CN.RollbackTrans
    prompt_err err, Name, "dcRateType_Click"
    Screen.MousePointer = vbDefault
End Sub

Private Sub dcRoomNumber_Click(Area As Integer)
    dcRateType_Click 0
End Sub

Public Sub dtpDateOut_Change()
    If dtpDateOut.Value < dtpDateIn.Value Then Exit Sub
    
    txtDays.Text = dtpDateOut.Value - dtpDateIn.Value
    hsDays.Value = txtDays.Text
    
    Call ComputeRate
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
On Error GoTo err

    RS.CursorLocation = adUseClient

    bind_dc "SELECT * FROM qry_Available_Rooms", "AvailableRoom", dcRoomNumber, "RoomNumber", True
    bind_dc "SELECT * FROM Countries", "Country", dcCountry, "CountryID", True
    bind_dc "SELECT * FROM [ID Type]", "IDType", dcIDType, "IDTypeID", True
    bind_dc "SELECT * FROM [Rate Type]", "RateType", dcRateType, "RateTypeID", True
    bind_dc "SELECT * FROM [Business Source]", "BusinessSource", dcBusSource, "BusinessSourceID", True
    bind_dc "SELECT * FROM [Vehicles]", "Vehicle", dcVehicle, "VehicleID", False
    
    dcCountry.Text = "Philippines"

    Dim rsRoomRates As New ADODB.Recordset

    'Check the form state
    If State = adStateAddMode Then
        RS.Open "SELECT * FROM Reservation WHERE ReservationNo = '" & PK & "'", CN, adOpenStatic, adLockOptimistic

        cmdUsrHistory.Enabled = False

        dtpDateIn.Value = Date
        dtpDateOut.Value = dtpDateIn.Value + 1

        GeneratePK

        rsRoomRates.Open "SELECT * FROM [Room Rates] WHERE RoomNumber = " & dcRoomNumber.BoundText & " AND RateTypeID = " & dcRateType.BoundText, CN, adOpenStatic, adLockOptimistic

        With rsRoomRates
            If .RecordCount > 0 Then
                txtRate.Text = toMoney(!RoomRate)
                txtAdults.Text = !NoofPerson
                hsAdults.Min = !NoofPerson
                hsAdults.Value = !NoofPerson
                txtAdults.Tag = !ExtraAdultRate
                txtChildrens.Tag = !ExtraChildRate
            End If
        End With

        dcBusSource.BoundText = 2

'        Call txtDays_Change

        Call ComputeRate
    Else
        RS.Open "SELECT * FROM Reservation WHERE ReservationNo = '" & PK & "'", CN, adOpenStatic, adLockOptimistic

        rsRoomRates.Open "SELECT * FROM [Room Rates] WHERE RoomNumber = " & RightSplitUF(dcRoomNumber.Tag) & " AND RateTypeID = " & dcRateType.BoundText, CN, adOpenStatic, adLockOptimistic

        With rsRoomRates
            If .RecordCount > 0 Then
                txtRate.Text = toMoney(!RoomRate)
                hsAdults.Min = !NoofPerson
                txtAdults.Tag = !ExtraAdultRate
                txtChildrens.Tag = !ExtraChildRate
            End If
        End With

        DisplayForEditing
        
        CmdCheckIn.Visible = True
        CmdPrint.Visible = True
        CmdCancel.Visible = True
        
        Call ComputeRate
    End If

    rsRoomRates.Close

    Exit Sub
                
err:
    prompt_err err, Name, "Form_Load"
    Screen.MousePointer = vbDefault
End Sub

'Procedure used to generate PK
Private Sub GeneratePK()
    PK = getIndex("Reservation")
    txtReservationNo.Text = GenerateID(PK, Format$(Date, "yy") & "-", "00000")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True And Shortcut = False Then
        frmReservationList.RefreshRecords
    End If
    
    Set frmReservation = Nothing
End Sub

Private Sub hsAdults_Change()
    txtAdults.Text = hsAdults.Value
    
'    Call ComputeAdultsRate
    Call ComputeRate
End Sub

Private Sub hsChildrens_Change()
    txtChildrens.Text = hsChildrens.Value
    
'    Call ComputeChildrensRate
    Call ComputeRate
End Sub

Private Sub hsDays_Change()
    dtpDateOut.Value = dtpDateIn.Value + hsDays.Value
    
    txtDays.Text = hsDays.Value
    
    Call ComputeRate
End Sub

Private Sub ComputeRate()
    If txtAdults.Tag = "" Then Exit Sub
    
    Dim intAdults As Integer
    
    If txtAdults.Text = hsAdults.Min Then
        intAdults = 0
    Else
        intAdults = CInt(txtAdults.Text) - hsAdults.Min
    End If
    
    txtTotalCharges.Text = (toMoney(txtRate.Text) * toNumber(txtDays.Text)) + (toMoney(txtAdults.Tag) * intAdults) + (toMoney(txtChildrens.Tag) * toNumber(txtChildrens.Text))
    txtTotalCharges.Text = toMoney(txtTotalCharges.Text)
    txtSubTotal.Text = toMoney(toNumber(txtTotalCharges.Text) + toNumber(txtOtherCharges.Text))
    txtTotal.Text = toMoney(toNumber(txtSubTotal.Text) - (toNumber(txtSubTotal.Text) * toNumber(txtDiscount.Text) / 100))
    txtBalance.Text = toMoney(toNumber(txtTotal.Text) - toNumber(txtAmountPaid.Text))
End Sub

Private Sub txtAmountPaid_Change()
    txtBalance.Text = toMoney(toNumber(txtTotal.Text) - toNumber(txtAmountPaid.Text))
End Sub

Private Sub txtDiscount_Change()
    Call ComputeRate
End Sub

Private Sub txtDiscount_GotFocus()
    HLText txtDiscount
End Sub

Private Sub txtDiscount_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtDiscount_Validate(Cancel As Boolean)
    txtDiscount.Text = toMoney(txtDiscount.Text)
End Sub

Private Sub txtRate_GotFocus()
    HLText txtRate
End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtRate_Validate(Cancel As Boolean)
    txtRate.Text = toMoney(txtRate.Text)
End Sub
