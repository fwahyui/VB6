VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form6"
   ClientHeight    =   4140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3720
   LinkTopic       =   "Form6"
   ScaleHeight     =   4140
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label hs8 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   18
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   3480
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label hs5 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   23
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label hs4 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   22
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label hs3 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   21
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label hs2 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   20
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label hs1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   19
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Line Line5 
      X1              =   3480
      X2              =   3480
      Y1              =   3240
      Y2              =   1440
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   120
      Y1              =   1440
      Y2              =   3240
   End
   Begin VB.Line Line3 
      X1              =   960
      X2              =   960
      Y1              =   3240
      Y2              =   1440
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Payment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Discount"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cost"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Duration"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Timeout"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Timein"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Date"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "User"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "RECEIPTS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   0
      Width           =   1455
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3480
      Y1              =   1440
      Y2              =   1440
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Labelwarnet1 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Labelwarnet2 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Phone"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Labelwarnet3 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   "Name    :"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "Address :"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label35 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone   :"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label36 
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail    :"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   2760
      Picture         =   "printout.frx":0000
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   750
   End
   Begin VB.Label hs7 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   25
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label hs6 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   24
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "<< THANK YOU >>"
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
      TabIndex        =   17
      Top             =   3720
      Width           =   3375
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MWARNET 2 - FREEWARE EDITION
'COPYRIGHT(C) 2007 MTechnologi Bali Indonesia
'Programed by A.A.Ngr.Manik Artawan
'e-mail : gungmanik@telkom.net
'---------------------------------------------
'THANK YOu FOR DOWNLOAD THIS SMALL APPLICATION
'---------------------------------------------


Private Sub Form_Load()
'---- Cybercafe ID
Labelwarnet.Caption = Form1.Labelwarnet.Caption
Labelwarnet1.Caption = Form1.Labelwarnet1.Caption
Labelwarnet2.Caption = Form1.Labelwarnet2.Caption
Labelwarnet3.Caption = Form1.Labelwarnet3.Caption
On Error Resume Next
'---- Load your stamp
Image1.Picture = LoadPicture(App.Path & "\stamp.jpg")
'---- Payment details
hs1.Caption = Form1.hx1.Text
hs2.Caption = Form1.hx2.Text
hs3.Caption = Form1.hx3.Text
hs4.Caption = Form1.hx4.Text
hs5.Caption = Form1.hx5.Text
hs6.Caption = Form1.hx6.Text
hs7.Caption = Form1.hx7.Text
hs8.Caption = Form1.hx8.Text
'---- PRINTING
On Error GoTo t
Form6.PrintForm
Form6.Hide
t:
End Sub
