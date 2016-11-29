VERSION 5.00
Begin VB.Form frmmain 
   Caption         =   "My Dental Clinic "
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   6720
      TabIndex        =   2
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Report"
      Height          =   495
      Left            =   5520
      TabIndex        =   1
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdPatient 
      Caption         =   "Patient"
      Height          =   495
      Left            =   4320
      TabIndex        =   0
      Top             =   3840
      Width           =   1215
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
 End
End Sub

Private Sub cmdPatient_Click()
 frmpatient.Show
End Sub

Private Sub cmdReport_Click()
 frmReport.Show
End Sub


Private Sub Form_Load()

End Sub
