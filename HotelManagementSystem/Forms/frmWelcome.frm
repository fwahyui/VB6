VERSION 5.00
Begin VB.Form frmWelcome 
   BorderStyle     =   0  'None
   Caption         =   "Welcome"
   ClientHeight    =   5565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12780
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5565
   ScaleWidth      =   12780
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image1 
      Height          =   3465
      Left            =   0
      Picture         =   "frmWelcome.frx":0000
      Top             =   990
      Width           =   4635
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   4620
      Picture         =   "frmWelcome.frx":10F1
      Stretch         =   -1  'True
      Top             =   4140
      Width           =   15360
   End
   Begin VB.Image Image3 
      Height          =   270
      Left            =   4350
      Picture         =   "frmWelcome.frx":1183
      Top             =   4410
      Width           =   4245
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   12000
      Left            =   0
      Top             =   4380
      Width           =   15360
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowForm()
    mdiMain.AddChild Me
End Sub

Private Sub Form_Activate()
    Active
    mdiMain.ActivateChild Me
End Sub

Private Sub Form_Deactivate()
    Deactive
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'pass keyinfo to mdiMain
    mdiMain.AFForm_KeyDown KeyCode, Shift
End Sub

Public Sub Active()
    mdiMain.ShowTBButton "ttfffff"
    
    With mdiMain
        .tbMenu.Buttons(2).Caption = "User's Guide"
        .tbMenu.Buttons(2).Image = 10
        
        .tbMenu.Buttons(3).Caption = "About"
        .tbMenu.Buttons(3).Image = 11
        
        .mnuRACN.Caption = "User's Guide"
        .mnuRAES.Caption = "About"
    End With
End Sub

Private Sub Deactive()
    mdiMain.HideTBButton "", True
    With mdiMain
        .tbMenu.Buttons(2).Caption = "New"
        .tbMenu.Buttons(2).Image = 1
        
        .tbMenu.Buttons(3).Caption = "Edit"
        .tbMenu.Buttons(3).Image = 2
    
        .mnuRACN.Caption = "Create New"
        .mnuRAES.Caption = "Edit Selected"
    End With
End Sub

Public Sub CommandPass(ByVal srcPerformWhat As String)
    Select Case srcPerformWhat
        Case "New"
            '
        Case "Edit"
            frmAbout.Show vbModal
    End Select
End Sub
