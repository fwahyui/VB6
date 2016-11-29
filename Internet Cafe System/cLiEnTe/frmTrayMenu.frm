VERSION 5.00
Begin VB.Form frmTrayMenu 
   Caption         =   "Tray Menu"
   ClientHeight    =   705
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   2760
   Icon            =   "frmTrayMenu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   705
   ScaleWidth      =   2760
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuTray 
      Caption         =   "TrayMenu"
      Begin VB.Menu mnuMax 
         Caption         =   "Maximize"
      End
      Begin VB.Menu mnuLogOut 
         Caption         =   "Log Out"
      End
   End
End
Attribute VB_Name = "frmTrayMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuLogOut_Click()
  frmPopUp.cmdLogout_Click
End Sub

Private Sub mnuMax_Click()
  frmPopUp.Tray.Visible = False
  frmPopUp.WindowState = 0
  frmPopUp.Show
End Sub
