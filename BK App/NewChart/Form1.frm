VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   9780
   StartUpPosition =   3  'Windows Default
   Begin Project1.Chart Chart1 
      Height          =   5175
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   9105
      _ExtentX        =   17119
      _ExtentY        =   9551
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub RefreshGraph()
Dim myArray(2) As String
Dim MyColor(2) As Long
Dim MyCaption(10) As String
Dim MyLegend(2) As String
Dim i As Integer
For i = 1 To 10
    MyCaption(i) = "Test " & i
Next i

    myArray(0) = "10,14,23,19,18,25,15,17,13,15"
    MyColor(0) = vbRed
    myArray(1) = "8,9,6,11,15,29,18,20,23,27"
    MyColor(1) = vbBlue
    myArray(2) = "12,8,7,13,14,22,12,10,25,23"
    MyColor(2) = vbGreen
    
    Chart1.MaxValue = 29
    Chart1.MinValue = 0
    Chart1.Rows = 10
    Chart1.Cols = 3
    Chart1.DrawGraph myArray, MyColor, MyCaption
    
    MyLegend(0) = "Prvi objekt u legendi"
    MyLegend(1) = "Drugi objekt u legendi"
    MyLegend(2) = "Treci objekt u legendi"
    Chart1.DrawLegend MyLegend, MyColor
End Sub

Private Sub Form_Load()
    RefreshGraph
End Sub

Private Sub Form_Resize()
    Chart1.Width = Me.ScaleWidth - 30
    Chart1.Height = Me.ScaleHeight - 30
End Sub
