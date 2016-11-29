VERSION 5.00
Object = "{02B5E320-7292-11CF-93D5-0020AF99504A}#1.0#0"; "MSCHART.OCX"
Begin VB.Form FX_MSchart 
   Caption         =   "MS Chart Example"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   9060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrintGraph 
      Caption         =   "Print"
      Height          =   735
      Left            =   7680
      Picture         =   "FX_MSChart.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Frame chartTypeReq 
      Height          =   1455
      Left            =   720
      TabIndex        =   4
      Top             =   5160
      Width           =   6375
      Begin VB.OptionButton graphType1 
         Caption         =   "2D Bar"
         Height          =   1245
         Index           =   1
         Left            =   75
         Picture         =   "FX_MSChart.frx":0ACE
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "grCol2d"
         Top             =   150
         Value           =   -1  'True
         Width           =   990
      End
      Begin VB.OptionButton graphType1 
         Caption         =   "2D Line"
         Height          =   1245
         Index           =   2
         Left            =   1117
         Picture         =   "FX_MSChart.frx":3108
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "grLine2D"
         Top             =   150
         Width           =   990
      End
      Begin VB.OptionButton graphType1 
         Caption         =   "2D Area"
         Height          =   1245
         Index           =   3
         Left            =   2144
         Picture         =   "FX_MSChart.frx":545A
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "grArea2D"
         Top             =   150
         Width           =   990
      End
      Begin VB.OptionButton graphType1 
         Caption         =   "2D Stack"
         Height          =   1245
         Index           =   4
         Left            =   3171
         Picture         =   "FX_MSChart.frx":78E4
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   165
         Width           =   990
      End
      Begin VB.OptionButton graphType1 
         Caption         =   "3D Bar"
         Height          =   1245
         Index           =   5
         Left            =   4198
         Picture         =   "FX_MSChart.frx":9C16
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   150
         Width           =   990
      End
      Begin VB.OptionButton graphType1 
         Caption         =   "2D Pie"
         Height          =   1245
         Index           =   6
         Left            =   5225
         Picture         =   "FX_MSChart.frx":BEB8
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   150
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdAddLegend 
      Caption         =   "Add Legend"
      Height          =   495
      Left            =   7680
      TabIndex        =   3
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdAddTitle 
      Caption         =   "Add Title"
      Height          =   495
      Left            =   7680
      TabIndex        =   2
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdAddData 
      Caption         =   "Add Demo Data"
      Height          =   855
      Left            =   7680
      Picture         =   "FX_MSChart.frx":DC5A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin MSChartLib.MSChart MSChart1 
      Height          =   4815
      Left            =   240
      OleObjectBlob   =   "FX_MSChart.frx":E8DC
      TabIndex        =   0
      Tag             =   "Print"
      Top             =   360
      Width           =   7215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Written by Garry Robinson --> www.gr-fx.com"
      Height          =   735
      Left            =   7560
      TabIndex        =   12
      Top             =   5640
      Width           =   1335
   End
End
Attribute VB_Name = "FX_MSchart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

 Dim tglLegend As Boolean, tglTitle As Boolean

Private Sub cmdAddData_Click()

On Error Resume Next

'  This subroutine shows you how to setup and pass data to the chart object
'  imax is the maximum of columns that you want to display as demo data
'  datascale will adjust the actual values to suit a different axis scale

   imax = 6
   Dim X() As Variant
   Dim iRow As Integer
  
   ReDim X(1 To 8, 1 To imax + 1)
  
    
  dataScale = 1
  
  iRow = 1
  X(1, iRow) = "Period"
  X(2, iRow) = "June 1998"
  X(3, iRow) = "June 1999"
  X(4, iRow) = "June 2000"
  X(5, iRow) = "June 2001"
  X(6, iRow) = "June 2002"
  X(7, iRow) = "June 2003"
  X(8, iRow) = "June 2004"
  
  For iRow = 2 To imax + 1
  
  X(1, iRow) = "Ann Depr%" & iRow
  X(2, iRow) = (12.2 + (iRow * Rnd) * 20) * dataScale
  X(3, iRow) = (45 + (iRow * Rnd) * 20) * dataScale
  X(4, iRow) = (36 + (iRow * Rnd) * 20) * dataScale
  X(5, iRow) = (28 + (iRow * Rnd) * 20) * dataScale
  X(6, iRow) = (38 + (iRow * Rnd) * 20) * dataScale
  X(7, iRow) = (25 + (iRow * Rnd) * 20) * dataScale
  X(8, iRow) = (16 + (iRow * Rnd) * 20) * dataScale
  
  Next iRow
  

addChartData:

' Reset the chart back to default to avoid any surprises
  MSChart1.ToDefaults
    
 
  Call addDataArray(MSChart1, X(), True)


End Sub



Private Sub cmdAddLegend_Click()

  tglLegend = Not tglLegend
  
  Call AddLegend(MSChart1, tglLegend)
End Sub

Private Sub cmdAddTitle_Click()
  tglTitle = Not tglTitle

  Call AddTitle(MSChart1, "Monthly Results", tglTitle)
End Sub



Private Sub cmdPrintGraph_Click()
On Error GoTo ErrorHandler

' First turn of all the controls that do not have a Tag property of Print

  Call PrintVisibility(False)

' Now print the objects that are still visible

  PrintForm

exitForm:
' And then turn all the other objects on again
  Call PrintVisibility(True)

Exit Sub
ErrorHandler:

  MsgBox "The form can't be printed."
  GoTo exitForm
  
End Sub
Private Sub PrintVisibility(visState As Boolean)
'  Hide or show all objects with a tag of Print
 
 Dim i As Integer

'  First turn off the other controls

 For i = 0 To Me.Count - 1
   If Me(i).Tag <> "Print" Then
     Me(i).Visible = visState
   End If
 Next

End Sub

Private Sub graphType1_Click(Index As Integer)
On Error Resume Next
  Dim graphInt As Integer, chartStr As String
       
  With MSChart1
  
  chartStr = graphType1(Index).Caption
  
  Select Case chartStr
 
      Case Is = "2D Area"
      
         .chartType = VtChChartType2dArea
         .Stacking = True
         
      Case Is = "2D Bar"
      
         .chartType = VtChChartType2dBar
         .Stacking = False
  
     Case Is = "3D Bar"
  
       .chartType = VtChSeriesType3dBar
       .Plot.Projection = VtProjectionTypeOblique
       .Stacking = True
    
     Case Is = "2D Stack"
  
         .chartType = VtChChartType2dBar
         .Stacking = True
      
     Case Is = "2D Line"
     
       .chartType = VtChChartType2dLine
       .Stacking = False
     
     Case Is = "2D Pie"
     
       .chartType = VtChChartType2dPie
       .Stacking = False
  
     Case Else
      
          MsgBox "Chart Type Not Supported"
          
  End Select
  End With
End Sub

