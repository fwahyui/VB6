VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ChangeGraf 
   Caption         =   "Demonstration Of MS Chart Manipulation (using Graph8 /97)"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   765
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   9810
   Begin VB.CommandButton cmdAddData 
      Caption         =   "Add Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   8745
      TabIndex        =   46
      Top             =   270
      Width           =   1020
   End
   Begin VB.CommandButton PrintIt 
      Caption         =   "Print Graph"
      Height          =   480
      Left            =   8370
      TabIndex        =   45
      Top             =   1335
      Width           =   900
   End
   Begin VB.CommandButton setColours 
      Caption         =   "Colours"
      Height          =   345
      Left            =   8340
      TabIndex        =   44
      Top             =   2055
      Width           =   930
   End
   Begin VB.PictureBox Picture1 
      Height          =   525
      Left            =   8580
      Picture         =   "graf1.frx":0000
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   42
      Top             =   5205
      Width           =   525
   End
   Begin VB.ListBox SeriesNum 
      Height          =   645
      ItemData        =   "graf1.frx":030A
      Left            =   8610
      List            =   "graf1.frx":030C
      TabIndex        =   41
      Top             =   3240
      Width           =   450
   End
   Begin ComctlLib.Slider grafTilt 
      Height          =   915
      Left            =   8280
      TabIndex        =   38
      Top             =   4215
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   1614
      _Version        =   327682
      Orientation     =   1
      LargeChange     =   15
      SmallChange     =   5
      Min             =   -90
      Max             =   90
      TickStyle       =   3
   End
   Begin ComctlLib.Slider grafRotate 
      Height          =   345
      Left            =   7035
      TabIndex        =   37
      Top             =   5145
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   609
      _Version        =   327682
      LargeChange     =   30
      SmallChange     =   10
      Max             =   360
      TickStyle       =   3
   End
   Begin VB.CommandButton cmdScale 
      Caption         =   "Scale"
      Height          =   465
      Left            =   270
      TabIndex        =   28
      Top             =   5175
      Width           =   870
   End
   Begin VB.TextBox MajUnit 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   450
      TabIndex        =   24
      Text            =   "Auto"
      Top             =   2775
      Width           =   720
   End
   Begin VB.TextBox botScale 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   270
      TabIndex        =   23
      Text            =   "Auto Scale"
      Top             =   4530
      Width           =   915
   End
   Begin VB.TextBox TopScale 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   255
      TabIndex        =   22
      Text            =   "Auto Scale"
      Top             =   1470
      Width           =   915
   End
   Begin VB.CommandButton tglValues 
      Caption         =   "Values"
      Height          =   330
      Left            =   8340
      TabIndex        =   21
      Top             =   2880
      Width           =   930
   End
   Begin VB.CommandButton tglLegend 
      Caption         =   "Legend"
      Height          =   330
      Left            =   8340
      TabIndex        =   20
      Top             =   2475
      Width           =   930
   End
   Begin VB.CommandButton cmdAddTitle 
      Caption         =   "Add Title"
      Height          =   345
      Left            =   1860
      TabIndex        =   19
      Top             =   5340
      Width           =   855
   End
   Begin VB.TextBox MainTitle 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   2775
      TabIndex        =   18
      Text            =   "Demonstration Of MS Chart/Graph"
      Top             =   5355
      Width           =   4035
   End
   Begin VB.Frame ChangeGrafs 
      Caption         =   "Choose A Graf Style"
      Height          =   1155
      Left            =   270
      TabIndex        =   1
      Top             =   30
      Width           =   8295
      Begin VB.OptionButton Line 
         Caption         =   "Line "
         Height          =   240
         Index           =   1
         Left            =   180
         TabIndex        =   36
         Top             =   780
         Width           =   1350
      End
      Begin VB.OptionButton Line3D 
         Caption         =   "3D"
         Height          =   195
         Index           =   1
         Left            =   1605
         TabIndex        =   35
         Top             =   825
         Width           =   570
      End
      Begin VB.OptionButton LineSt 
         Caption         =   "Stacked"
         Height          =   225
         Index           =   1
         Left            =   2205
         TabIndex        =   34
         Top             =   795
         Width           =   930
      End
      Begin VB.OptionButton LineMarks 
         Caption         =   "Marks"
         Height          =   225
         Index           =   1
         Left            =   3165
         TabIndex        =   33
         Top             =   795
         Width           =   795
      End
      Begin VB.OptionButton Bubble 
         Caption         =   "Bubble"
         Height          =   210
         Left            =   4380
         TabIndex        =   32
         Top             =   810
         Width           =   855
      End
      Begin VB.OptionButton Bubble3D 
         Caption         =   "3D"
         Height          =   240
         Left            =   5220
         TabIndex        =   31
         Top             =   780
         Width           =   690
      End
      Begin VB.OptionButton Stock 
         Caption         =   "Stock"
         Height          =   240
         Left            =   7290
         TabIndex        =   30
         Top             =   780
         Width           =   795
      End
      Begin VB.OptionButton Cylinder 
         Caption         =   "Cylinder"
         Height          =   210
         Left            =   6345
         TabIndex        =   29
         Top             =   810
         Width           =   900
      End
      Begin VB.OptionButton SurfTop 
         Caption         =   "Top"
         Height          =   240
         Left            =   7290
         TabIndex        =   17
         Top             =   540
         Width           =   690
      End
      Begin VB.OptionButton Surf 
         Caption         =   "Surface"
         Height          =   210
         Left            =   6345
         TabIndex        =   16
         Top             =   570
         Width           =   900
      End
      Begin VB.OptionButton Area 
         Caption         =   "Area"
         Height          =   210
         Left            =   6345
         TabIndex        =   15
         Top             =   300
         Width           =   720
      End
      Begin VB.OptionButton Area3D 
         Caption         =   "3D"
         Height          =   240
         Left            =   7290
         TabIndex        =   14
         Top             =   300
         Width           =   690
      End
      Begin VB.OptionButton XYScat 
         Caption         =   "Scatter"
         Height          =   210
         Left            =   4380
         TabIndex        =   13
         Top             =   540
         Width           =   855
      End
      Begin VB.OptionButton XYScatSm 
         Caption         =   "Lines"
         Height          =   240
         Left            =   5220
         TabIndex        =   12
         Top             =   540
         Width           =   690
      End
      Begin VB.OptionButton Pie3D 
         Caption         =   "3D"
         Height          =   240
         Left            =   5220
         TabIndex        =   11
         Top             =   285
         Width           =   690
      End
      Begin VB.OptionButton Pie 
         Caption         =   "Pie"
         Height          =   210
         Left            =   4380
         TabIndex        =   10
         Top             =   300
         Width           =   615
      End
      Begin VB.OptionButton ColCl 
         Caption         =   "Vertical Bar"
         Height          =   240
         Index           =   0
         Left            =   195
         TabIndex        =   9
         Top             =   540
         Width           =   1350
      End
      Begin VB.OptionButton ColCl3D 
         Caption         =   "3D"
         Height          =   195
         Index           =   0
         Left            =   1605
         TabIndex        =   8
         Top             =   540
         Width           =   570
      End
      Begin VB.OptionButton ColSt 
         Caption         =   "Stacked"
         Height          =   225
         Index           =   0
         Left            =   2205
         TabIndex        =   7
         Top             =   540
         Width           =   930
      End
      Begin VB.OptionButton ColSt3D 
         Caption         =   "3D"
         Height          =   225
         Index           =   0
         Left            =   3165
         TabIndex        =   6
         Top             =   540
         Width           =   525
      End
      Begin VB.OptionButton BarSt3D 
         Caption         =   "3D"
         Height          =   225
         Left            =   3165
         TabIndex        =   5
         Top             =   285
         Width           =   525
      End
      Begin VB.OptionButton BarSt 
         Caption         =   "Stacked"
         Height          =   225
         Left            =   2205
         TabIndex        =   4
         Top             =   285
         Width           =   930
      End
      Begin VB.OptionButton BarCl3D 
         Caption         =   "3D"
         Height          =   195
         Left            =   1605
         TabIndex        =   3
         Top             =   285
         Width           =   570
      End
      Begin VB.OptionButton BarCl 
         Caption         =   "Horizontal Bar"
         Height          =   240
         Left            =   195
         TabIndex        =   2
         Top             =   240
         Width           =   1350
      End
   End
   Begin VB.Label Label6 
      Caption         =   "www.gr-fx.com"
      Height          =   240
      Left            =   8325
      TabIndex        =   43
      Top             =   5730
      Width           =   1125
   End
   Begin VB.Label Label5 
      Caption         =   "Tilt"
      Height          =   270
      Left            =   8700
      TabIndex        =   40
      Top             =   4530
      Width           =   555
   End
   Begin VB.Label Label4 
      Caption         =   "Rotate"
      Height          =   270
      Left            =   7485
      TabIndex        =   39
      Top             =   5475
      Width           =   555
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Major Unit"
      Height          =   255
      Left            =   195
      TabIndex        =   27
      Top             =   3105
      Width           =   945
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Low Value"
      Height          =   255
      Left            =   225
      TabIndex        =   26
      Top             =   4875
      Width           =   945
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "High Value"
      Height          =   255
      Left            =   225
      TabIndex        =   25
      Top             =   1815
      Width           =   945
   End
   Begin VB.OLE chart1 
      Class           =   "MSGraph.Chart.8"
      Height          =   3810
      Left            =   1230
      OleObjectBlob   =   "graf1.frx":030E
      SizeMode        =   1  'Stretch
      TabIndex        =   0
      Tag             =   "Print"
      Top             =   1335
      Width           =   6990
   End
End
Attribute VB_Name = "ChangeGraf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim MyGrafObj As Object

Private Sub Area_Click()
 MyGrafObj.ChartType = xlAreaStacked
End Sub

Private Sub Area3D_Click()
 MyGrafObj.ChartType = xl3DArea
End Sub

Private Sub BarCl_Click()
 MyGrafObj.ChartType = xlBarClustered
End Sub

Private Sub BarCl3D_Click()
 MyGrafObj.ChartType = xl3DBarClustered
End Sub


Private Sub BarSt_Click()
 MyGrafObj.ChartType = xlBarStacked
End Sub

Private Sub BarSt3D_Click()
 MyGrafObj.ChartType = xl3DBarStacked
End Sub

Private Sub Bubble_Click()
 MyGrafObj.ChartType = xlBubble
End Sub

Private Sub Bubble3D_Click()
 MyGrafObj.ChartType = xlBubble3DEffect
End Sub



Private Sub cmdAddTitle_Click()

  ' Toggle the title on and off
  
  With MyGrafObj
    .HasTitle = Not .HasTitle
  
    If .HasTitle Then
    
'     Display the title
      .ChartTitle.Text = CStr(Me!MainTitle)
    End If
  End With
End Sub

Private Sub cmdAddData_Click()
'  Add some data to the graph

Dim msg, NewLine, Tabb

  Tabb = vbTab
  NewLine = vbNewLine
  
  msg = Tabb + "Beer" & Tabb & "Wine" & Tabb & "Spirits"
  msg = msg & NewLine + "Australia" & Tabb & 100 & Tabb & 65 & Tabb & 14
  msg = msg & NewLine + "USA" & Tabb & 120 & Tabb & 85 & Tabb & 33
  msg = msg & NewLine + "UK" & Tabb & 130 & Tabb & 53 & Tabb & 24

  With chart1
    
'     .DoVerb vbOLEShow
     .AppIsRunning
     If .AppIsRunning Then
     
       .DataText = msg
       .Update
     Else
     
       MsgBox "Graph is not active"
     End If
     
  End With
  
End Sub
Private Sub cmdScale_Click()

' Change the scale of the vertical graph axes

On Error GoTo cmdScale_Error:

With MyGrafObj
  
  If Left(Me![botScale], 4) = "Auto" Or Me![botScale] = "0" Then 'Set the minimum scale
    .Axes(2).MinimumScaleIsAuto = True
  Else
    .Axes(2).MinimumScale = CDbl(Me![botScale])
  End If
  
  If Left(Me![TopScale], 4) = "Auto" Or Me![TopScale] = "0" Then  'Set the maximum scale
    .Axes(2).MaximumScaleIsAuto = True
  Else
    .Axes(2).MaximumScale = CDbl(Me![TopScale])
  End If
 
  If Left(Me![MajUnit], 4) = "Auto" Or Me![MajUnit] = "0" Then  'Set the major unit
    .Axes(2).MajorUnitIsAuto = True
  Else
    .Axes(2).MajorUnit = CDbl(Me![MajUnit])
  End If
  
End With

 Exit Sub
 
cmdScale_Error:

   MsgBox "Problem with scaling ... " & Error, vbInformation, "Try Another Graph"
   Exit Sub

End Sub

Private Sub ColCl_Click(Index As Integer)
 MyGrafObj.ChartType = xlColumnClustered
End Sub



Private Sub ColCl3D_Click(Index As Integer)
 MyGrafObj.ChartType = xl3DColumnClustered

End Sub


Private Sub ColSt_Click(Index As Integer)
 MyGrafObj.ChartType = xlColumnStacked
End Sub

Private Sub ColSt3D_Click(Index As Integer)
 MyGrafObj.ChartType = xl3DColumnStacked
End Sub

Private Sub Cylinder_Click()
 MyGrafObj.ChartType = xlCylinderCol

End Sub

Private Sub Form_Load()

' Define an object so that reference can be used throughout the form

Set MyGrafObj = Me![chart1].object.Application.Chart

' Setup the list box choices
  SeriesNum.AddItem 1
  SeriesNum.AddItem 2
  SeriesNum.AddItem 3
  SeriesNum = 1
  
'  Set the standard format for the graph

  With chart1
    .Format = "CF_TEXT"    ' Search CF_Text for more info on formats supported
'    .SizeMode = vbOLESizeStretch
'    .CreateEmbed "", "MSGRAPH"
    
  End With
  
End Sub

Private Sub grafRotate_Click()

 On Error GoTo grafRotate_error
 
  With MyGrafObj
    .Rotation = Me!grafRotate
  End With
  
 Exit Sub
 
grafRotate_error:
 
   MsgBox "This graf does not support rotations ... " & Error, vbInformation, "Try Another Graph"
   Exit Sub
  
End Sub

Private Sub grafTilt_Click()

 On Error GoTo grafTilt_error
 
  With MyGrafObj
    .Elevation = Me!grafTilt
  End With
  
 Exit Sub
 
grafTilt_error:
 
   MsgBox "This graf does not support rotations", vbInformation, "Try Another Graph"
   Exit Sub
  
End Sub

Private Sub Line_Click(Index As Integer)
 MyGrafObj.ChartType = xlLine
End Sub

Private Sub Line3D_Click(Index As Integer)
  MyGrafObj.ChartType = xl3DLine
End Sub

Private Sub LineMarks_Click(Index As Integer)
 MyGrafObj.ChartType = xlLineMarkers
End Sub

Private Sub LineSt_Click(Index As Integer)
 MyGrafObj.ChartType = xlLineStacked
End Sub


Private Sub Pie_Click()
 MyGrafObj.ChartType = xlPie
End Sub

Private Sub Pie3D_Click()
 MyGrafObj.ChartType = xl3DPie
End Sub



Private Sub PrintIt_Click()
  Dim msg   ' Declare variable.
  On Error GoTo ErrorHandler   ' Set up error handler.
  
' First turn of all the controls that do not have a Tag property of Print

  Call PrintVisibility(False)
  
' Now print the objects that are still visible
  PrintForm   ' Print form.
  
' And then turn all the other objects on again
  Call PrintVisibility(True)

  Exit Sub
ErrorHandler:
   msg = "The form can't be printed."
   MsgBox msg   ' Display message.
   Resume Next
End Sub

Private Sub PrintVisibility(visState As Boolean)
  Dim i As Integer

' First turn off the other controls

  For i = 0 To Me.Count - 1
    If Me(i).Tag <> "Print" Then
      Me(i).Visible = visState
    End If
  Next

End Sub

Private Sub setColours_Click()
  
  Dim i As Integer, j As Integer, k As Integer
  
  On Error GoTo exitSetColours
  
  With MyGrafObj

    k = 1
    For i = 1 To 3
      For j = 1 To 4
        .SeriesCollection(i).Points(j).Interior.Color = QBColor(k)
        k = k + 1
      Next j
    Next i
    .Refresh
  End With
  
exitSetColours:

End Sub

Private Sub Stock_Click()
 MyGrafObj.ChartType = xlStockHLC
End Sub

Private Sub Surf_Click()
 MyGrafObj.ChartType = xlSurface
End Sub

Private Sub SurfTop_Click()
 MyGrafObj.ChartType = xlSurfaceTopView
End Sub

Private Sub tglLegend_Click()
      
' Toggle the legend
  With MyGrafObj
  
    .HasLegend = Not .HasLegend
  End With
End Sub

Private Sub tglValues_Click()

Static tglVal As Boolean, seriesVal As Long
  
On Error GoTo tglValues_error
  
  
' Show the legend for the first series collection

  tglVal = Not tglVal

  With MyGrafObj
  
    seriesVal = Me!SeriesNum
    If tglVal Then
      .SeriesCollection(seriesVal).ApplyDataLabels xlDataLabelsShowValue, False
     
    Else
      .SeriesCollection(seriesVal).ApplyDataLabels xlDataLabelsShowNone, True
    End If
  
  End With

Exit Sub

tglValues_error:

   MsgBox "This graf does not values fields ... " & Error, vbInformation, "Try Another Graph Format"
   Exit Sub


End Sub

Private Sub XYScat_Click()
 MyGrafObj.ChartType = xlXYScatter
End Sub

Private Sub XYScatSm_Click()
 MyGrafObj.ChartType = xlXYScatterSmooth
End Sub
