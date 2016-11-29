VERSION 5.00
Begin VB.UserControl XChart 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8400
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   5580
   ScaleWidth      =   8400
   ToolboxBitmap   =   "XChart.ctx":0000
   Begin VB.PictureBox picCommands 
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   60
      ScaleHeight     =   330
      ScaleWidth      =   1935
      TabIndex        =   6
      Top             =   60
      Width           =   1935
      Begin VB.CommandButton cmdCmd 
         Height          =   315
         Index           =   5
         Left            =   1605
         Picture         =   "XChart.ctx":0312
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdCmd 
         Height          =   315
         Index           =   0
         Left            =   0
         Picture         =   "XChart.ctx":045C
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdCmd 
         Height          =   315
         Index           =   3
         Left            =   975
         Picture         =   "XChart.ctx":05A6
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdCmd 
         Height          =   315
         Index           =   1
         Left            =   330
         Picture         =   "XChart.ctx":06F0
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdCmd 
         Height          =   315
         Index           =   4
         Left            =   1290
         Picture         =   "XChart.ctx":083A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdCmd 
         Height          =   315
         Index           =   2
         Left            =   660
         Picture         =   "XChart.ctx":0984
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   315
      End
      Begin VB.Image imgCmd 
         Height          =   240
         Index           =   4
         Left            =   1470
         Picture         =   "XChart.ctx":0ACE
         Top             =   585
         Width           =   240
      End
      Begin VB.Image imgCmd 
         Height          =   240
         Index           =   3
         Left            =   1215
         Picture         =   "XChart.ctx":0C18
         Top             =   600
         Width           =   240
      End
      Begin VB.Image imgCmd 
         Height          =   240
         Index           =   2
         Left            =   930
         Picture         =   "XChart.ctx":0D62
         Top             =   600
         Width           =   240
      End
      Begin VB.Image imgCmd 
         Height          =   240
         Index           =   1
         Left            =   660
         Picture         =   "XChart.ctx":0EAC
         Top             =   600
         Width           =   240
      End
      Begin VB.Image imgCmd 
         Height          =   240
         Index           =   0
         Left            =   420
         Picture         =   "XChart.ctx":0FF6
         Top             =   600
         Width           =   240
      End
   End
   Begin VB.PictureBox picLegend 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F5F5&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFF0F0&
      ForeColor       =   &H00FF7040&
      Height          =   5430
      Left            =   3360
      ScaleHeight     =   5430
      ScaleWidth      =   2130
      TabIndex        =   1
      Top             =   0
      Width           =   2130
      Begin VB.VScrollBar vsbContainer 
         Height          =   5445
         LargeChange     =   5
         Left            =   1905
         Max             =   100
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   0
         Width           =   225
      End
      Begin VB.PictureBox picContainer 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00F0F5F5&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5205
         Left            =   240
         ScaleHeight     =   5205
         ScaleWidth      =   1665
         TabIndex        =   2
         Top             =   120
         Width           =   1665
         Begin VB.Label lblDescription 
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   315
            TabIndex        =   3
            Top             =   135
            Visible         =   0   'False
            Width           =   1290
         End
         Begin VB.Shape Box 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FFFFFF&
            BorderStyle     =   0  'Transparent
            Height          =   195
            Index           =   0
            Left            =   75
            Shape           =   5  'Rounded Square
            Top             =   150
            Visible         =   0   'False
            Width           =   195
         End
      End
      Begin VB.Label lblSlider 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "«"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   5430
         Left            =   0
         TabIndex        =   5
         ToolTipText     =   "Display Legend"
         Top             =   0
         Width           =   90
      End
   End
   Begin VB.Label lblInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   1800
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Main"
      Visible         =   0   'False
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectionInfo 
         Caption         =   "Selection information"
      End
      Begin VB.Menu mnuViewLegend 
         Caption         =   "Display Legend"
      End
      Begin VB.Menu mnuSeparator2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "Properties"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuLegend 
      Caption         =   "&Legend"
      Begin VB.Menu mnuLegendHide 
         Caption         =   "Hide"
      End
   End
End
Attribute VB_Name = "XChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type PointAPI   'API Point structure
    X   As Long
    Y   As Long
End Type

Private Const PI    As Double = 3.14159265358979
Private Const RADS  As Double = PI / 180    '<Degrees> * RADS = radians

Private Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As PointAPI, ByVal nCount As Long) As Long

Private uColumns()        As Double       'array of column height values
                                          'used to determine hittest feature.

'--------------------------------------------------------------------------------
'added by M. Costa on 21/06/2002
Private uDataFormat       As String       'the data format for numeric values
Private dblMeanValue      As Double       'the mean value
Private uMeanOn           As Boolean      'marker indicating if the mean value must be displayed
Private uMeanColor        As Long         'the mean line color
Private Const MEAN_CAPTION = "Mean"
Private uMeanCaption      As String       'the mean caption used in the legend
Private uPicture          As StdPicture   'the background picture
Private uPictureTile      As Boolean      'marker indicating if the background picture must be tiled
                                          '(TRUE) or stretched (FALSE)
Private uBarPicture       As StdPicture   'the background picture
Private uBarPictureTile   As Boolean      'marker indicating if the bar picture must be tiled
Private uBarShadow        As Boolean      'marker indicating if the bar must have the shadow
                                          '(shadow takes effect only if line width is 1!)
Private uBarShadowColor   As Long         'the bar shadow color
Private uAutoRedraw       As Boolean      'marker indicating if the chart is auto-redrawn
                                          'upon every property change
Private uRangeY           As Integer      'the absolute range between Y-axis min. ad max. values
Private uDataType         As Integer      'indicates the data distribution in the Y axis
Private Const DT_BOTH = 0                 ' 0 = range(-Y0, +Y1)
Private Const DT_NEG = 1                  ' 1 = range(-Y0, -Y1)
Private Const DT_POS = 2                  ' 2 = range(+Y0, +Y1)

Private uMinYValue        As Double       'minimum y value
Private uLineColor        As Long         'the color of the line
Private uLineStyle        As Integer      'the line style
Private uBarSymbolColor   As Long         'the color of the symbol
Private uBarColor         As Long         'the backcolor of the bars
Private uBarFillStyle     As Integer      'the bars fill style
Private uSelectedBarColor As Long         'the selected bar backcolor
Private uMinorGridColor   As Long         'the minor intersect grid color
Private uMajorGridColor   As Long         'the major intersect grid color
Private uMinorGridOn      As Boolean      'marker indicating display of minor grid
Private uMajorGridOn      As Boolean      'marker indicating display of major grid
Private uLegendBackColor  As Long         'the legend background color
Private uLegendForeColor  As Long         'the legend foreground color
Private uInfoBackColor    As Long         'the information box background color
Private uInfoForeColor    As Long         'the information box foreground color
Private uXAxisLabelColor  As Long         'the X axis label color
Private uYAxisLabelColor  As Long         'the Y axis label color
Private uXAxisItemsColor  As Long         'the X axis items color
Private uYAxisItemsColor  As Long         'the Y axis items color
Private uChartTitleColor  As Long         'the chart title color
Private uChartSubTitleColor As Long       'the chart subtitle color
Private uSaveAsCaption    As String       'the SaveAs dialog box caption
Private uInfoItems        As String       'the information items (to be displayed in the info box)
Private Const INFO_ITEMS = "Value|Description|Mean"

Public Enum ChartMenuConstants             'the enumerated for menu type
    xcPopUpMenu = 0
    xcButtonMenu
End Enum

Private uMenuType         As ChartMenuConstants 'the menu type.
Private uMenuItems        As String       'the menu's items.
Private Const MENU_ITEMS = "&Save as...|&Print|&Copy|Selection &information|&Legend|&Properties|&Hide"

Private Const XC_BAR = 1
Private Const XC_SYMBOL = 2
Private Const XC_LINE = 4
Private Const XC_OVAL = 8
Private Const XC_TRIANGLE = 16
Private Const XC_RHOMBUS = 32
Private Const XC_TRAPEZIUM = 64
Public Enum ChartTypeConstants            'the enumerated for chart type
    xcBar = XC_BAR
    xcSymbol = XC_SYMBOL
    xcLine = XC_LINE
    xcBarLine = XC_BAR + XC_LINE
    xcSymbolLine = XC_SYMBOL + XC_LINE
    xcOval = XC_OVAL
    xcOvalLine = XC_OVAL + XC_LINE
    xcTriangle = XC_TRIANGLE
    xcTriangleLine = XC_TRIANGLE + XC_LINE
    xcRhombus = XC_RHOMBUS
    xcRhombusLine = XC_RHOMBUS + XC_LINE
    xcTrapezium = XC_TRAPEZIUM
    xcTrapeziumLine = XC_TRAPEZIUM + XC_LINE
End Enum

Private uChartType        As ChartTypeConstants 'the chart type.
Private uBarSymbol        As String * 1   'the symbol to be displayed when uChartType=xcSymbol
Private uBarWidthPercentage As Integer    'the column width (in percentage) just for bar type
Private uLineWidth        As Integer      'the line width (used when uChartType=xcLine and for bar border in case of uChartType=xcBar)

Private Const IDX_SAVE = 0                'the command buttons' indexs
Private Const IDX_PRINT = 1
Private Const IDX_COPY = 2
Private Const IDX_INFO = 3
Private Const IDX_LEGEND = 4
Private Const IDX_PROPERTIES = 5
'--------------------------------------------------------------------------------

Private uColWidth         As Single       'the calculated width of each column
Private uRowHeight        As Single       'the calculated height of each column
Private uTopMargin        As Single       '--------------------------------------
Private uBottomMargin     As Single       'margins used around the chart content
Private uLeftMargin       As Single       '
Private uRightMargin      As Single       '--------------------------------------
Private uContentBorder    As Boolean      'border around the chart content?
Private uSelectable       As Boolean      'marker indicating whether user can select a column
Private uHotTracking      As Boolean      'marker indicating use of hot tracking
Private uSelectedColumn   As Integer      'marker indicating the selected column
Private uOldSelection     As Long
Private uDisplayDescript  As Boolean      'display description when selectable
Private uChartTitle       As String       'chart title
Private uChartSubTitle    As String       'chart sub title
Private uAxisXOn          As Boolean      'marker indicating display of x axis
Private uAxisYOn          As Boolean      'marker indicating display of y axis
Private uColorBars        As Boolean      'marker indicating use of different coloured bars
Private uIntersectMajor   As Single       'major intersect value
Private uIntersectMinor   As Single       'minor intersect value
Private uMaxYValue        As Double       'maximum y value
Private uXAxisLabel       As String       'label to be displayed below the X-Axis
Private uYAxisLabel       As String       'label to be displayed left of the Y-Axis
Private cItems            As Collection   'collection of chart items

Private offsetX           As Long
Private offsetY           As Long

Private bLegendAdded      As Boolean
Private bLegendClicked    As Boolean
Private bDisplayLegend    As Boolean
Private bResize           As Boolean

Private bProcessingOver   As Boolean      'marker to speed up mouse over effects

Public Type ChartItem
    ItemID As String
    SelectedDescription As String
    XAxisDescription As String
    Value As Double
End Type

Public Event ItemClick(cItem As ChartItem)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Function AddItem(cItem As ChartItem) As Boolean
    
    Dim oChartItem As ChartItem
    
    If uMeanOn = True Then
        If cItems.Count > 0 Then
            cItems.Remove (cItems.Count)
        End If
    End If

    cItems.Add cItem
    CalcMean
    
    If uMeanOn = True Then
        If uMeanCaption = Empty Then uMeanCaption = MEAN_CAPTION
        oChartItem.Value = dblMeanValue
        oChartItem.ItemID = uMeanCaption
        oChartItem.XAxisDescription = uMeanCaption
        oChartItem.SelectedDescription = uMeanCaption
        cItems.Add oChartItem
    End If
    
End Function

Public Property Let AutoRedraw(blnVal As Boolean)
    If blnVal <> uAutoRedraw Then
        uAutoRedraw = blnVal
        DrawChart
        PropertyChanged "AutoRedraw"
    End If
End Property

Public Property Get AutoRedraw() As Boolean
    AutoRedraw = uAutoRedraw
End Property

Public Property Get BarShadow() As Boolean
    BarShadow = uBarShadow
End Property

Public Property Get BarShadowColor() As OLE_COLOR
    BarShadowColor = uBarShadowColor
End Property


Public Property Let BarShadow(blnVal As Boolean)
    If blnVal <> uBarShadow Then
        uBarShadow = blnVal
        DrawChart
        PropertyChanged "BarShadow"
    End If
End Property
Public Property Let BarShadowColor(lngVal As OLE_COLOR)
    If lngVal <> uBarShadowColor Then
        uBarShadowColor = lngVal
        DrawChart
        PropertyChanged "BarShadowColor"
    End If
End Property

Private Sub CalcMean()
    
    On Error Resume Next
    
    Dim intIdx As Integer
    
    dblMeanValue = 0
    For intIdx = 0 To cItems.Count - 1
        dblMeanValue = dblMeanValue + cItems.Item(intIdx).Value
    Next
    dblMeanValue = dblMeanValue / cItems.Count
    
End Sub

Public Property Get DataFormat() As String
    DataFormat = uDataFormat
End Property

Public Property Let DataFormat(stgVal As String)
    uDataFormat = stgVal
    PropertyChanged "DataFormat"
End Property

Private Sub DisplayInfo(intIdx As Integer)

    Dim sDescription    As String
    Dim varItems        As Variant
    Dim oChartItem      As ChartItem
    
    'it's important to let the info label invisible at beginning to avoid flickering effect
    lblInfo.Visible = False
    If uDisplayDescript Then
        If intIdx > -1 Then
            oChartItem = cItems.Item(intIdx + 1)
            'this kind of error trapping is useful in case the user
            'did not define any item in the menu items string, so the default is used
            On Error GoTo DrawChart_error
    
            If uInfoItems = Empty Then uInfoItems = INFO_ITEMS
            varItems = Split(uInfoItems, "|")
            sDescription = CStr(varItems(0)) & ": " & Format(oChartItem.Value, uDataFormat)
            If Len(oChartItem.SelectedDescription) > 0 Then
                sDescription = CStr(varItems(1)) & ": " & oChartItem.SelectedDescription & vbCrLf & sDescription
            End If
            If (uMeanOn = True) And (intIdx < cItems.Count - 1) Then
                sDescription = sDescription & vbCrLf & CStr(varItems(2)) & ": " & Format(dblMeanValue, uDataFormat)
            End If
        End If
        If sDescription <> Empty Then
            lblInfo.Caption = sDescription
            lblInfo.Width = UserControl.TextWidth(sDescription) + 5 * Screen.TwipsPerPixelX
            lblInfo.Height = UserControl.TextHeight(sDescription) * 1.2
            lblInfo.Visible = True
        End If
    End If
    Exit Sub

DrawChart_error:
    uInfoItems = INFO_ITEMS
    Resume Next

End Sub

Private Sub DrawOval(sngX1 As Single, sngX2 As Single, sngY1 As Single, sngY2 As Single, sngBase As Single, sngHeight As Single, lngBorderColor As Long)
    
    On Error Resume Next
    
    Dim x1 As Single
    Dim y1 As Single
    Dim x2 As Single
    Dim y2 As Single
    Dim sngH As Single
    Dim sngW As Single
    Dim lngFillColor As Long

    x1 = sngX1
    y1 = sngY1
    x2 = sngX2
    y2 = sngY2
    sngW = sngBase
    sngH = sngHeight
    x1 = x1 + (sngW / 2)
    y1 = y1 + (sngH / 2)
    If uBarShadow = True Then
        lngFillColor = UserControl.FillColor
        UserControl.FillStyle = vbFSSolid
        UserControl.FillColor = uBarShadowColor
        UserControl.Circle (x1, y1), sngH / 2, uBarShadowColor, , , _
                            IIf((sngH > sngW), (sngH / sngW), (sngW / sngH))
        UserControl.FillColor = lngFillColor
        UserControl.FillStyle = uBarFillStyle
        x1 = x1 - 2 * Screen.TwipsPerPixelX
        sngW = sngW - 2 * Screen.TwipsPerPixelX
        sngH = sngH - 2 * Screen.TwipsPerPixelX
    End If
    'the aspect ratio depend on whether the base is greater than the height
    UserControl.Circle (x1, y1), sngH / 2, lngBorderColor, , , _
                        IIf((sngH > sngW), (sngH / sngW), (sngW / sngH))
    
End Sub

Private Sub DrawPicture(sngX1 As Single, sngX2 As Single, sngY1 As Single, sngY2 As Single, blnTile As Boolean, pic As StdPicture)

    On Error Resume Next
    
    Dim x1 As Single
    Dim x2 As Single
    Dim y1 As Single
    Dim y2 As Single
    Dim sngH As Single
    Dim sngW As Single
    Dim xTemp As Single
    Dim yTemp As Single
    
    If blnTile = True Then
        'I found the ratio of 1.75 to adjust size, but I really don't know why!!!
        sngH = Round(pic.Height / 1.75)
        sngW = Round(pic.Width / 1.75)
        If (sngH Mod Screen.TwipsPerPixelY) <> 0 Then
            sngH = Round(sngH / Screen.TwipsPerPixelY) * Screen.TwipsPerPixelY
        End If
        If (sngW Mod Screen.TwipsPerPixelX) <> 0 Then
            sngW = Round(sngW / Screen.TwipsPerPixelX) * Screen.TwipsPerPixelX
        End If
        y1 = sngY1
        y2 = sngY2
        x2 = sngX2
        Do While y1 < y2
            x1 = sngX1
            Do While x1 < x2
                If (x1 + sngW) > x2 Then
                    xTemp = (x2 - x1)
                Else
                    xTemp = sngW
                End If
                xTemp = IIf(xTemp < Screen.TwipsPerPixelX, Screen.TwipsPerPixelX, xTemp)
                If (y1 + sngH) > y2 Then
                    yTemp = (y2 - y1)
                Else
                    yTemp = sngH
                End If
                yTemp = IIf(yTemp < Screen.TwipsPerPixelY, Screen.TwipsPerPixelY, yTemp)
'If (yTemp Mod Screen.TwipsPerPixelY) <> 0 Then
'    yTemp = Round(yTemp / Screen.TwipsPerPixelY) * Screen.TwipsPerPixelY
'End If
'If (xTemp Mod Screen.TwipsPerPixelX) <> 0 Then
'    xTemp = Round(xTemp / Screen.TwipsPerPixelX) * Screen.TwipsPerPixelX
'End If
                UserControl.PaintPicture pic, _
                            x1, y1, _
                            xTemp, _
                            yTemp, _
                            0, 0, xTemp, yTemp
                x1 = (x1 + sngW)
            Loop
            y1 = (y1 + sngH)
        Loop
    Else
        'stretch the picture
        UserControl.PaintPicture pic, _
                            sngX1, sngY1, _
                            IIf((sngX2 - sngX1) < Screen.TwipsPerPixelX, Screen.TwipsPerPixelX, (sngX2 - sngX1)), _
                            IIf((sngY2 - sngY1) < Screen.TwipsPerPixelY, Screen.TwipsPerPixelY, (sngY2 - sngY1))
    End If

End Sub

Private Sub DrawRectangle(dblData As Double, sngX1 As Single, sngX2 As Single, sngY1 As Single, sngY2 As Single, lngBorderColor As Long, blnOverridePicture As Boolean)
        
    On Error Resume Next
    
    Dim x1 As Single
    Dim y1 As Single
    Dim x2 As Single
    Dim y2 As Single
    
    x1 = sngX1
    y1 = sngY1
    x2 = sngX2
    y2 = sngY2
    If uBarShadow = True Then
        x2 = x2 - 2 * Screen.TwipsPerPixelX
    End If
    If (blnOverridePicture = True) Or (uBarPicture Is Nothing) Then
        UserControl.Line (x1 + 1 * Screen.TwipsPerPixelX, y1)-(x2 - 1 * Screen.TwipsPerPixelX, y2), , BF
    Else
        Call DrawPicture(x1, x2, y1, y2, uBarPictureTile, uBarPicture)
        'if the fill  style is solid, the image is overriden when drawing the outer box
        If UserControl.FillStyle = vbFSSolid Then _
            UserControl.FillStyle = vbFSTransparent
    End If
    UserControl.Line (x1, y1)-(x2 - 1 * Screen.TwipsPerPixelX, y2), lngBorderColor, B
    If uBarShadow = True Then
        If dblData >= 0 Then
            y1 = y1 + 2 * Screen.TwipsPerPixelX
            UserControl.Line (x2, y1)-(x2 + 2 * Screen.TwipsPerPixelX, y2), uBarShadowColor, BF
        Else
            y2 = y2 - 2 * Screen.TwipsPerPixelX
            UserControl.Line (x2, y1)-(x2 + 2 * Screen.TwipsPerPixelX, y2), uBarShadowColor, BF
        End If
    End If
    UserControl.FillStyle = uBarFillStyle

End Sub

Private Sub DrawRhombus(dblData As Double, sngX1 As Single, sngX2 As Single, sngY1 As Single, sngY2 As Single)

    On Error Resume Next
    
    Dim lRet As Long
    Dim sngXTemp As Single
    Dim sngYTemp As Single
    Dim uaPts(3) As PointAPI
    Dim lngFillColor As Long
    Dim intScaleMode As Integer
    
    'the polygon function works only with pixels!
    intScaleMode = UserControl.ScaleMode
    UserControl.ScaleMode = vbPixels
    
    'setup the 4 points of the Rhombus anti-clockwise
    '     (1)
    '    /   \
    '   /     \
    ' (0)     (2)
    '   \     /
    '    \   /
    '     (3)
    sngXTemp = sngX1 + ((sngX2 - sngX1) / 2)
    sngYTemp = sngY1 + ((sngY2 - sngY1) / 2)
    uaPts(0).X = sngX1 / Screen.TwipsPerPixelX
    uaPts(0).Y = sngYTemp / Screen.TwipsPerPixelY
    uaPts(1).X = sngXTemp / Screen.TwipsPerPixelX
    uaPts(1).Y = sngY1 / Screen.TwipsPerPixelY
    uaPts(2).X = sngX2 / Screen.TwipsPerPixelX
    uaPts(2).Y = sngYTemp / Screen.TwipsPerPixelY
    uaPts(3).X = sngXTemp / Screen.TwipsPerPixelX
    uaPts(3).Y = sngY2 / Screen.TwipsPerPixelY
    
    If uBarShadow = True Then
        lngFillColor = UserControl.FillColor
        UserControl.FillColor = uBarShadowColor
        lRet = Polygon(UserControl.hDC, uaPts(0), 4)
        UserControl.FillColor = lngFillColor
        'resize the Rhombus
        uaPts(1).X = uaPts(1).X - 2
        uaPts(2).X = uaPts(2).X - 3
        uaPts(3).X = uaPts(3).X - 2
        If dblData > 0 Then
            uaPts(1).Y = uaPts(1).Y + 2
            uaPts(3).Y = uaPts(3).Y - 2
        Else
            uaPts(1).Y = uaPts(1).Y - 2
            uaPts(3).Y = uaPts(3).Y + 2
        End If
    End If
    
    'draw the filled Rhombus
    lRet = Polygon(UserControl.hDC, uaPts(0), 4)
    
    'reset the scalemode
    UserControl.ScaleMode = intScaleMode
    
    'free the memory
    Erase uaPts

End Sub

Private Sub DrawTrapezium(dblData As Double, sngX1 As Single, sngX2 As Single, sngY1 As Single, sngY2 As Single)

    On Error Resume Next
    
    Dim lRet As Long
    Dim sngXTemp As Single
    Dim sngYTemp As Single
    Dim lngFillColor As Long
    Dim uaPts(3) As PointAPI
    Dim intScaleMode As Integer
    
    'the polygon function works only with pixels!
    intScaleMode = UserControl.ScaleMode
    UserControl.ScaleMode = vbPixels
    
    'setup the 4 points of the trapezio
    sngXTemp = (sngX2 - sngX1) / 4      'consider the 25% as X-offset
    'set the points anti-clockwise
    '     (1)-----(2)
    '    /           \
    '   /             \
    ' (0)-------------(3)
    uaPts(0).X = sngX1 / Screen.TwipsPerPixelX
    uaPts(1).X = (sngX1 + sngXTemp) / Screen.TwipsPerPixelX
    uaPts(2).X = (sngX2 - sngXTemp) / Screen.TwipsPerPixelX
    uaPts(3).X = sngX2 / Screen.TwipsPerPixelX
    If dblData > 0 Then
        uaPts(0).Y = sngY2 / Screen.TwipsPerPixelY
        uaPts(1).Y = sngY1 / Screen.TwipsPerPixelY
        uaPts(2).Y = sngY1 / Screen.TwipsPerPixelY
        uaPts(3).Y = sngY2 / Screen.TwipsPerPixelY
    Else
        uaPts(0).Y = sngY1 / Screen.TwipsPerPixelY
        uaPts(1).Y = sngY2 / Screen.TwipsPerPixelY
        uaPts(2).Y = sngY2 / Screen.TwipsPerPixelY
        uaPts(3).Y = sngY1 / Screen.TwipsPerPixelY
    End If
    
    If uBarShadow = True Then
        lngFillColor = UserControl.FillColor
        UserControl.FillColor = uBarShadowColor
        lRet = Polygon(UserControl.hDC, uaPts(0), 4)
        UserControl.FillColor = lngFillColor
        'resize the trapezio
        uaPts(1).X = uaPts(1).X - 2
        uaPts(2).X = uaPts(2).X - 2
        uaPts(3).X = uaPts(3).X - 2
        If dblData > 0 Then
            uaPts(1).Y = uaPts(1).Y + 2
            uaPts(2).Y = uaPts(2).Y + 2
        Else
            uaPts(1).Y = uaPts(1).Y - 2
            uaPts(2).Y = uaPts(2).Y - 2
        End If
    End If
    
    'draw the filled trapezio
    lRet = Polygon(UserControl.hDC, uaPts(0), 4)
    
    'reset the scalemode
    UserControl.ScaleMode = intScaleMode
    
    'free the memory
    Erase uaPts

End Sub


Private Sub DrawTriangle(dblData As Double, sngX1 As Single, sngX2 As Single, sngY1 As Single, sngY2 As Single)

    On Error Resume Next
    
    Dim lRet As Long
    Dim uaPts(2) As PointAPI
    Dim lngFillColor As Long
    Dim intScaleMode As Integer

    'the polygon function works only with pixels!
    intScaleMode = UserControl.ScaleMode
    UserControl.ScaleMode = vbPixels
    
    'setup the 3 points of the triangle anti-clockwise
    '     (1)
    '    /   \
    '   /     \
    ' (0)-----(2)
    uaPts(0).X = sngX1 / Screen.TwipsPerPixelX
    uaPts(1).X = sngX2 / Screen.TwipsPerPixelX
    uaPts(2).X = (sngX1 + ((sngX2 - sngX1) / 2)) / Screen.TwipsPerPixelX
    If dblData > 0 Then
        uaPts(0).Y = sngY2 / Screen.TwipsPerPixelY
        uaPts(1).Y = sngY2 / Screen.TwipsPerPixelY
        uaPts(2).Y = sngY1 / Screen.TwipsPerPixelY
    Else
        uaPts(0).Y = sngY1 / Screen.TwipsPerPixelY
        uaPts(1).Y = sngY1 / Screen.TwipsPerPixelY
        uaPts(2).Y = sngY2 / Screen.TwipsPerPixelY
    End If
    
    If uBarShadow = True Then
        lngFillColor = UserControl.FillColor
        UserControl.FillColor = uBarShadowColor
        lRet = Polygon(UserControl.hDC, uaPts(0), 3)
        UserControl.FillColor = lngFillColor
        'resize the triangle
        uaPts(1).X = uaPts(1).X - 2
        uaPts(2).X = uaPts(2).X - 2
        If dblData > 0 Then
            uaPts(2).Y = uaPts(2).Y + 2
        Else
            uaPts(2).Y = uaPts(2).Y - 2
        End If
    End If
    
    'draw the filled triangle
    lRet = Polygon(UserControl.hDC, uaPts(0), 3)
    
    'reset the scalemode
    UserControl.ScaleMode = intScaleMode
    
    'free the memory
    Erase uaPts

End Sub

Public Property Get LineStyle() As DrawStyleConstants
    LineStyle = uLineStyle
End Property

Public Property Let LineStyle(intVal As DrawStyleConstants)
    If uLineStyle <> intVal Then
        uLineStyle = intVal
        DrawChart
        PropertyChanged "LineStyle"
    End If
End Property

Public Property Let LineWidth(intVal As Integer)
    If intVal <> uLineWidth Then
        If intVal > 0 And intVal <= 9 Then
            uLineWidth = intVal
            DrawChart
            PropertyChanged "LineWidth"
        End If
    End If
End Property

Public Property Get LineWidth() As Integer
    LineWidth = uLineWidth
End Property
Public Property Get MeanOn() As Boolean
    MeanOn = uMeanOn
End Property

Public Property Get MeanCaption() As String
    MeanCaption = uMeanCaption
End Property


Public Property Get MeanColor() As OLE_COLOR
    MeanColor = uMeanColor
End Property


Public Property Let MeanOn(blnVal As Boolean)
    If blnVal <> uMeanOn Then
        uMeanOn = blnVal
        DrawChart
        PropertyChanged "MeanOn"
    End If
End Property

Public Property Let MeanCaption(stgVal As String)
    If stgVal <> uMeanCaption Then
        uMeanCaption = stgVal
        DrawChart
        PropertyChanged "MeanCaption"
    End If
End Property


Public Property Let MeanColor(lngVal As OLE_COLOR)
    If lngVal <> uMeanColor Then
        uMeanColor = lngVal
        DrawChart
        PropertyChanged "MeanColor"
    End If
End Property


Public Property Get MinorGridOn() As Boolean
    MinorGridOn = uMinorGridOn
End Property
Public Property Get MajorGridOn() As Boolean
    MajorGridOn = uMajorGridOn
End Property

Public Property Let MinorGridOn(blnVal As Boolean)
    If blnVal <> uMinorGridOn Then
        uMinorGridOn = blnVal
        DrawChart
        PropertyChanged "MinorGridOn"
    End If
End Property

Public Property Let MajorGridOn(blnVal As Boolean)
    If blnVal <> uMajorGridOn Then
        uMajorGridOn = blnVal
        DrawChart
        PropertyChanged "MajorGridOn"
    End If
End Property

Public Property Get MinorGrid() As Boolean

End Property

Public Property Set Picture(ByVal picVal As StdPicture)
    Set uPicture = picVal
    DrawChart
End Property


Public Property Set BarPicture(ByVal picVal As StdPicture)
    Set uBarPicture = picVal
    DrawChart
End Property



Public Property Get Picture() As Picture
    Set Picture = uPicture
End Property
Public Property Get BarPicture() As Picture
    Set BarPicture = uBarPicture
End Property

Public Property Get BarWidthPercentage() As Integer
    BarWidthPercentage = uBarWidthPercentage
End Property

Public Property Get BarSymbol() As String
    BarSymbol = uBarSymbol
End Property

Public Property Let BarSymbol(stgVal As String)
    If stgVal <> uBarSymbol Then
        uBarSymbol = stgVal
        DrawChart
        PropertyChanged "BarSymbol"
    End If
End Property

Public Property Let ChartType(intVal As ChartTypeConstants)
    If intVal <> uChartType Then
        uChartType = intVal
        DrawChart
        PropertyChanged "ChartType"
    End If
End Property

Public Property Let BarWidthPercentage(intVal As Integer)
    If intVal > 0 And intVal <= 100 Then
        If intVal <> uBarWidthPercentage Then
            uBarWidthPercentage = intVal
            DrawChart
            PropertyChanged "BarWidthPercentage"
        End If
    End If
End Property
Public Property Get ChartType() As ChartTypeConstants
    ChartType = uChartType
End Property
Public Function EditCopy() As Boolean
    Clipboard.SetData UserControl.Image
End Function

Private Sub FixData()

    If uMinYValue < 0 And uMaxYValue < 0 Then
        uDataType = DT_NEG
        uRangeY = (Abs(uMinYValue) - Abs(uMaxYValue))
    ElseIf uMinYValue >= 0 And uMaxYValue >= 0 Then
        uDataType = DT_POS
        uRangeY = (Abs(uMaxYValue) - Abs(uMinYValue))
    Else
        uDataType = DT_BOTH
        uRangeY = (Abs(uMaxYValue) + Abs(uMinYValue))
    End If

    If uRangeY = 0 Then uRangeY = 1
    If uIntersectMajor = 0 Then uIntersectMajor = uRangeY / 10
    If uIntersectMinor = 0 Then uIntersectMinor = uIntersectMajor / 5
    
    
End Sub

Private Sub FixMenu()
    
    'this kind of error trapping is useful in case the user
    'did not define any item in the menu items string, so the default is used
    On Error GoTo FixMenu_error
    
    Dim varItems As Variant
    
    If uMenuItems = Empty Then
        uMenuItems = MENU_ITEMS
    End If
    varItems = Split(uMenuItems, "|")
    
    If varItems(0) <> Empty Then
        mnuSaveAs.Caption = CStr(varItems(0))
    Else
        mnuSaveAs.Caption = "&Save as..."
    End If
    cmdCmd(IDX_SAVE).ToolTipText = Replace(mnuSaveAs.Caption, "&", "")
    
    If varItems(1) <> Empty Then
        mnuPrint.Caption = CStr(varItems(1))
    Else
        mnuPrint.Caption = "&Print"
    End If
    cmdCmd(IDX_PRINT).ToolTipText = Replace(mnuPrint.Caption, "&", "")
    
    If varItems(2) <> Empty Then
        mnuEditCopy.Caption = CStr(varItems(2))
    Else
        mnuEditCopy.Caption = "&Copy"
    End If
    cmdCmd(IDX_COPY).ToolTipText = Replace(mnuEditCopy.Caption, "&", "")
    
    If varItems(3) <> Empty Then
        mnuSelectionInfo.Caption = CStr(varItems(3))
    Else
        mnuSelectionInfo.Caption = "Selection &information"
    End If
    cmdCmd(IDX_INFO).ToolTipText = Replace(mnuSelectionInfo.Caption, "&", "")
    
    If varItems(4) <> Empty Then
        mnuViewLegend.Caption = CStr(varItems(4))
    Else
        mnuViewLegend.Caption = "&Legend"
    End If
    cmdCmd(IDX_LEGEND).ToolTipText = Replace(mnuViewLegend.Caption, "&", "")
    
    If varItems(5) <> Empty Then
        mnuProperties.Caption = CStr(varItems(5))
    Else
        mnuProperties.Caption = "&Properties"
    End If
    
    If varItems(6) <> Empty Then
        mnuLegendHide.Caption = CStr(varItems(6))
    Else
        mnuLegendHide.Caption = "&Hide"
    End If
    
    If uMenuType = xcButtonMenu Then
        picCommands.Visible = True
        picCommands.BackColor = UserControl.BackColor
        picCommands.Move 60, 60
        lblInfo.Move picCommands.Left + picCommands.ScaleWidth + 60, 60
    Else
        picCommands.Visible = False
        lblInfo.Move 60, 60
    End If
    Exit Sub
    
FixMenu_error:
    uMenuItems = MENU_ITEMS
    Resume Next

End Sub

Private Function InColumn(X As Single, Y As Single) As Integer

    Dim sngY As Single
    Dim sngY1 As Single
    Dim sngY2 As Single
    Dim intCol As Integer
    Dim intSelectedCol As Integer

    intSelectedCol = -1
    If (uChartType And XC_BAR) = XC_BAR _
    Or (uChartType And XC_OVAL) = XC_OVAL _
    Or (uChartType And XC_RHOMBUS) = XC_RHOMBUS _
    Or (uChartType And XC_TRAPEZIUM) = XC_TRAPEZIUM _
    Or (uChartType And XC_TRIANGLE) = XC_TRIANGLE Then
        If (Y <= UserControl.ScaleHeight - uBottomMargin) And (Y >= uTopMargin) _
        And (uSelectable = True) Then
            intCol = (X - uLeftMargin) \ (uColWidth)
            sngY1 = uColumns(intCol, 0)
            sngY2 = uColumns(intCol, 1)
            If sngY1 > sngY2 Then
                sngY = sngY1
                sngY1 = sngY2
                sngY2 = sngY
            End If
            If (Y >= sngY1 And Y <= sngY2) Then
                intSelectedCol = intCol
            End If
        End If
    End If
    InColumn = intSelectedCol

End Function

Public Property Let MarginTop(lMargin As Long)
    uTopMargin = lMargin * Screen.TwipsPerPixelY
    DrawChart
    PropertyChanged "MarginTop"
End Property
Public Property Get MarginTop() As Long
    MarginTop = uTopMargin / Screen.TwipsPerPixelY
End Property

Public Property Let MarginBottom(lMargin As Long)
    uBottomMargin = lMargin * Screen.TwipsPerPixelY
    DrawChart
    PropertyChanged "MarginBottom"
End Property
Public Property Get MarginBottom() As Long
    MarginBottom = uBottomMargin / Screen.TwipsPerPixelY
End Property

Public Property Let MarginLeft(lMargin As Long)
    uLeftMargin = lMargin * Screen.TwipsPerPixelX
    DrawChart
    PropertyChanged "MarginLeft"
End Property
Public Property Get MarginLeft() As Long
    MarginLeft = uLeftMargin / Screen.TwipsPerPixelX
End Property

Public Property Let MarginRight(lMargin As Long)
    uRightMargin = lMargin * Screen.TwipsPerPixelX
    DrawChart
    PropertyChanged "MarginRight"
End Property
Public Property Get MarginRight() As Long
    MarginRight = uRightMargin / Screen.TwipsPerPixelX
End Property

Public Property Let ContentBorder(blnVal As Boolean)
    If blnVal <> uContentBorder Then
        uContentBorder = blnVal
        DrawChart
        PropertyChanged "ContentBorder"
    End If
End Property
Public Property Get ContentBorder() As Boolean
    ContentBorder = uContentBorder
End Property

Public Property Get MenuType() As ChartMenuConstants
    MenuType = uMenuType
End Property

Public Property Let MenuType(intVal As ChartMenuConstants)
    If intVal <> uMenuType Then
        uMenuType = intVal
        FixMenu
        PropertyChanged "MenuType"
    End If
End Property
Public Property Let PictureTile(blnVal As Boolean)
    If blnVal <> uPictureTile Then
        uPictureTile = blnVal
        DrawChart
        PropertyChanged "PictureTile"
    End If
End Property

Public Property Let BarPictureTile(blnVal As Boolean)
    If blnVal <> uBarPictureTile Then
        uBarPictureTile = blnVal
        DrawChart
        PropertyChanged "BarPictureTile"
    End If
End Property


Public Property Get PictureTile() As Boolean
    PictureTile = uPictureTile
End Property
Public Property Get BarPictureTile() As Boolean
    BarPictureTile = uBarPictureTile
End Property

Public Property Let Selectable(blnVal As Boolean)
    If blnVal <> uSelectable Then
        uSelectable = blnVal
        DrawChart
        PropertyChanged "Selectable"
    End If
End Property
Public Property Get Selectable() As Boolean
    Selectable = uSelectable
End Property

Public Property Let HotTracking(blnVal As Boolean)
    If blnVal <> uHotTracking Then
        uHotTracking = blnVal
        DrawChart
        PropertyChanged "HotTracking"
    End If
End Property
Public Property Get HotTracking() As Boolean
    HotTracking = uHotTracking
End Property

Public Property Let SelectedColumn(lngColumn As Long)
    
    Dim oItem As ChartItem
    On Error Resume Next
    
    If lngColumn <> uSelectedColumn Then
        uSelectedColumn = lngColumn
        DrawChart
        PropertyChanged "SelectedColumn"
        
        If Err.Number Then
            uSelectedColumn = -1
        Else
            If (uMeanOn = True) And (uSelectedColumn = cItems.Count - 1) Then
                'do nothing in case of mean bar selected
            Else
                oItem = cItems(lngColumn + 1)
                RaiseEvent ItemClick(oItem)
            End If
        End If
    End If

End Property
Public Property Get SelectedColumn() As Long
    SelectedColumn = uSelectedColumn
End Property

Public Property Let ChartTitle(stgVal As String)
    If stgVal <> uChartTitle Then
        uChartTitle = stgVal
        DrawChart
        PropertyChanged "ChartTitle"
    End If
End Property
Public Property Get ChartTitle() As String
    ChartTitle = uChartTitle
End Property
Public Property Let MenuItems(stgVal As String)
    uMenuItems = stgVal
    FixMenu
    PropertyChanged "MenuItems"
End Property
Public Property Let InfoItems(stgVal As String)
    uInfoItems = stgVal
    PropertyChanged "InfoItems"
End Property

Public Property Get InfoItems() As String
    InfoItems = uInfoItems
End Property
Public Property Get MenuItems() As String
    MenuItems = uMenuItems
End Property

Public Property Let ChartSubTitle(stgVal As String)
    If stgVal <> uChartSubTitle Then
        uChartSubTitle = stgVal
        DrawChart
        PropertyChanged "ChartSubTitle"
    End If
End Property
Public Property Get ChartSubTitle() As String
    ChartSubTitle = uChartSubTitle
End Property

Public Property Let IntersectMajor(sngVal As Single)
    If sngVal <> uIntersectMajor Then
        uIntersectMajor = sngVal
        DrawChart
        PropertyChanged "IntersectMajor"
    End If
End Property
Public Property Get IntersectMajor() As Single
    IntersectMajor = uIntersectMajor
End Property

Public Property Let IntersectMinor(sngVal As Single)
    If sngVal <> uIntersectMinor Then
        uIntersectMinor = sngVal
        DrawChart
        PropertyChanged "IntersectMinor"
    End If
End Property
Public Property Get IntersectMinor() As Single
    IntersectMinor = uIntersectMinor
End Property

Public Property Let AxisYOn(blnVal As Boolean)
    If blnVal <> uAxisYOn Then
        uAxisYOn = blnVal
        DrawChart
        PropertyChanged "AxisYOn"
    End If
End Property
Public Property Get AxisYOn() As Boolean
    AxisYOn = uAxisYOn
End Property

Public Property Let AxisXOn(blnVal As Boolean)
    If blnVal <> uAxisXOn Then
        uAxisXOn = blnVal
        DrawChart
        PropertyChanged "AxisXOn"
    End If
End Property
Public Property Get AxisXOn() As Boolean
    AxisXOn = uAxisXOn
End Property

Public Property Let MaxY(dblMax As Double)
    If dblMax > uMinYValue Then
        uMaxYValue = dblMax
        DrawChart
        PropertyChanged "MaxY"
    End If
End Property
Public Property Let MinY(dblMin As Double)
    If dblMin < uMaxYValue Then
        uMinYValue = dblMin
        DrawChart
        PropertyChanged "MinY"
    End If
End Property

Public Property Get MinY() As Double
    MinY = uMinYValue
End Property


Public Property Get MaxY() As Double
    MaxY = uMaxYValue
End Property

Public Property Let SelectionInformation(blnVal As Boolean)
    If blnVal <> uDisplayDescript Then
        uDisplayDescript = blnVal
        DrawChart
        PropertyChanged "SelectionInformation"
    End If
End Property
Public Property Get SelectionInformation() As Boolean
    SelectionInformation = uDisplayDescript
End Property

Public Property Let AxisLabelY(stgCaption As String)
    If stgCaption <> uYAxisLabel Then
        uYAxisLabel = stgCaption
        DrawChart
        PropertyChanged "AxisLabelY"
    End If
End Property
Public Property Get AxisLabelY() As String
    AxisLabelY = uYAxisLabel
End Property

Public Property Let AxisLabelX(stgCaption As String)
    If stgCaption <> uXAxisLabel Then
        uXAxisLabel = stgCaption
        DrawChart
        PropertyChanged "AxisLabelX"
    End If
End Property
Public Property Let AxisLabelXColor(lngVal As OLE_COLOR)
    If lngVal <> uXAxisLabelColor Then
        uXAxisLabelColor = lngVal
        DrawChart
        PropertyChanged "AxisLabelXColor"
    End If
End Property

Public Property Let AxisLabelYColor(lngVal As OLE_COLOR)
    If lngVal <> uYAxisLabelColor Then
        uYAxisLabelColor = lngVal
        DrawChart
        PropertyChanged "AxisLabelYColor"
    End If
End Property


Public Property Let AxisItemsYColor(lngVal As OLE_COLOR)
    If lngVal <> uYAxisItemsColor Then
        uYAxisItemsColor = lngVal
        DrawChart
        PropertyChanged "AxisItemsYColor"
    End If
End Property



Public Property Let AxisItemsXColor(lngVal As OLE_COLOR)
    If lngVal <> uXAxisItemsColor Then
        uXAxisItemsColor = lngVal
        DrawChart
        PropertyChanged "AxisItemsXColor"
    End If
End Property
Public Property Get AxisItemsYColor() As OLE_COLOR
    AxisItemsYColor = uYAxisItemsColor
End Property
Public Property Get AxisItemsXColor() As OLE_COLOR
    AxisItemsXColor = uXAxisItemsColor
End Property

Public Property Get AxisLabelYColor() As OLE_COLOR
    AxisLabelYColor = uYAxisLabelColor
End Property



Public Property Get AxisLabelXColor() As OLE_COLOR
    AxisLabelXColor = uXAxisLabelColor
End Property




Public Property Get AxisLabelX() As String
    AxisLabelX = uXAxisLabel
End Property

Public Property Let BackColor(lngVal As OLE_COLOR)
    If lngVal <> UserControl.BackColor Then
        UserControl.BackColor = lngVal
        DrawChart
        PropertyChanged "BackColor"
    End If
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Get MajorGridColor() As OLE_COLOR
    MajorGridColor = uMajorGridColor
End Property

Public Property Get ChartTitleColor() As OLE_COLOR
    ChartTitleColor = uChartTitleColor
End Property
Public Property Get SaveAsCaption() As String
    SaveAsCaption = uSaveAsCaption
End Property
Public Property Let SaveAsCaption(stgVal As String)
    uSaveAsCaption = stgVal
    PropertyChanged "SaveAsCaption"
End Property
Public Property Let ChartTitleColor(lngVal As OLE_COLOR)
    If lngVal <> uChartTitleColor Then
        uChartTitleColor = lngVal
        DrawChart
        PropertyChanged "ChartTitleColor"
    End If
End Property
Public Property Let ChartSubTitleColor(lngVal As OLE_COLOR)
    If lngVal <> uChartSubTitleColor Then
        uChartSubTitleColor = lngVal
        DrawChart
        PropertyChanged "ChartSubTitleColor"
    End If
End Property

Public Property Get ChartSubTitleColor() As OLE_COLOR
    ChartSubTitleColor = uChartSubTitleColor
End Property

Public Property Get MinorGridColor() As OLE_COLOR
    MinorGridColor = uMinorGridColor
End Property

Public Property Let MinorGridColor(lngVal As OLE_COLOR)
    If lngVal <> uMinorGridColor Then
        uMinorGridColor = lngVal
        DrawChart
        PropertyChanged "MinorGridColor"
    End If
End Property


Public Property Let MajorGridColor(lngVal As OLE_COLOR)
    If lngVal <> uMajorGridColor Then
        uMajorGridColor = lngVal
        DrawChart
        PropertyChanged "MajorGridColor"
    End If
End Property



Public Property Get BarColor() As OLE_COLOR
    BarColor = uBarColor
End Property

Public Property Get LegendBackColor() As OLE_COLOR
    LegendBackColor = uLegendBackColor
End Property


Public Property Get LegendForeColor() As OLE_COLOR
    LegendForeColor = uLegendForeColor
End Property



Public Property Let LegendForeColor(lngVal As OLE_COLOR)
    If lngVal <> uLegendForeColor Then
        uLegendForeColor = lngVal
        DrawChart
        PropertyChanged "LegendForeColor"
    End If
End Property




Public Property Let InfoBackColor(lngVal As OLE_COLOR)
    If lngVal <> uInfoBackColor Then
        uInfoBackColor = lngVal
        DrawChart
        PropertyChanged "InfoBackColor"
    End If
End Property
Public Property Let InfoForeColor(lngVal As OLE_COLOR)
    If lngVal <> uInfoForeColor Then
        uInfoForeColor = lngVal
        DrawChart
        PropertyChanged "InfoForeColor"
    End If
End Property

Public Property Get InfoBackColor() As OLE_COLOR
    InfoBackColor = uInfoBackColor
End Property

Public Property Get InfoForeColor() As OLE_COLOR
    InfoForeColor = uInfoForeColor
End Property

Public Property Let LegendBackColor(lngVal As OLE_COLOR)
    If lngVal <> uLegendBackColor Then
        uLegendBackColor = lngVal
        DrawChart
        PropertyChanged "LegendBackColor"
    End If
End Property

Public Property Get SelectedBarColor() As OLE_COLOR
    SelectedBarColor = uSelectedBarColor
End Property


Public Property Let SelectedBarColor(lngVal As OLE_COLOR)
    If lngVal <> uSelectedBarColor Then
        uSelectedBarColor = lngVal
        PropertyChanged "SelectedBarColor"
    End If
End Property

Public Property Let BarColor(lngVal As OLE_COLOR)
    If lngVal <> uBarColor Then
        uBarColor = lngVal
        DrawChart
        PropertyChanged "BarColor"
    End If
End Property


Public Property Let ColorBars(blnVal As Boolean)
    If blnVal <> uColorBars Then
        uColorBars = blnVal
        DrawChart
        PropertyChanged "ColorBars"
    End If
End Property
Public Property Get ColorBars() As Boolean
    ColorBars = uColorBars
End Property

Private Sub Swap(ByRef var1 As Variant, ByRef var2 As Variant)
    
    Dim varDummy As Variant
    
    varDummy = var1
    var1 = var2
    var2 = varDummy

End Sub

Private Sub cmdCmd_Click(Index As Integer)

    Select Case Index
        Case IDX_SAVE
            mnuSaveAs_Click
        
        Case IDX_PRINT
            mnuPrint_Click
    
        Case IDX_COPY
            mnuEditCopy_Click
    
        Case IDX_INFO
            mnuSelectionInfo_Click
        
        Case IDX_LEGEND
            mnuViewLegend_Click
        
        Case IDX_PROPERTIES
            mnuProperties_Click
        
    End Select

End Sub

Private Sub lblDescription_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim lScrollvalue As Integer
    
    If Button = vbLeftButton Then
        If uSelectable Then
            uSelectedColumn = Index
            uOldSelection = uSelectedColumn
            lScrollvalue = vsbContainer.Value
            bLegendClicked = True
            DrawChart
            'display information
            Call DisplayInfo(Index)
            bLegendClicked = False
            vsbContainer.Value = lScrollvalue
        End If
    End If
End Sub

Private Sub lblInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        offsetX = X
        offsetY = Y
        lblInfo.Drag
    Else
        PopupMenu mnuMain
    End If
End Sub


Private Sub lblSlider_Click()
    mnuViewLegend.Checked = Not mnuViewLegend.Checked
    bDisplayLegend = mnuViewLegend.Checked
    ShowLegend Not (bDisplayLegend)
    DrawChart
End Sub

Private Sub mnuEditCopy_Click()
    Clipboard.SetData UserControl.Image
End Sub

Private Sub mnuLegendHide_Click()
    mnuViewLegend.Checked = Not mnuViewLegend.Checked
    bDisplayLegend = mnuViewLegend.Checked
    ShowLegend True
    DrawChart
End Sub



Private Sub mnuPrint_Click()
    On Error Resume Next
    
    Screen.MousePointer = vbHourglass
    Printer.PaintPicture UserControl.Image, 0, 0, UserControl.Width, UserControl.Height
    Printer.EndDoc
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuProperties_Click()
    'frmProperties.Show vbModal
End Sub

Private Sub mnuSaveAs_Click()
   
    Dim sFilters As String
    Dim OFN As OPENFILENAME
    Dim lRet As Long
    
    'used after call
    Dim buff As String
    Dim sLname As String
    Dim sSname As String
    Dim strBuffer As String
    Dim blnReturn As Boolean
    
    'create string of filters for the dialog
    sFilters = "Windows Bitmap" & vbNullChar & "*.bmp" & vbNullChar & vbNullChar
    If uSaveAsCaption = Empty Then
        uSaveAsCaption = "Save graph"
    End If
    
    With OFN
        .nStructSize = Len(OFN)
        .hWndOwner = UserControl.hWnd
        .sFilter = sFilters
        .nFilterIndex = 0
        .sFile = "XChart.bmp" & Space$(1024) & vbNullChar & vbNullChar
        .nMaxFile = Len(.sFile)
        .sDefFileExt = "bmp" & vbNullChar & vbNullChar
        .sFileTitle = vbNullChar & Space$(512) & vbNullChar & vbNullChar
        .nMaxTitle = Len(OFN.sFileTitle)
        .sInitialDir = strBuffer & vbNullChar & vbNullChar
        .sDialogTitle = uSaveAsCaption
        .flags = OFS_FILE_SAVE_FLAGS
    End With
   
    'call the API
    blnReturn = GetSaveFileName(OFN)
    
    If blnReturn Then
        SavePicture UserControl.Image, OFN.sFile
    End If

End Sub

Private Sub mnuSelectionInfo_Click()
    
    mnuSelectionInfo.Checked = Not mnuSelectionInfo.Checked
    uDisplayDescript = mnuSelectionInfo.Checked
    Call DisplayInfo(uSelectedColumn)
    
End Sub

Private Sub mnuViewLegend_Click()
    mnuViewLegend.Checked = Not mnuViewLegend.Checked
    bDisplayLegend = mnuViewLegend.Checked
    ShowLegend Not (bDisplayLegend)
    DrawChart
End Sub


Private Sub picContainer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuLegend
    End If
End Sub

Private Sub picLegend_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuLegend
    End If
End Sub

Private Sub UserControl_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Left = X - offsetX
    Source.Top = Y - offsetY
End Sub

Private Sub UserControl_Initialize()
    Set cItems = New Collection
End Sub

Private Sub UserControl_InitProperties()
    
    Dim X As Integer
    Dim oChartItem As ChartItem
    
    uTopMargin = 50 * Screen.TwipsPerPixelY
    uBottomMargin = 55 * Screen.TwipsPerPixelY
    uLeftMargin = 55 * Screen.TwipsPerPixelX
    uRightMargin = 55 * Screen.TwipsPerPixelX
    uContentBorder = True
    uSelectable = False
    uHotTracking = False
    uSelectedColumn = -1
    uOldSelection = -1
    uChartTitle = UserControl.Name
    uChartSubTitle = ""
    uAxisYOn = True
    uAxisXOn = True
    uColorBars = False
    uIntersectMajor = 10
    uIntersectMinor = 2
    uMaxYValue = 100
    UserControl.BackColor = vbWindowBackground
    UserControl.ForeColor = vbWindowText
    '----------------------------------------------------
    'added by M. Costa on 21/06/2002
    uMinYValue = 0
    uBarColor = vbGreen
    uSelectedBarColor = vbYellow
    uMajorGridColor = vbWhite
    uMinorGridColor = vbBlack
    uLegendBackColor = UserControl.BackColor
    uLegendForeColor = UserControl.ForeColor
    uInfoBackColor = vbInfoBackground
    uInfoForeColor = vbInfoText
    uXAxisLabelColor = UserControl.ForeColor
    uYAxisLabelColor = UserControl.ForeColor
    uXAxisItemsColor = UserControl.ForeColor
    uYAxisItemsColor = UserControl.ForeColor
    uChartTitleColor = UserControl.ForeColor
    uChartSubTitleColor = UserControl.ForeColor
    uBarSymbolColor = uBarColor
    uLineColor = uBarColor
    uMenuType = xcPopUpMenu
    uChartType = xcBar
    uBarSymbol = "*"
    uBarWidthPercentage = 100
    uMenuItems = Empty
    uInfoItems = Empty
    uSaveAsCaption = Empty
    uAutoRedraw = True
    Set uBarPicture = Nothing
    uBarPictureTile = False
    Set uPicture = Nothing
    uPictureTile = False
    uMinorGridOn = True
    uMajorGridOn = True
    uLineWidth = 1
    uBarFillStyle = vbFSSolid
    uBarFillStyle = vbCross
    uLineStyle = vbSolid
    uBarShadow = True
    uBarShadowColor = vbBlack
    uMeanOn = False
    uMeanCaption = Empty
    uDataFormat = Empty
    '----------------------------------------------------
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim oItem As ChartItem
    Dim intSelectedCol As Integer
    
    If Button = vbLeftButton Then
        
        On Error GoTo TrackExit
        
        intSelectedCol = InColumn(X, Y)
        If intSelectedCol >= 0 Then
            If Not bProcessingOver Then
                bProcessingOver = True
                uSelectedColumn = intSelectedCol
                If Not uSelectedColumn = uOldSelection Then
                    DrawChart
                    uOldSelection = uSelectedColumn
                    If (uMeanOn = True) And (uSelectedColumn = cItems.Count - 1) Then
                        'do nothing in case of mean bar selected
                    Else
                        oItem = cItems(uSelectedColumn + 1)
                        RaiseEvent ItemClick(oItem)
                    End If
                End If
                bProcessingOver = False
             End If
        End If
    ElseIf Button = vbRightButton Then
        If uMenuType = xcPopUpMenu Then
            FixMenu
            mnuSelectionInfo.Visible = (uSelectable = True)
            PopupMenu mnuMain
        End If
    End If
    RaiseEvent MouseDown(Button, Shift, X, Y)
    
TrackExit:
    Exit Sub

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If (uHotTracking = True) Or (Button = vbLeftButton) Then
        'either in case of hot tracking or not, simulate the mouse left button down
        Call UserControl_MouseDown(vbLeftButton, Shift, X, Y)
    End If

End Sub

Public Sub Refresh()
    DrawChart
End Sub

Public Sub Clear()
    Set cItems = Nothing
    Set cItems = New Collection
    ClearLegendItems
    'the following forces the drawing chart routine to not enhance the description
    'in the legend (if it is visible); the legend items were already deleted!
    uSelectedColumn = -1
    DrawChart
End Sub

Public Sub DrawChart()
    
    Dim x1              As Single
    Dim x2              As Single
    Dim y1              As Single
    Dim y2              As Single
    Dim xTemp           As Single
    Dim yTemp           As Single
    Dim xPrev           As Single
    Dim yPrev           As Single
    Dim sngRowHeight    As Single
    Dim CurrentColor    As Integer
    Dim iCols           As Integer
    Dim X               As Integer
    Dim oChartItem      As ChartItem
    Dim sngColWidth     As Single
    
    'do not redraw the chart if not required
    If uAutoRedraw = False Then Exit Sub

    'calculate the data distribution in the y-axis
    FixData
    
    lblInfo.ForeColor = uInfoForeColor
    lblInfo.BackColor = uInfoBackColor
    lblDescription(0).ForeColor = uLegendForeColor
    
    iCols = cItems.Count
    
    mnuSelectionInfo.Checked = uDisplayDescript
    lblInfo.Visible = False
    If uDisplayDescript And uSelectedColumn > -1 Then lblInfo.Visible = True
    
    'hide existing legend
    If bDisplayLegend Then
        vsbContainer.Visible = False
        picContainer.Visible = False
    End If
    
    If Not bResize Then ClearLegendItems

    uRowHeight = ((UserControl.ScaleHeight - (uTopMargin + uBottomMargin)) / uRangeY)
    If iCols Then
        uColWidth = ((UserControl.ScaleWidth - (uLeftMargin + uRightMargin)) / iCols)
    End If
    
    UserControl.Cls
    If uPicture Is Nothing Then
    Else
        'paint the background image
        Call DrawPicture(uLeftMargin, UserControl.ScaleWidth - uRightMargin, _
                         uTopMargin, UserControl.ScaleHeight - uBottomMargin, _
                         uPictureTile, uPicture)
    End If

    If iCols Then ReDim uColumns(iCols - 1, 1)

    On Error Resume Next
    
    'dump chart title
    UserControl.ForeColor = uChartTitleColor
    If bDisplayLegend Then
        x1 = ((UserControl.ScaleWidth - picContainer.ScaleWidth) / 2)
    Else
        x1 = (UserControl.ScaleWidth / 2)
    End If
    UserControl.CurrentX = x1 - (UserControl.TextWidth(uChartTitle) / 2)
    UserControl.CurrentY = 0
    UserControl.FontBold = True
    UserControl.Print uChartTitle
    UserControl.FontBold = False
    
    'dump chart subtitle
    UserControl.ForeColor = uChartSubTitleColor
    UserControl.FontSize = UserControl.FontSize - 2
    If bDisplayLegend Then
        UserControl.CurrentX = ((UserControl.ScaleWidth - picContainer.ScaleWidth) / 2) - (UserControl.TextWidth(uChartSubTitle) / 2)
    Else
        UserControl.CurrentX = (UserControl.ScaleWidth / 2) - (UserControl.TextWidth(uChartSubTitle) / 2)
    End If
    UserControl.Print uChartSubTitle
    UserControl.FontSize = UserControl.FontSize + 2
    
    If uAxisYOn Then
        'draw Y axis
        UserControl.ForeColor = uYAxisItemsColor
        For X = uMinYValue To uMaxYValue
            x1 = uLeftMargin + (2 * Screen.TwipsPerPixelX)
            x2 = UserControl.ScaleWidth - uRightMargin
            If uDataType = DT_NEG Then
                y1 = (UserControl.ScaleHeight - uBottomMargin) + ((Abs(X) - Abs(uMinYValue)) * uRowHeight)
            Else
                y1 = (UserControl.ScaleHeight - uBottomMargin) - ((X - uMinYValue) * uRowHeight)
            End If
            If (X = uMinYValue) Or (X = uMaxYValue) Or ((X Mod uIntersectMajor) = 0) Then
                If uMajorGridOn Then
                    UserControl.Line (x1, y1)-(x2, y1), uMajorGridColor
                End If
                UserControl.FontSize = UserControl.FontSize - 2
                UserControl.CurrentX = uLeftMargin - UserControl.TextWidth(X) - (5 * Screen.TwipsPerPixelX)
                UserControl.CurrentY = y1 - (UserControl.TextHeight("0") / 2)
                UserControl.Print X
                UserControl.FontSize = UserControl.FontSize + 2
            ElseIf ((uMaxYValue - X) Mod uIntersectMinor = 0) Then
                If uMinorGridOn Then
                    UserControl.Line (x1, y1)-(x2, y1), uMinorGridColor
                End If
            End If
        Next X
    End If

    On Error GoTo 0
    If uContentBorder Then
        UserControl.Line (uLeftMargin, uTopMargin)-(UserControl.ScaleWidth - uRightMargin, UserControl.ScaleHeight - uBottomMargin), uMajorGridColor, B
    End If
    
    'draw bars, lines, symbols,...
    For X = 0 To cItems.Count - 1
        oChartItem = cItems(X + 1)
        x1 = (X * uColWidth) + uLeftMargin + (2 * Screen.TwipsPerPixelX)    'increment by 2 pixs.
        x2 = x1 + uColWidth - (2 * Screen.TwipsPerPixelX)                   'decrement by 2 pixs.
        If uDataType = DT_POS Then
            sngRowHeight = uRowHeight * (oChartItem.Value - uMinYValue)
            y2 = UserControl.ScaleHeight - uBottomMargin
            y1 = y2 - sngRowHeight
        ElseIf uDataType = DT_NEG Then
            sngRowHeight = uRowHeight * (Abs(CDbl(oChartItem.Value)) - Abs(uMaxYValue))
            y1 = uTopMargin
            y2 = y1 + sngRowHeight
        Else
            sngRowHeight = (-CDbl(oChartItem.Value) * uRowHeight)
            y1 = UserControl.ScaleHeight - uBottomMargin
            y1 = y1 - uRowHeight * Abs(uMinYValue)
            y2 = y1 + sngRowHeight
        End If
        sngRowHeight = Abs(sngRowHeight)
        'be sure the y1 coordinate is always less than y2
        If y2 < y1 Then Call Swap(y1, y2)

        'save coordinates of bar (only Y since X is calculated)
        uColumns(X, 0) = y1
        uColumns(X, 1) = y2

        If ((uChartType And XC_BAR) = XC_BAR) _
        Or (uChartType And XC_OVAL) = XC_OVAL _
        Or (uChartType And XC_RHOMBUS) = XC_RHOMBUS _
        Or (uChartType And XC_TRAPEZIUM) = XC_TRAPEZIUM _
        Or (uChartType And XC_TRIANGLE) = XC_TRIANGLE Then
            'draw the bars in the right shape

            'adjust x-coordinates depending on bar width percentage
            sngColWidth = uColWidth * uBarWidthPercentage / 100
            xTemp = x1 + ((uColWidth - sngColWidth) / 2)
            x2 = x2 - ((uColWidth - sngColWidth) / 2)
            'Selected bar outline
            UserControl.DrawWidth = uLineWidth
            UserControl.FillStyle = uBarFillStyle
            If X = uSelectedColumn And uSelectable Then
                UserControl.FillColor = uSelectedBarColor
                If (uChartType And XC_OVAL) = XC_OVAL Then
                    Call DrawOval(xTemp, x2, y1, y2, sngColWidth, sngRowHeight, uBarColor)
                ElseIf (uChartType And XC_BAR) = XC_BAR Then
                    Call DrawRectangle(oChartItem.Value, xTemp, x2, y1, y2, uBarColor, (uMeanOn = True) And (X = cItems.Count - 1))
                ElseIf (uChartType And XC_TRIANGLE) = XC_TRIANGLE Then
                    Call DrawTriangle(oChartItem.Value, xTemp, x2, y1, y2)
                ElseIf (uChartType And XC_TRAPEZIUM) = XC_TRAPEZIUM Then
                    Call DrawTrapezium(oChartItem.Value, xTemp, x2, y1, y2)
                ElseIf (uChartType And XC_RHOMBUS) = XC_RHOMBUS Then
                    Call DrawRhombus(oChartItem.Value, xTemp, x2, y1, y2)
                End If
                UserControl.DrawWidth = 1
                UserControl.FillStyle = vbFSTransparent

                'display information
                Call DisplayInfo(X)
            Else
                If (uMeanOn = True) And (X = cItems.Count - 1) Then
                    UserControl.FillColor = uMeanColor
                Else
                    UserControl.FillColor = IIf(uColorBars, QBColor(CurrentColor), uBarColor)
                End If
                UserControl.FillStyle = uBarFillStyle
                UserControl.DrawWidth = uLineWidth
                If (uChartType And XC_OVAL) = XC_OVAL Then
                    Call DrawOval(xTemp, x2, y1, y2, sngColWidth, sngRowHeight, uSelectedBarColor)
                ElseIf (uChartType And XC_BAR) = XC_BAR Then
                    Call DrawRectangle(oChartItem.Value, xTemp, x2, y1, y2, uSelectedBarColor, (uMeanOn = True) And (X = cItems.Count - 1))
                ElseIf (uChartType And XC_TRIANGLE) = XC_TRIANGLE Then
                    Call DrawTriangle(oChartItem.Value, xTemp, x2, y1, y2)
                ElseIf (uChartType And XC_TRAPEZIUM) = XC_TRAPEZIUM Then
                    Call DrawTrapezium(oChartItem.Value, xTemp, x2, y1, y2)
                ElseIf (uChartType And XC_RHOMBUS) = XC_RHOMBUS Then
                    Call DrawRhombus(oChartItem.Value, xTemp, x2, y1, y2)
                End If
                UserControl.DrawWidth = 1
                UserControl.FillStyle = vbFSTransparent
            End If
        End If
        If (uChartType And XC_SYMBOL) = XC_SYMBOL Then
            'draw the symbol in the higher (absolute) point
            If uDataType = DT_NEG Then
                yTemp = y2
            ElseIf uDataType = DT_POS Then
                yTemp = y1
            Else
                yTemp = IIf((oChartItem.Value > 0), y1, y2)
            End If
            xTemp = x1 + (uColWidth / 2) - (UserControl.TextWidth(uBarSymbol) / 2)
            yTemp = yTemp - (UserControl.TextHeight(uBarSymbol) / 2)
            If (uMeanOn = True) And (X = cItems.Count - 1) Then
                UserControl.ForeColor = uMeanColor
            Else
                UserControl.ForeColor = uBarSymbolColor
            End If
            UserControl.CurrentX = xTemp
            UserControl.CurrentY = yTemp
            UserControl.FontSize = UserControl.FontSize + 2
            UserControl.Print uBarSymbol
            UserControl.FontSize = UserControl.FontSize - 2
        End If
        If (uChartType And XC_LINE) = XC_LINE Then
            'draw the lines
            If uDataType = DT_NEG Then
                yTemp = y2
            ElseIf uDataType = DT_POS Then
                yTemp = y1
            Else
                yTemp = IIf((oChartItem.Value > 0), y1, y2)
            End If
            xTemp = x1 + (uColWidth / 2)
            'check if it's the first data: if it is, do not draw the line
            If (X > 0) And (uMeanOn = True And X < cItems.Count - 1) Then
                UserControl.DrawStyle = uLineStyle
                UserControl.DrawWidth = uLineWidth
                UserControl.Line (xPrev, yPrev)-(xTemp, yTemp), uLineColor
                UserControl.DrawWidth = 1
                UserControl.DrawStyle = vbSolid
            End If
            xPrev = xTemp
            yPrev = yTemp
        End If

        'display X-axis labels and ticks
        If uAxisXOn Then
            UserControl.ForeColor = uXAxisItemsColor
            UserControl.FontSize = UserControl.FontSize - 1
            
            xTemp = (((x2 - x1) / 2) + x1) / Screen.TwipsPerPixelX
            yTemp = (UserControl.ScaleHeight - uBottomMargin + UserControl.TextWidth(oChartItem.XAxisDescription) / 1.25) / Screen.TwipsPerPixelY
            
            PrintRotText UserControl.hDC, oChartItem.XAxisDescription, xTemp, yTemp, 270
            
            yTemp = (UserControl.ScaleHeight - uBottomMargin) + Screen.TwipsPerPixelX
            UserControl.Line (xTemp * Screen.TwipsPerPixelX, yTemp)-(xTemp * Screen.TwipsPerPixelX, yTemp + 2 * Screen.TwipsPerPixelX), uMajorGridColor
            UserControl.FontSize = UserControl.FontSize + 1
        End If
        'Add Legend item
        If Not bResize Then
            If (uMeanOn = True) And (X = cItems.Count - 1) Then
                UserControl.FillColor = uMeanColor
            ElseIf ((uChartType And XC_BAR) = XC_BAR) _
            Or (uChartType And XC_OVAL) = XC_OVAL _
            Or (uChartType And XC_RHOMBUS) = XC_RHOMBUS _
            Or (uChartType And XC_TRAPEZIUM) = XC_TRAPEZIUM _
            Or (uChartType And XC_TRIANGLE) = XC_TRIANGLE Then
                'do nothing, since FillColor is already set
            ElseIf (uChartType And XC_LINE) = XC_LINE Then
                UserControl.FillColor = uLineColor
            ElseIf (uChartType And XC_SYMBOL) = XC_SYMBOL Then
                UserControl.FillColor = uBarSymbolColor
            End If
            AddLegendItem oChartItem.SelectedDescription, _
                          UserControl.FillColor, uLegendForeColor
        End If
        
        If uColorBars = True Then
            CurrentColor = CurrentColor + 1
            If CurrentColor >= 15 Then CurrentColor = 0
        End If
    Next X

    'Print the x axis label
    If Len(uXAxisLabel) Then
        UserControl.FontSize = UserControl.FontSize - 1
        UserControl.CurrentY = UserControl.ScaleHeight - UserControl.TextHeight(uXAxisLabel) * 1.5
        If bDisplayLegend Then
            UserControl.CurrentX = ((UserControl.ScaleWidth - picContainer.ScaleWidth) / 2) - (UserControl.TextWidth(uXAxisLabel) / 2)
        Else
            UserControl.CurrentX = (UserControl.ScaleWidth / 2) - (UserControl.TextWidth(uXAxisLabel) / 2)
        End If
        UserControl.ForeColor = uXAxisLabelColor
        UserControl.Print uXAxisLabel
        UserControl.FontSize = UserControl.FontSize + 1
    End If
    
    'print the y axis label
    If Len(uYAxisLabel) > 0 Then
        UserControl.FontSize = UserControl.FontSize - 1
        UserControl.ForeColor = uYAxisLabelColor
        PrintRotText UserControl.hDC, uYAxisLabel, UserControl.TextHeight(uYAxisLabel) / Screen.TwipsPerPixelX, UserControl.ScaleHeight / 2 / Screen.TwipsPerPixelY, 90
        UserControl.FontSize = UserControl.FontSize + 1
    End If

    If bDisplayLegend = True Then
        If uSelectable And uSelectedColumn > -1 Then
            
            Dim perScreen As Integer
            Dim scrollValue As Integer
                        
            perScreen = Abs((picLegend.ScaleHeight / ((Box(0).Height + (10 * Screen.TwipsPerPixelY)))) - 1)
                        
            If (uSelectedColumn + 1) > perScreen Then
                scrollValue = ((uSelectedColumn + 1) * ((Box(0).Height / Screen.TwipsPerPixelY) + 10)) - (Box(perScreen).Top / Screen.TwipsPerPixelY)
                If scrollValue > vsbContainer.Max Then scrollValue = vsbContainer.Max
                vsbContainer.Value = scrollValue
            Else
                vsbContainer.Value = 0
            End If
                        
            picContainer.Cls
            picContainer.Line ((Box(uSelectedColumn).Left - 3 * Screen.TwipsPerPixelX), (Box(uSelectedColumn).Top - 3 * Screen.TwipsPerPixelY))-(lblDescription(uSelectedColumn).Left + lblDescription(uSelectedColumn).Width + 2 * Screen.TwipsPerPixelX, Box(uSelectedColumn).Top + Box(uSelectedColumn).Height + 2 * Screen.TwipsPerPixelY), uSelectedBarColor, B
        End If
        picContainer.Visible = True
    End If
    
End Sub

Public Function ShowLegend(Optional bHidden As Boolean = False)
    
    lblSlider.Height = picLegend.ScaleHeight
    picLegend.Line (0, 0)-(picLegend.ScaleWidth - Screen.TwipsPerPixelX, picLegend.ScaleHeight - Screen.TwipsPerPixelY), &HFFE0E0, B
    
    If bHidden Then bDisplayLegend = False Else bDisplayLegend = True
    
    If bDisplayLegend Then
        picLegend.BackColor = uLegendBackColor
        picContainer.BackColor = uLegendBackColor
        uRightMargin = uRightMargin + picLegend.ScaleWidth
        picLegend.Move UserControl.ScaleWidth - picLegend.Width + Screen.TwipsPerPixelX, 0, picLegend.Width, UserControl.ScaleHeight
        lblSlider = Chr(187)
    Else
        uRightMargin = uRightMargin - picLegend.Width
        picLegend.Move UserControl.ScaleWidth - lblSlider.Width
        lblSlider = Chr(171)
    End If

End Function

Private Sub AddLegendItem(sDescription As String, lngBackColor As OLE_COLOR, lngForeColor As OLE_COLOR)
    Dim X As Integer
    Dim ShortDescript As String
    
    ShortDescript = sDescription
    If Len(ShortDescript) > 17 Then ShortDescript = Left(ShortDescript, 15) & ".."
    
    If bLegendAdded Then
        X = Box.Count
        Load Box(X)
        Load lblDescription(X)
        
        Box(X).BackColor = lngBackColor
        Box(X).Top = Box(X - 1).Top + Box(X - 1).Height + 10 * Screen.TwipsPerPixelY
        lblDescription(X).Top = Box(X).Top
    Else
        X = 0
        Box(X).BackColor = lngBackColor
        bLegendAdded = True
    End If
    lblDescription(X).ForeColor = lngForeColor
    lblDescription(X) = ShortDescript
    lblDescription(X).ToolTipText = sDescription
    
    Box(X).Visible = True
    lblDescription(X).Visible = True
            
    picContainer.Height = ((Box(0).Height + (10 * Screen.TwipsPerPixelY)) * Box.Count - 1) + 10 * Screen.TwipsPerPixelY
    If picContainer.ScaleHeight > picLegend.ScaleHeight Then
        vsbContainer.Max = (picContainer.ScaleHeight / Screen.TwipsPerPixelY) - (picLegend.ScaleHeight / Screen.TwipsPerPixelY)
        If Not vsbContainer.Visible Then vsbContainer.Visible = True
    Else
        vsbContainer.Visible = False
    End If
End Sub

Private Sub ClearLegendItems()
    Dim X As Integer
    
    On Error Resume Next    'we are expecting an error for item 1
    
    If bLegendAdded Then
        bLegendAdded = False
        
        For X = 1 To Box.Count
            Unload Box(X)
            Unload lblDescription(X)
            If Err.Number Then Err.Clear
            vsbContainer.Value = 0
            Box(0).Visible = False
            lblDescription(0).Visible = False
        Next X
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    On Error Resume Next
    
    With PropBag
        uTopMargin = .ReadProperty("uTopMargin")
        uBottomMargin = .ReadProperty("uBottomMargin")
        uLeftMargin = .ReadProperty("uLeftMargin")
        uRightMargin = .ReadProperty("uRightMargin")
        uContentBorder = .ReadProperty("uContentBorder")
        uSelectable = .ReadProperty("uSelectable", False)
        uHotTracking = .ReadProperty("uHotTracking", False)
        uSelectedColumn = .ReadProperty("uSelectedColumn", -1)
        uChartTitle = .ReadProperty("uChartTitle", UserControl.Name)
        uChartSubTitle = .ReadProperty("uChartSubTitle", uChartSubTitle)
        uAxisYOn = .ReadProperty("uAxisXOn", uAxisXOn)
        uAxisXOn = .ReadProperty("uAxisYOn", uAxisYOn)
        uColorBars = .ReadProperty("uColorBars", False)
        uIntersectMajor = .ReadProperty("uIntersectMajor", 10)
        uIntersectMinor = .ReadProperty("uIntersectMinor", 2)
        uMaxYValue = .ReadProperty("uMaxYValue", 100)
        uDisplayDescript = .ReadProperty("uDisplayDescript", False)
        uXAxisLabel = .ReadProperty("uXAxisLabel")
        uYAxisLabel = .ReadProperty("uYAxisLabel")
        UserControl.BackColor = .ReadProperty("BackColor")
        UserControl.ForeColor = .ReadProperty("ForeColor")
        '----------------------------------------------------
        'added by M. Costa on 21/06/2002
        uMinYValue = .ReadProperty("MinY")
        uBarColor = .ReadProperty("BarColor")
        uSelectedBarColor = .ReadProperty("SelectedBarColor")
        uMajorGridColor = .ReadProperty("MajorGridColor")
        uMinorGridColor = .ReadProperty("MinorGridColor")
        uLegendBackColor = .ReadProperty("LegendBackColor")
        uLegendForeColor = .ReadProperty("LegendForeColor")
        uInfoBackColor = .ReadProperty("InfoBackColor")
        uInfoForeColor = .ReadProperty("InfoForeColor")
        uXAxisLabelColor = .ReadProperty("XAxisLabelColor")
        uYAxisLabelColor = .ReadProperty("YAxisLabelColor")
        uXAxisItemsColor = .ReadProperty("XAxisItemsColor")
        uYAxisItemsColor = .ReadProperty("YAxisItemsColor")
        uChartTitleColor = .ReadProperty("ChartTitleColor")
        uChartSubTitleColor = .ReadProperty("ChartSubTitleColor")
        uChartType = .ReadProperty("ChartType")
        uMenuType = .ReadProperty("MenuType")
        uMenuItems = .ReadProperty("MenuItems")
        uInfoItems = .ReadProperty("InfoItems")
        uSaveAsCaption = .ReadProperty("SaveAsCaption")
        uAutoRedraw = .ReadProperty("AutoRedraw")
        uBarWidthPercentage = .ReadProperty("BarWidthPercentage")
        uBarSymbol = .ReadProperty("BarSymbol")
        Set uBarPicture = .ReadProperty("BarPicture", Nothing)
        uBarPictureTile = .ReadProperty("BarPictureTile")
        Set uPicture = .ReadProperty("Picture", Nothing)
        uPictureTile = .ReadProperty("PictureTile")
        uMinorGridOn = .ReadProperty("MinorGridOn")
        uMajorGridOn = .ReadProperty("MajorGridOn")
        uLineWidth = .ReadProperty("LineWidth")
        uLineColor = .ReadProperty("LineColor")
        uBarSymbolColor = .ReadProperty("BarSymbolColor")
        uBarFillStyle = .ReadProperty("BarFillStyle")
        uLineStyle = .ReadProperty("LineStyle")
        uBarShadow = .ReadProperty("BarShadow")
        uBarShadowColor = .ReadProperty("BarShadowColor")
        uMeanOn = .ReadProperty("MeanOn")
        uMeanColor = .ReadProperty("MeanColor")
        uMeanCaption = .ReadProperty("MeanCaption")
        uDataFormat = .ReadProperty("DataFormat")
        '----------------------------------------------------
        uOldSelection = -1
    End With

End Sub

Private Sub UserControl_Resize()
    If bDisplayLegend Then
        picLegend.Left = UserControl.ScaleWidth - picLegend.Width
    Else
        picLegend.Left = UserControl.ScaleWidth - lblSlider.Width
    End If
    picLegend.Height = UserControl.ScaleHeight
    vsbContainer.Height = picLegend.ScaleHeight
    lblSlider.Height = picLegend.ScaleHeight

    bResize = True
    DrawChart
    bResize = False

End Sub

Private Sub UserControl_Show()
    DrawChart
    FixMenu
End Sub

Private Sub UserControl_Terminate()
    Set cItems = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    With PropBag
        .WriteProperty "uTopMargin", uTopMargin
        .WriteProperty "uBottomMargin", uBottomMargin
        .WriteProperty "uLeftMargin", uLeftMargin
        .WriteProperty "uRightMargin", uRightMargin
        .WriteProperty "uContentBorder", uContentBorder
        .WriteProperty "uSelectable", uSelectable
        .WriteProperty "uHotTracking", uHotTracking
        .WriteProperty "uSelectedColumn", uSelectedColumn
        .WriteProperty "uChartTitle", uChartTitle
        .WriteProperty "uChartSubTitle", uChartSubTitle
        .WriteProperty "uAxisXOn", uAxisXOn
        .WriteProperty "uAxisYOn", uAxisYOn
        .WriteProperty "uColorBars", uColorBars
        .WriteProperty "uIntersectMajor", uIntersectMajor
        .WriteProperty "uIntersectMinor", uIntersectMinor
        .WriteProperty "uMaxYValue", uMaxYValue
        .WriteProperty "uDisplayDescript", uDisplayDescript
        .WriteProperty "uXAxisLabel", uXAxisLabel
        .WriteProperty "uYAxislabel", uYAxisLabel
        .WriteProperty "BackColor", UserControl.BackColor
        .WriteProperty "ForeColor", UserControl.ForeColor
        '----------------------------------------------------
        'added by M. Costa on 21/06/2002
        .WriteProperty "MinY", uMinYValue
        .WriteProperty "BarColor", uBarColor
        .WriteProperty "SelectedBarColor", uSelectedBarColor
        .WriteProperty "MajorGridColor", uMajorGridColor
        .WriteProperty "MinorGridColor", uMinorGridColor
        .WriteProperty "LegendBackColor", uLegendBackColor
        .WriteProperty "LegendForeColor", uLegendForeColor
        .WriteProperty "InfoBackColor", uInfoBackColor
        .WriteProperty "InfoForeColor", uInfoForeColor
        .WriteProperty "XAxisLabelColor", uXAxisLabelColor
        .WriteProperty "YAxisLabelColor", uYAxisLabelColor
        .WriteProperty "XAxisItemsColor", uXAxisItemsColor
        .WriteProperty "YAxisItemsColor", uYAxisItemsColor
        .WriteProperty "ChartTitleColor", uChartTitleColor
        .WriteProperty "ChartSubTitleColor", uChartSubTitleColor
        .WriteProperty "ChartType", uChartType
        .WriteProperty "MenuType", uMenuType
        .WriteProperty "MenuItems", uMenuItems
        .WriteProperty "InfoItems", uInfoItems
        .WriteProperty "SaveAsCaption", uSaveAsCaption
        .WriteProperty "AutoRedraw", uAutoRedraw
        .WriteProperty "BarWidthPercentage", uBarWidthPercentage
        .WriteProperty "BarSymbol", uBarSymbol
        .WriteProperty "BarPicture", uBarPicture, Nothing
        .WriteProperty "BarPictureTile", uBarPictureTile
        .WriteProperty "Picture", uPicture, Nothing
        .WriteProperty "PictureTile", uPictureTile
        .WriteProperty "MinorGridOn", uMinorGridOn
        .WriteProperty "MajorGridOn", uMajorGridOn
        .WriteProperty "LineWidth", uLineWidth
        .WriteProperty "LineColor", uLineColor
        .WriteProperty "BarSymbolColor", uBarSymbolColor
        .WriteProperty "BarFillStyle", uBarFillStyle
        .WriteProperty "LineStyle", uLineStyle
        .WriteProperty "BarShadow", uBarShadow
        .WriteProperty "BarShadowColor", uBarShadowColor
        .WriteProperty "MeanOn", uMeanOn
        .WriteProperty "MeanColor", uMeanColor
        .WriteProperty "MeanCaption", uMeanCaption
        .WriteProperty "DataFormat", uDataFormat
        '----------------------------------------------------
    End With

End Sub

Private Sub vsbContainer_Change()
    picContainer.Top = -vsbContainer.Value * Screen.TwipsPerPixelY
End Sub

Private Sub vsbContainer_Scroll()
    picContainer.Top = -vsbContainer.Value * Screen.TwipsPerPixelY
End Sub

Public Property Get LineColor() As OLE_COLOR
    LineColor = uLineColor
End Property

Public Property Let LineColor(lngVal As OLE_COLOR)
    If lngVal <> uLineColor Then
        uLineColor = lngVal
        DrawChart
        PropertyChanged "LineColor"
    End If
End Property

Public Property Get BarSymbolColor() As OLE_COLOR
    BarSymbolColor = uBarSymbolColor
End Property

Public Property Let BarSymbolColor(lngVal As OLE_COLOR)
    If uBarSymbolColor <> lngVal Then
        uBarSymbolColor = lngVal
        DrawChart
        PropertyChanged "BarSymbolColor"
    End If
End Property

Public Property Get BarFillStyle() As FillStyleConstants
    BarFillStyle = uBarFillStyle
End Property

Public Property Let BarFillStyle(intVal As FillStyleConstants)
    If uBarFillStyle <> intVal Then
        uBarFillStyle = intVal
        DrawChart
        PropertyChanged "BarFillStyle"
    End If
End Property
