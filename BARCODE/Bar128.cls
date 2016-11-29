VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bar128"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim Code_A As String
Dim Code_B() As Variant
Private Const CodeC = 99
Private Const CodeB = 100
Private Const CodeA = 101
Private Const FNC1 = 102
Private Const StartA = 103
Private Const StartB = 104
Private Const StartC = 105

Dim BarH As Long
Dim zBarText As String
Dim xObj As Object

Dim xPos As Long, xtop As Long, mCnt As Integer
Dim xStart As Integer, posCtr As Integer, xTotal As Long, chkSum As Long
'Dim me As PictureBox
Function Code128(zObj As Object, zBarH As Integer, BarText As String)
   Set xObj = zObj
   Init_Table
   zBarText = BarText
   xObj.Cls
   BarH = zBarH * 10
   xtop = 10
   
   xObj.BackColor = vbWhite
   xObj.AutoRedraw = True
   xObj.ScaleMode = 3
   
   xObj.Height = (xObj.TextHeight(zBarText) + BarH + 25) * Screen.TwipsPerPixelY
   xObj.Width = ((Test_String(zBarText) * 11)) * 20  'Screen.TwipsPerPixelX
   
   Paint_Code zBarText
   'xPic.Picture = me.Picture
End Function
Private Function Test_String(xstr As String)
    Dim ii As Long, jj As Integer, ctr As Integer
    ctr = 0
    jj = 0
    For ii = 1 To Len(xstr)
        If InStr("0123456789", Mid(xstr, ii, 1)) > 0 Then
           ctr = ctr + 1
        Else
           jj = jj + IIf(ctr = 0, 1, ctr)
           ctr = 0
        End If
    Next
    If (ctr >= 4 And ii >= Len(xstr)) Then
        If ctr Mod 2 <> 0 Then
           ctr = ctr - 1
        End If
        jj = jj + (ctr / 2)
    End If
    jj = jj + 3
    Test_String = jj
End Function

Private Sub Paint_Code(xstr As String)
    Dim ii As Long, jj As Integer, ctr As Integer
    xTotal = 0
    xPos = 1
    xStart = 0
    ctr = 0
    posCtr = 0
    mCnt = 0
    For ii = 1 To Len(xstr)
        If InStr("0123456789", Mid(xstr, ii, 1)) > 0 Then
           ctr = ctr + 1
        Else
           For jj = ii - ctr To ii
                PrintB Mid(xstr, jj, 1)
                mCnt = mCnt + 1
           Next
           ctr = 0
        End If
    Next
    If (ctr >= 4 And ii >= Len(xstr)) Then
        If ctr Mod 2 <> 0 Then
           mCnt = mCnt + 1
           PrintB Mid(xstr, ii - ctr, 1)
           ctr = ctr - 1
        End If
        PrintC Mid(xstr, ii - ctr, ctr)
    End If
    chkSum = xTotal Mod 103
    Draw_Bar CStr(Code_B(chkSum))
    Draw_Bar "1100011101011"
    'MsgBox me.TextWidth(xstr)
    xObj.CurrentX = ((xPos + 20) / 2) - xObj.TextWidth(xstr) / 2   ' Horizontal position.
    xObj.CurrentY = 15 + BarH    ' Vertical position.
    xObj.Print xstr   ' Print message.
    'Picture = Me.Image
End Sub
Private Sub PrintB(xstr As String)
        posCtr = posCtr + 1
        xTotal = xTotal + ((InStr(Code_A, xstr) - 1) * posCtr)
        If xStart <> StartA Then
           
           If xStart = 0 Then
              xTotal = xTotal + StartA
              xStart = StartA
              Draw_Bar CStr(Code_B(StartA))
            Else
              xStart = CodeA
              Draw_Bar CStr(Code_B(CodeA))
              posCtr = posCtr + 1
              xTotal = xTotal + (CodeA * posCtr)
            End If
        End If
        Draw_Bar CStr(Code_B(InStr(Code_A, xstr) - 1))
End Sub
Private Sub PrintC(xstr As String)
Dim jj As Integer
    If xStart <> StartC Then
       If xStart = 0 Then
          xTotal = xTotal + StartC
          xStart = StartC
          Draw_Bar CStr(Code_B(StartC))
        Else
          xStart = CodeC
          Draw_Bar CStr(Code_B(CodeC))
          posCtr = posCtr + 1
          xTotal = xTotal + (CodeC * posCtr)
        End If
    End If
    SetC xstr
    For jj = 1 To Len(xstr) Step 2
       posCtr = posCtr + 1
       xTotal = xTotal + CInt(Mid(xstr, jj, 2)) * posCtr
    Next

End Sub
Private Sub SetC(xstr As String)
   For ii = 1 To Len(xstr) Step 2
       Draw_Bar CStr(Code_B(CInt(Mid(xstr, ii, 2))))
       mCnt = mCnt + 1
   Next
End Sub
Private Sub Draw_Bar(Encoding As String)
    Dim ii As Integer
    For ii = 1 To Len(Encoding)
        xPos = xPos + 1
        xObj.Line (xPos + 10, xtop)-(xPos + 10, xtop + BarH), IIf(Mid(Encoding, ii, 1), vbBlack, vbWhite)
    Next
End Sub
Private Sub Init_Table()
    Code_A = " !""#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\}^_`abcdefghijklmnopqrstuvwxyz{|}~"
    Code_B = Array( _
             "11011001100", "11001101100", "11001100110", "10010011000", "10010001100", "10001001100", _
             "10011001000", "10011000100", "10001100100", "11001001000", "11001000100", "11000100100", _
             "10110011100", "10011011100", "10011001110", "10111001100", "10011101100", "10011100110", _
             "11001110010", "11001011100", "11001001110", "11011100100", "11001110100", "11101101110", _
             "11101001100", "11100101100", "11100100110", "11101100100", "11100110100", "11100110010", _
             "11011011000", "11011000110", "11000110110", "10100011000", "10001011000", "10001000110", _
             "10110001000", "10001101000", "10001100010", "11010001000", "11000101000", "11000100010", _
             "10110111000", "10110001110", "10001101110", "10111011000", "10111000110", "10001110110", _
             "11101110110", "11010001110", "11000101110", "11011101000", "11011100010", "11011101110", _
             "11101011000", "11101000110", "11100010110", "11101101000", "11101100010", "11100011010", _
             "11101111010", "11001000010", "11110001010", "10100110000", "10100001100", "10010110000", _
             "10010000110", "10000101100", "10000100110", "10110010000", "10110000100", "10011010000", _
             "10011000010", "10000110100", "10000110010", "11000010010", "11001010000", "11110111010", _
             "11000010100", "10001111010", "10100111100", "10010111100", "10010011110", "10111100100", _
             "10011110100", "10011110010", "11110100100", "11110010100", "11110010010", "11011011110", _
             "11011110110", "11110110110", "10101111000", "10100011110", "10001011110", "10111101000", _
             "10111100010", "11110101000", "11110100010", "10111011110", "10111101110", "11101011110", _
             "11110101110", "11010000100", "11010010000", "11010011100" _
             )
End Sub


