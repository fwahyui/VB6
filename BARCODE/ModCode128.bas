Attribute VB_Name = "ModCode128"
Dim CharSet As String
Dim arrEncoding() As Variant
Private Const CodeC = 99
Private Const CodeB = 100
Private Const CodeA = 101
Private Const FNC1 = 102
Private Const StartA = 103
Private Const StartB = 104
Private Const StartC = 105
Private Const StopChar = 106
Private Const EndChar = 107

Dim BarH As Long
Dim zBarText As String
Dim zHasCaption As Boolean

Dim xStart As Integer
Dim xObj As PictureBox

Dim myCols As Collection
Sub Bar128(zObj As Object, zBarH As Integer, BarText As String, Optional ByVal HasCaption As Boolean = False)
   Set xObj = zObj
   
   zBarText = BarText
   zHasCaption = HasCaption
   
   Init_Table 'Initialize Encoding
   Eval_String BarText
   
   
   xObj.Picture = Nothing
   BarH = zBarH * 72 'Inches to Pixel of barcode
   xObj.BackColor = vbWhite
   xObj.AutoRedraw = True
   xObj.ScaleMode = 3 'Pixel
   
   If HasCaption Then
      '(Barcode height + Text height)
      xObj.Height = (xObj.TextHeight(zBarText) + BarH) * Screen.TwipsPerPixelY
   Else
      xObj.Height = BarH * Screen.TwipsPerPixelY
   End If
   xObj.Width = (myCols.Count * 11) * Screen.TwipsPerPixelX
   Draw_Barcode
   zObj.Picture = zObj.Image
End Sub
Private Sub Draw_Barcode()
    Dim encoding As String, i As Integer, j As Integer
    xObj.CurrentX = 0: xObj.CurrentY = 0
    For i = 1 To myCols.Count
        encoding = arrEncoding(myCols(i))
        For j = 1 To Len(encoding)
            xObj.Line (0, 0)-(xObj.CurrentX + 1, BarH), IIf(Mid(encoding, ii, 1), vbBlack, vbWhite)
        Next
    Next
End Sub

Private Sub Init_Table()
    CharSet = " !""#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}~"
    arrEncoding = Array( _
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
             "11110101110", "11010000100", "11010010000", "11010011100", "11000111010", "11" _
             )
End Sub

Private Sub Eval_String(xstr As String)
    'Check the string for alpha-numeric
    '
    Dim i As Integer, num As Integer
    Dim tmpbuffer As String, StartCode As Integer
    
    
    Set myCols = New Collection
    num = 0: StartCode = 0
    tmpbuffer = ""

    For i = 1 To Len(xstr)
        If InStr("0123456789", Mid(xstr, i, 1)) > 0 Then
           num = num + 1
           tmpbuffer = tmpbuffer & Mid(xstr, i, 1)
           If num = 2 Then
              Add_Num tmpbuffer
              num = 0: tmpbuffer = ""
           End If
        Else
            If num = 1 Then
               Add_Char tmpbuffer
            ElseIf num = 2 Then
               Add_Num tmpbuffer
            End If
            
            Add_Char Mid(xstr, i, 1)
            num = 0: tmpbuffer = ""
        End If
    Next
    If num = 1 Then
       Add_Char tmpbuffer
    ElseIf num = 2 Then
       Add_Num tmpbuffer
    End If
    ' Add the Checksum
    Dim Calc As Long, chkSum As Long
    For i = 0 To myCols.Count - 1
        If i = 0 Then
           Calc = myCols(i + 1)
        Else
           Calc = Calc + (myCols(i + 1) * i)
        End If
    Next
    chkSum = Calc Mod 103
    If chkSum <> 0 Then Add_Num chkSum
    Add_Num StopChar
    Add_Num EndChar
End Sub
Private Sub Add_Char(xstr As String)
        If xStart <> StartB And xStart <> CodeB Then
           If xStart = 0 Then
              xStart = StartB
            Else
              xStart = CodeB
            End If
            myCols.Add xStart
            List1.AddItem xStart
            List2.AddItem arrEncoding(xStart)
        End If
        List1.AddItem xstr
        'List2.AddItem arrEncoding(InStr(CharSet, xstr) - 1)
        
        myCols.Add InStr(CharSet, xstr) - 1
End Sub
Private Sub Add_Num(xstr As String)
        If xStart <> StartC And xStart <> CodeC Then
           If xStart = 0 Then
              xStart = StartC
            Else
              xStart = CodeC
            End If
            myCols.Add xStart
            List1.AddItem xStart
            List2.AddItem arrEncoding(xStart)
        End If
        List1.AddItem xstr
        List2.AddItem arrEncoding(CInt(xstr))
        myCols.Add CInt(xstr)
End Sub




