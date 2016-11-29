Attribute VB_Name = "Code39"
Dim CharSet As String
Dim arrEncoding() As Variant
Option Explicit

Dim zBarH As Long
Dim zBarText As String
Dim zObj As Object
Dim zHasCaption As Boolean
Dim zWithCheckSum As Boolean

Private Const ChkChar = 43
Dim myCols As Collection

Sub Bar39(Obj As Object, BarH As Double, BarText As String, Optional WithCheckSum As Boolean = False, Optional ByVal HasCaption As Boolean = False)
   Set zObj = Obj
   zWithCheckSum = WithCheckSum
   zBarText = BarText
   zHasCaption = HasCaption
   zBarH = BarH * 72 'Inches to Pixel of barcode
   
   Init_Table
   
   If Not CheckCode Then Exit Sub 'Check the String if Valid in Code 39
   
   Eval_String
   
   zObj.Picture = Nothing
   zObj.BackColor = vbWhite
   zObj.AutoRedraw = True
   zObj.ScaleMode = 3
   
   If zHasCaption Then
      zObj.Height = (zObj.TextHeight(zBarText) + zBarH + 5) * Screen.TwipsPerPixelY
   Else
      zObj.Height = zBarH * Screen.TwipsPerPixelY
   End If
   
   zObj.Height = zObj.Height + 10 ' Border
   zObj.Width = ((myCols.Count + 1) * 12) * Screen.TwipsPerPixelX
   
   Draw_Barcode
   zObj.Picture = zObj.Image
End Sub
Function CheckCode() As Boolean
    'Allowed Character sets are only those specified in Charset Variable less the (*)
    Dim ii As Integer
    zBarText = Replace(zBarText, "*", "")
    For ii = 1 To Len(zBarText)
        If InStr(CharSet, Mid(zBarText, ii, 1)) = 0 Then
           GoTo Err_Found
        End If
    Next
    CheckCode = True
    Exit Function
Err_Found:
    Err.Raise vbObjectError + 513, "Bar 39", _
      "An Invalid Character Found in Bar Text"
    CheckCode = False
End Function
Private Sub Eval_String()
    Dim i As Long, chkSum As Integer, xTotal As Integer, posCtr As Integer
    Set myCols = New Collection
 
    xTotal = 0
    posCtr = 0
    
    myCols.Add ChkChar 'Start of Barcode
    
    For i = 1 To Len(zBarText)
        posCtr = InStr(CharSet, Mid(zBarText, i, 1)) - 1
        xTotal = xTotal + posCtr
        myCols.Add posCtr
    Next
    
    chkSum = xTotal Mod 43
    If zWithCheckSum Then myCols.Add chkSum 'Check sum
    myCols.Add ChkChar 'End of Barcode
End Sub
Private Sub Draw_Barcode()
    Dim encoding As String, i As Integer, j As Integer, xPos As Integer
    xPos = 5 'zBorder / 2
    For i = 1 To myCols.Count
        encoding = arrEncoding(myCols(i))
        For j = 1 To Len(encoding)
            xPos = xPos + 1
            zObj.Line (xPos, 5)-(xPos, zBarH), IIf(Mid(encoding, j, 1), vbBlack, vbWhite)
        Next
    Next
    
    If zHasCaption Then
        zObj.CurrentX = ((myCols.Count * 12) - zObj.TextWidth(zBarText)) / 2   '(zObj.Width - zObj.TextWidth(zBarText) / 2)    ' Horizontal position.
        zObj.CurrentY = zObj.CurrentY + 5    ' Vertical position.
        zObj.Print zBarText   ' Print message.
    End If
End Sub
Private Sub Init_Table()
    CharSet = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%*"
    arrEncoding = Array( _
             "101001101101", "110100101011", "101100101011", "110110010101", "101001101011", "110100110101", _
             "101100110101", "101001011011", "110100101101", "101100101101", "110101001011", "101101001011", _
             "110110100101", "101011001011", "110101100101", "101101100101", "101010011011", "110101001101", _
             "101101001101", "101011001101", "110101010011", "101101010011", "110110101001", "101011010011", _
             "110101101001", "101101101001", "101010110011", "110101011001", "101101011001", "101011011001", _
             "110010101011", "100110101011", "110011010101", "100101101011", "110010110101", "100110110101", _
             "100101011011", "110010101101", "100110101101", "100100100101", "100100101001", "100101001001", _
             "101001001001", "100101101101" _
             )
End Sub






