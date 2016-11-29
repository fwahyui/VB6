Attribute VB_Name = "EAN13and8"
Option Explicit
Dim LeftHand_Odd() As Variant
Dim LeftHand_Even() As Variant
Dim Right_Hand() As Variant
Dim Parity() As Variant
Dim zBarH As Long
Dim zBarText As String
Dim zObj As Object
Dim zHasCaption As Boolean
Dim zBarType As String 'EAN13 atau EAN8
Private Const ChkChar = 43
Dim xPos As Long, xTotal As Integer
Dim StartX As Integer, EndX As Integer
Dim chkSum  As Integer
Sub BarEAN(BarType As String, Obj As Object, BarH As Integer, BarText As String, Optional ByVal HasCaption As Boolean = False)
   Set zObj = Obj
   Init_Table
   
   zBarText = BarText
   zHasCaption = HasCaption
   zBarType = BarType
   zBarH = BarH * 72
   
   If Not CheckCode Then Exit Sub
   
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
   zObj.Width = (Len(zBarText) * 7) + (zObj.TextWidth(Mid(zBarText, 1, 1)) * 2) + 30
   zObj.Width = zObj.Width * Screen.TwipsPerPixelX
   
   If zBarType = "EAN13" Then
      Paint_Bar13 zBarText
   Else
      Paint_Bar8 zBarText
   End If
   
   zObj.Picture = zObj.Image
End Sub
Function CheckCode() As Boolean
    Dim ii As Integer, J As Integer
    J = IIf(zBarType = "EAN13", 12, 7)
    If Len(zBarText) <> J Then
        Err.Raise vbObjectError + 513, zBarType, _
          "Should be " & J & " Digit Numbers"
        GoTo Err_Found
    End If
    For ii = 1 To Len(zBarText)
        If InStr("0123456789", Mid(zBarText, ii, 1)) = 0 Then
            Err.Raise vbObjectError + 513, zBarType, _
              "An Invalid Character Found in Bar Text"
           GoTo Err_Found
        End If
    Next
    CheckCode = True
    Exit Function
Err_Found:
    CheckCode = False
End Function
Private Sub Paint_Bar13(ByVal xstr As String)
    Dim ii As Long, jj As Integer, ctr As Integer, xEven As Boolean, xParity As String
 
    xTotal = 0
    xPos = 11
    
    If zHasCaption Then
        zObj.CurrentX = xPos
        zObj.CurrentY = zBarH - zObj.TextHeight(zBarText)
        zObj.Print Mid(xstr, 1, 1)
        xPos = xPos + zObj.TextWidth(Mid(xstr, 1, 1)) + 1
    End If
    Draw_Bar "101", True
    StartX = zObj.CurrentX
    
    zObj.CurrentY = 15 + zBarH
    xParity = Parity(CInt(Mid(xstr, 1, 1)))
    
    
    For ii = 1 To Len(xstr)
        If ((Len(xstr) + 1) - ii) Mod 2 = 0 Then
           xTotal = xTotal + (CInt(Mid(xstr, ii, 1)))
        Else ' Odd
           xTotal = xTotal + CInt(Mid(xstr, ii, 1) * 3)
        End If
        If ii = 8 Then
           Draw_Bar "01010", True
        End If
        jj = CInt(Mid(xstr, ii, 1))
        If ii > 1 And ii < 8 Then
           Draw_Bar CStr(IIf(Mid(xParity, ii - 1, 1) = "E", LeftHand_Even(jj), LeftHand_Odd(jj))), False
        ElseIf ii > 1 And ii >= 8 Then
           Draw_Bar CStr(Right_Hand(jj)), False
        End If
    Next
    chkSum = 0
    jj = xTotal Mod 10
    If jj <> 0 Then
       chkSum = 10 - jj
    End If
    Draw_Bar CStr(Right_Hand(chkSum)), False
    
    EndX = zObj.CurrentX
    Draw_Bar "101", True
    
   
   If zHasCaption Then
        Dim centerX As Double
        centerX = ((EndX - StartX) / 2)
        zObj.CurrentX = ((centerX - zObj.TextWidth(Mid(xstr, 2, 6))) / 2) + StartX  '23
        zObj.CurrentY = 2 + zBarH
        zObj.Print Mid(xstr, 2, 6)
        
        zObj.CurrentX = ((centerX - zObj.TextWidth(Mid(xstr, 8, 6))) / 2) + StartX + centerX
        zObj.CurrentY = 2 + zBarH
        zObj.Print Mid(xstr, 8, 6) & chkSum
    End If
End Sub
Private Sub Paint_Bar8(ByVal xstr As String)
    Dim ii As Long, jj As Integer, ctr As Integer, xEven As Boolean, xParity As String
 
    xTotal = 0
    xPos = 11
    
    
    Draw_Bar "101", True ' Start
    StartX = zObj.CurrentX
    
    zObj.CurrentX = xPos
    zObj.CurrentY = zBarH + 15
    xParity = Parity(7) 'CInt(Mid(xstr, 1, 1)))
    
    
    For ii = 1 To Len(xstr)
        If ((Len(xstr) + 1) - ii) Mod 2 = 0 Then 'EVEN
           xTotal = xTotal + (CInt(Mid(xstr, ii, 1)))
        Else ' Odd
           xTotal = xTotal + CInt(Mid(xstr, ii, 1) * 3)
        End If
        If ii = 5 Then
           Draw_Bar "01010", True ' Middle
        End If
        jj = CInt(Mid(xstr, ii, 1))
        If ii < 5 Then
           Draw_Bar CStr(LeftHand_Odd(jj)), False
        ElseIf ii >= 5 Then
           Draw_Bar CStr(Right_Hand(jj)), False
        End If
    Next
    chkSum = 0
    jj = xTotal Mod 10
    If jj <> 0 Then
       chkSum = 10 - jj
    End If
    Draw_Bar CStr(Right_Hand(chkSum)), False
    
    EndX = zObj.CurrentX
    Draw_Bar "101", True
    
    If zHasCaption Then
        Dim centerX As Double
        centerX = ((EndX - StartX) / 2)
        zObj.CurrentX = ((centerX - zObj.TextWidth(Mid(xstr, 1, 4))) / 2) + StartX
        zObj.CurrentY = 2 + zBarH
        zObj.Print Mid(xstr, 1, 4)
        
        zObj.CurrentX = ((centerX - zObj.TextWidth(Mid(xstr, 4, 4))) / 2) + StartX + centerX
        zObj.CurrentY = 2 + zBarH
        zObj.Print Mid(xstr, 5, 3) & chkSum
    End If
End Sub

Private Sub Draw_Bar(encoding As String, Guard As Boolean)
    Dim ii As Integer
    For ii = 1 To Len(encoding)
        xPos = xPos + 1
        zObj.Line (xPos, 10)-(xPos, zBarH + IIf(Guard, 5, 0)), IIf(Mid(encoding, ii, 1), vbBlack, vbWhite)
    Next
End Sub
Private Sub Init_Table()
    LeftHand_Odd = Array("0001101", "0011001", "0010011", "0111101", "0100011", "0110001", "0101111", "0111011", "0110111", "0001011")
    LeftHand_Even = Array("0100111", "0110011", "0011011", "0100001", "0011101", "0111001", "0000101", "0010001", "0001001", "0010111")
    Right_Hand = Array("1110010", "1100110", "1101100", "1000010", "1011100", "1001110", "1010000", "1000100", "1001000", "1110100")
    Parity = Array("OOOOOO", "OOEOEE", "OOEEOE", "OOEEEO", "OEOOEE", "OEEOOE", "OEEEOO", "OEOEOE", "OEOEEO", "OEEOEO")
End Sub
