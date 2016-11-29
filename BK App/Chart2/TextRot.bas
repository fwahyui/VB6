Attribute VB_Name = "basTextRot"
Option Explicit
Public uDisplayDescript  As Boolean      'Display description when selectable

'API Constants:
Private Const LF_FACESIZE   As Long = 32&
Private Const SYSTEM_FONT   As Long = 13&
Private Const ANTIALIASED_QUALITY = 4

'Type Structures:
Private Type PointAPI
    X   As Long
    Y   As Long
End Type

Private Type SizeStruct
    Width   As Long
    Height  As Long
End Type

Private Type LOGFONT
    lfHeight            As Long
    lfWidth             As Long
    lfEscapement        As Long
    lfOrientation       As Long
    lfWeight            As Long
    lfItalic            As Byte
    lfUnderline         As Byte
    lfStrikeOut         As Byte
    lfCharSet           As Byte
    lfOutPrecision      As Byte
    lfClipPrecision     As Byte
    lfQuality           As Byte
    lfPitchAndFamily    As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

'API Declarations:
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SizeStruct) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Public Function PrintRotText(ByVal hDC As Long, ByVal Text As String, ByVal CenterX As Long, ByVal CenterY As Long, ByVal RotDegrees As Single) As Boolean
'Parameters:
'
'hDC = Device context where printing will occur.
'       This may be any object with an hDC (Form,
'       PictureBox, UserControl, etc.)
'
'Text = Text string to be printed.
'
'CenterX, CenterY = Center point of text in pixels.
'
'RotDegrees = Rotation amount in degrees (0.0 to 359.9999999)
'   (counter-clockwise; zero = horizontal (no rotation)).

Dim bOkSoFar    As Boolean      'Flag to continue.
Dim hFontOld    As Long         'Handle to original font.
Dim hFontNew    As Long         'Handle to new font.
Dim lfFont      As LOGFONT      'LOGFONT structure for new font.
Dim ptOrigin    As PointAPI     'Point of origin for drawing text.
Dim ptCenter    As PointAPI     'Center point of text.
Dim szText      As SizeStruct   'Width and Height of text.

    'Get the current LOGFONT structure from the device.
    'To accomplish this, first select a stock font into the
    'device, which will return a handle to it's current font.
    hFontOld = SelectObject(hDC, GetStockObject(SYSTEM_FONT))
    
    'If successful getting the font from the device...
    If hFontOld <> 0 Then
        
        'Now get the LOGFONT structure from the font.
        bOkSoFar = (GetObjectAPI(hFontOld, Len(lfFont), lfFont) <> 0)
        
        'Put the original font back into the device.
        Call SelectObject(hDC, hFontOld)
        
        'Reset for use later
        hFontOld = 0
    End If
    
    'Continue only if successful getting the LOGFONT structure.
    If bOkSoFar Then
        'Change the escapement and orientation of the font.
        lfFont.lfEscapement = RotDegrees * 10
        lfFont.lfOrientation = lfFont.lfEscapement
        lfFont.lfQuality = ANTIALIASED_QUALITY
        
        'Now create a font object from the LOGFONT structure.
        hFontNew = CreateFontIndirect(lfFont)
        
        'If font creation was successful...
        If hFontNew <> 0 Then
            'Select the new font into the device.
            hFontOld = SelectObject(hDC, hFontNew)
            'If successful selecting the new font into the device...
            If hFontOld <> 0 Then
                'Get the size of the text in logical units (pixels).
                bOkSoFar = (GetTextExtentPoint32(hDC, Text, Len(Text), szText) <> 0)
                
                'If successful getting the size of the text...
                If bOkSoFar Then
                    'Calculate the point of origin for the text
                    'as it would be if the text was horizontal.
                    With ptOrigin
                        .X = CenterX - (szText.Width / 2)
                        .Y = CenterY - (szText.Height / 2)
                    End With
                    
                    'Convert CenterX, CenterY to a point structure
                    '(needed for call to RotatePoint).
                    With ptCenter
                        .X = CenterX
                        .Y = CenterY
                    End With
                    
                    'Rotate the point of origin to match
                    'the desired rotation (RotDegrees).
                    Call RotatePoint(ptCenter, ptOrigin, RotDegrees)
                
                    'Now Print the rotated text and return success/failure.
                    PrintRotText = (TextOut(hDC, ptOrigin.X, _
                      ptOrigin.Y, Text, Len(Text)) <> 0)
                
                End If
                'Put the original font back into the device.
                hFontNew = SelectObject(hDC, hFontOld)
            End If
            'Clean up memory by deleting the created font.
            Call DeleteObject(hFontNew)
        End If
    End If
            
End Function

Private Sub RotatePoint(ptAxis As PointAPI, ptRotate As PointAPI, fDegrees As Single)

' ***************************************************
' *                 RotatePoint                     *
' *                                                 *
' *  Created by: Rocky Clark (Kath-Rock Software)   *
' *                                                 *
' *  Rotate ptRotate around ptAxis, fDegrees from   *
' *  its current position.                          *
' *                                                 *
' * This procedure may be used and distributed, as  *
' * is, in your code, as long as these credits and  *
' * the code itself remain unchanged.               *
' *                                                 *
' ***************************************************

Dim fDX     As Single   'Delta X
Dim fDY     As Single   'Delta Y
Dim fRads   As Single   'Radians
Const dPi   As Double = 3.14159265358979 'Pi


    'Convert degrees to radians.
    fRads = fDegrees * (dPi / 180#)
    
    'Calculate the deltas from the center point.
    fDX = ptRotate.X - ptAxis.X
    fDY = ptRotate.Y - ptAxis.Y
    
    'Rotate the point.
    ptRotate.X = ptAxis.X + ((fDX * Cos(fRads)) + (fDY * Sin(fRads)))
    ptRotate.Y = ptAxis.Y + -((fDX * Sin(fRads)) - (fDY * Cos(fRads)))
    
End Sub

