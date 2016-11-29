Attribute VB_Name = "modPublicVar"
Option Explicit


Public CurrUser                     As USER_INFO
Public DBPath                       As String
Public Enc                          As New clsBlowfish
Public CurrBiz                      As BUSINESS_INFO

Public CN                           As New Connection

Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long

Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
    Public Const IDC_HAND = 32649&
    Public Const IDC_ARROW = 32512&

