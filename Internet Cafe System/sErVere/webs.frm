VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmWebServer 
   BorderStyle     =   0  'None
   Caption         =   "Webserver"
   ClientHeight    =   705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1155
   Icon            =   "webs.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   705
   ScaleWidth      =   1155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock Hoe 
      Index           =   0
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   80
   End
   Begin MSWinsockLib.Winsock Pimp 
      Left            =   600
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   80
   End
End
Attribute VB_Name = "frmWebServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private fso As New FileSystemObject

Public Function LoadFile(filename1 As String) As String
On Error GoTo Heaven

  If Not fso.FileExists(filename1) Then Err.Raise 76

  Open filename1 For Binary As #1
    LoadFile = Input(FileLen(filename1), #1)
  Close #1
  Exit Function
Heaven:
  LoadFile = " "
End Function

Private Sub Form_Load()
On Error Resume Next
Dim i As Integer
  Pimp.Close
  Pimp.LocalPort = 80
  Pimp.Listen
  DoEvents
  For i = 1 To 200
    Load Hoe(i)
  Next i
End Sub

Private Sub Hoe_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim strData As String
Dim strGet As String
Dim spc2 As Long
Dim page As String
  Hoe(Index).GetData strData
  If Mid(strData, 1, 3) = "GET" Then
    strGet = InStr(strData, "GET ")
    spc2 = InStr(strGet + 5, strData, " ")
    page = Trim(Mid(strData, strGet + 5, spc2 - (strGet + 4)))
    If Right(page, 1) = "/" Then page = Left(page, Len(page) - 1)
    
    If page = "/" Or page = "" Then page = "Index.html"
    
    If Mid(page, InStrRev(page, ".") + 1) = "htm" Then page = page & "l"
    
    If Not fso.FileExists(App.Path & "\" & page) Then page = "Error_404.html"
    
    Hoe(Index).SendData LoadFile(App.Path & "\" & page)
  End If
End Sub

Private Sub Hoe_SendComplete(Index As Integer)
  Hoe(Index).Close
End Sub

Private Sub Pimp_ConnectionRequest(ByVal requestID As Long)
Dim i As Integer
For i = 0 To 200
If Hoe(i).State = sckClosed Then
Hoe(i).Close
Hoe(i).Accept (requestID)
Exit Sub
End If
Next i
End Sub

