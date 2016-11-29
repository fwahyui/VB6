Attribute VB_Name = "Module1"
Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function ExitWindowsEx& Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long)
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessID As Long, ByVal dwType As Long) As Long
Public Const RSP_SIMPLE_SERVICE = 1
Public Const RSP_UNREGISTER_SERVICE = 0
Public DbConn As ADODB.Connection
Public DbRst As ADODB.Recordset
Public StrConn As String
Public TypeTrans As Boolean 'Open=True;Limited=False
Public Mark As Integer
Public Shells As Object
Public Locked As Boolean
Public ExtraCommand As String
Public Cur_ID As Byte

Private Declare Function SystemParametersInfo Lib _
"user32" Alias "SystemParametersInfoA" (ByVal uAction _
As Long, ByVal uParam As Long, ByVal lpvParam As Any, _
ByVal fuWinIni As Long) As Long

Public Sub DoDrag(TheForm As Form)
  ReleaseCapture
  SendMessage TheForm.hwnd, &HA1, 2, 0&
End Sub

Public Sub Connect_DB()
  Set DbConn = New ADODB.Connection
  With DbConn
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Config.mdb;Jet OLEDB:Database Password=GallanosA;"
    .CommandTimeout = 0
    .CursorLocation = adUseClient
    .Open
  End With
End Sub

Sub Main()
  'filters to run only one app itself
  If App.PrevInstance = True Then End
  
  Locked = True
  Connect_DB

  Set DbRst = New ADODB.Recordset
  StrConn = "SELECT * FROM tblClient"
  DbRst.Open StrConn, DbConn, adOpenDynamic, adLockOptimistic
  
  Cur_ID = DbRst!CompID
  
  If DbRst!SvrIP = "143" Then
    frmConfig.Show
  Else
    frmMain.Show
  End If
End Sub

Public Function Check_Data(abc As String) As String
Dim i As Byte
Dim tmp As Byte
Dim loc As Byte
  
  tmp = 0
  loc = 0
  For i = 3 To Len(abc)
    If Mid(abc, i, 1) = "@" Then
      tmp = tmp + 1
      If loc = 0 Then loc = i
    End If
  Next i

  If tmp = 1 Then
    Check_Data = abc
  Else
    ExtraCommand = Mid(abc, loc + 1, Len(abc) - loc)
    Check_Data = Mid(abc, 1, loc)
  End If
End Function

Sub BlockCtrl_Alt_Del(bDisabled As Boolean)
  Dim X As Long
  X = SystemParametersInfo(97, bDisabled, CStr(1), 0)

End Sub

Public Sub RemoveProgramFromList()
  Dim lngProcessID As Long
  Dim lngReturn As Long
 
  lngProcessID = GetCurrentProcessId()
  'lngReturn = RegisterServiceProcess(0, RSP_SIMPLE_SERVICE)
End Sub




