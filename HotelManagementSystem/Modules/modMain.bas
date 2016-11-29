Attribute VB_Name = "modMain"
Option Explicit

Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'api declarations
Public Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Public CloseMe  As Boolean

Public Sub Main()
    'use system appearance style
    InitCommonControls
    
    frmSplash.Show
    frmSplash.Refresh

    DBPath = GetINI("Configuration", "Path")      'get path from file
    If Trim(DBPath) = "" Or IsNull(DBPath) Then
JumpHere:
      frmLocate.Show 1                            'browse database
    End If
    
    If OpenDB = vbRetry Then GoTo JumpHere
    
    'create DSN for reports
    createDSN

'    Load mdiMain
    
    Delay 2

    frmLogin.Show 1
    
    If CloseMe = True Then End

    mdiMain.Show
    
    Unload frmSplash
    Set frmSplash = Nothing
End Sub

'Public Sub Main_AfterSD()
'
'
'    'Open Database File
'    If OpenDB = False Then
'        Exit Sub
'    End If
'
'
'    'TestUnit
'    mdiMain.ShowForm
'End Sub

Public Sub SetINI(strMain As String, strSub As String, strvalue As String)
    WritePrivateProfileString strMain, strSub, strvalue, App.Path & "\config.txt"
End Sub

Public Sub Delay(PauseTime)
    Dim Start, Finish, TotalTime

    Start = Timer   ' Set start time.
    Do While Timer < Start + PauseTime
        DoEvents    ' Yield to other processes.
    Loop
End Sub

