VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'(c) 1999, 2000 by Joe Hart
'
'You do NOT have rights to redistribute this code, in whole or in part without
'my permission.  You also may not recompile the code and release
'it as another program without my permission.  If you would like to modify this code
'and distribute it in either as source code or as a compiled program please contact
'me at bghost@ti.cz before doing so.  I would appreciate being notified of any modifications
'even if you do not intend to redistribute it.
'
'No warrantees either expressed or implied are granted with this software.
'Any damage arrising from use of this software is your own risk.
'
'INI class
'
'Allows simple use of INI files.
'
Option Explicit

Private INI_FileName
Public Enum EDataType
    dtString
    dtBoolean
    dtNumeric
End Enum

Public Function GetString(SectionHeader As String, VariableName As String, Optional Default As String = "", Optional DataType As EDataType = dtString) As String
'get a value from the INI file
   On Error GoTo Hell
   Dim A$
   Dim Ret As String
   ' Blank the return string
   Ret = String(255, vbNullChar)
   'Get requested information, trimming the returned
   A$ = Left$(Ret, GetPrivateProfileString(SectionHeader, ByVal VariableName, Default, Ret, 255, INI_FileName))
   Select Case DataType
   Case dtString
      'do nothing as this is default
   Case dtBoolean
      'verify value contains true or false, default is true unless otherwise specified
      Dim B$
      B$ = UCase(A$)
      If Default = sEmpty Then Default = "True"
      If B$ <> "TRUE" And B$ <> "FALSE" Then A$ = Default
   Case dtNumeric
      'verify the value is a number
      If Val(A$) = 0 Then A$ = Default
   End Select
   'return the value
   GetString = A$
Hell:
End Function

Public Function SaveString(SectionHeader As String, VariableName As String, Value As String) As Integer
   SaveString = WritePrivateProfileString(SectionHeader, VariableName, Value, INI_FileName)
End Function

Private Sub Class_Initialize()
   INI_FileName = App.Path & "\" & App.EXEName & ".INI"
End Sub
