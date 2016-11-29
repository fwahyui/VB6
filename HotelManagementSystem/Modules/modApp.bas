Attribute VB_Name = "modApp"
Option Explicit

Private Const sAS_AutoBackup As String = "AutoBackup"

'Public Function WriteErrorLog(sModuleName As String, sRoutineName As String, sDetail As String)
'    frmErrMsg.ShowForm sModuleName, sRoutineName, sDetail
'End Function

Public Function SetAutoBackup(ByVal NewValue As Boolean)

    Dim sValue As String
    
    SaveSetting App.Title, "AppSetting", sAS_AutoBackup, IIf(NewValue, "T", "F")

End Function


Public Function GetAutoBackup() As Boolean
    
    Dim sValue As String
    
    'default
    GetAutoBackup = -1
    
    sValue = GetSetting(App.Title, "AppSetting", sAS_AutoBackup, "T")

    GetAutoBackup = IIf(sValue = "T", True, False)

End Function

