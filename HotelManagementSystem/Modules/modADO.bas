Attribute VB_Name = "modADO"
Option Explicit

Public Function OpenDB() As Boolean
    Dim isOpen      As Boolean
    Dim ANS         As VbMsgBoxResult
    isOpen = False
    On Error GoTo err
    
    Do Until isOpen = True
      CN.CursorLocation = adUseClient
            
      CN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBPath & ";Persist Security Info=False;Jet OLEDB:Database Password=jaypee"
      isOpen = True
    Loop
    OpenDB = isOpen
    Exit Function
err:
    ANS = MsgBox("Error Number: " & err.Number & vbCrLf & "Description: " & err.Description, _
  vbCritical + vbRetryCancel)
  If ANS = vbCancel Then
    OpenDB = vbCancel
  ElseIf ANS = vbRetry Then
    OpenDB = vbRetry
  End If
End Function

Public Sub CloseDB()
    'Close the connection
    CN.Close
    Set CN = Nothing
End Sub

'Function that return the current index for a certain table
Public Function getIndex(ByVal srcTable As String) As Long
    On Error GoTo err
    Dim RS As New Recordset
    Dim RI As Long
    
    RS.CursorLocation = adUseClient
    RS.Open "SELECT * FROM [KEY GENERATOR] WHERE TableName = '" & srcTable & "'", CN, adOpenStatic, adLockOptimistic
    
    RI = RS.Fields("NextNo")
    CN.BeginTrans
    RS.Fields("NextNo") = RI + 1
    RS.Update
    CN.CommitTrans
    getIndex = RI
    
    srcTable = ""
    RI = 0
    Set RS = Nothing
    Exit Function
err:
        ''Error when incounter a null value
        If err.Number = 94 Then
            getIndex = 1
            Resume Next
        Else
            MsgBox err.Description
        End If
        CN.RollbackTrans
End Function

'Function used to get the sum  of fields
Public Function getSumOfFields(ByVal sTable As String, ByVal sField As String, ByRef sCN As ADODB.Connection, Optional inclField As String, Optional sCondition As String) As Double
    On Error GoTo err
    Dim RS As New ADODB.Recordset

    RS.CursorLocation = adUseClient
    If sCondition <> "" Then sCondition = " GROUP BY " & inclField & " HAVING(" & sCondition & ")"
    If inclField <> "" Then inclField = "," & inclField
    RS.Open "SELECT Sum(" & sTable & "." & sField & ") AS fTotal" & inclField & " FROM " & sTable & sCondition, sCN, adOpenStatic, adLockOptimistic
    If RS.RecordCount > 0 Then
        RS.MoveFirst
        Do While Not RS.EOF
            getSumOfFields = getSumOfFields + RS.Fields("fTotal")
            RS.MoveNext
        Loop
    Else
        getSumOfFields = 0
    End If
    
    Set RS = Nothing
    Exit Function
err:
        'Error when incounter a null value
        If err.Number = 94 Then getSumOfFields = 0: Resume Next
End Function

'Procedure used to generate DSN
Public Sub GenerateDSN()
Open App.Path & "\rptCN.dsn" For Output As #1
    Print #1, "[ODBC]"
    Print #1, "DRIVER=Microsoft Access Driver (*.mdb)"
    Print #1, "UID=admin"
    Print #1, "UserCommitSync=Yes"
    Print #1, "Threads=3"
    Print #1, "SafeTransactions=0"
    Print #1, "PageTimeout=5"
    Print #1, "MaxScanRows=8"
    Print #1, "MaxBufferSize=2048"
    Print #1, "FIL=MS Access"
    Print #1, "DriverId=25"
    Print #1, "DefaultDir=" & App.Path & "\Data"
    Print #1, "DBQ=" & App.Path & "\Data\Data.mdb"
Close #1
End Sub

'Procedure used to remove DSN
Public Sub RemoveDSN()
On Error Resume Next
Kill App.Path & "\rptCN.dsn"
End Sub


