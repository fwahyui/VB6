VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSoal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim SQL As String
Public No As Integer, Soal As String

Public Function Add(pNo As Integer, pSoal As String) As Integer
    On Error GoTo Salah
    SQL = "INSERT INTO tbSoal (nosoal,soal) VALUES (" & pNo & ",'" & pSoal & "')"
    DB_CONNECTION.Execute SQL
    Add = 1
    Exit Function
Salah:
    Add = 0
End Function

Public Function List(Optional pIsReadOnly As Boolean = True, Optional pWHERE As String = "") As ADODB.Recordset
On Error GoTo HELL
Dim RS As New ADODB.Recordset

SQL = "SELECT nosoal as [No Soal],soal FROM tbSoal " & pWHERE
RS.CursorLocation = adUseClient
If pIsReadOnly Then
    RS.Open SQL, DB_CONNECTION, adOpenForwardOnly, adLockReadOnly
Else
    RS.Open SQL, DB_CONNECTION, adOpenDynamic, adLockOptimistic
End If

Set List = RS
Set RS = Nothing
Exit Function
HELL:
    Set RS = Nothing
    Err.Raise BASE_ERROR, SRC_ERROR, Err.Description
End Function


Public Function Load(pNo As Integer) As Integer
On Error GoTo HELL
Dim RS As New ADODB.Recordset

RS.CursorType = adOpenForwardOnly
RS.LockType = adLockReadOnly
SQL = "SELECT Nosoal,Soal FROM TBsoal WHERE Nosoal=" & pNo & ""
Set RS = DB_CONNECTION.Execute(SQL)
With RS
    If .EOF Then
        Err.Raise BASE_ERROR, SRC_ERROR, "Data tidak ada"
    Else
        No = !Nosoal
        Soal = !Soal
    End If
    .Close
End With
Set RS = Nothing
Load = 1
Exit Function
HELL:
    Set RS = Nothing
    Load = 0
    Err.Raise BASE_ERROR, SRC_ERROR, Err.Description
End Function
Public Function Edit(pNo As Integer, pSoal As String) As Integer
    On Error GoTo Salah
    On Error GoTo Salah
    SQL = "UPDATE tbSoal SET nosoal=" & pNo & ",soal='" & pSoal & "' WHERE nosoal=" & pNo & ""
    DB_CONNECTION.Execute SQL
    Edit = 1
    Exit Function
Salah:
    Edit = 0
End Function

Public Function Delete(pNo As String) As Integer
On Error GoTo HELL

SQL = "DELETE FROM TBSoal WHERE nosoal='" & pNo & "'"
DB_CONNECTION.Execute SQL
Delete = 1
Exit Function
HELL:
    Delete = 0
End Function



