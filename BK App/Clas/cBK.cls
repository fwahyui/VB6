VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim SQL As String
Public IDKonseling As Integer
Public Jumlah As Long
Public tgl As Date
Public NIS As Integer
Public NIP As String, Permasalahan As String, Penanganan As String, Keterangan As String

Public Function Add _
(IDKonseling As Long, tgl As Date, NIS As Integer, NIP As String, _
Permasalahan As String, Penanganan As String, Keterangan As String) As Integer
    On Error GoTo Salah
    SQL = "INSERT INTO TBKonseling (IDKonseling,tgl,NIS,NIP,Permasalahan,"
    SQL = SQL & "Penanganan,Keterangan) VALUES (" & IDKonseling & ","
    SQL = SQL & "'" & tgl & "'," & NIS & ",'" & NIP & "','" & Permasalahan & "',"
    SQL = SQL & " '" & Penanganan & "','" & Keterangan & "')"
    DB_CONNECTION.Execute SQL
    Add = 1
    Exit Function
Salah:
    Add = 0
End Function

Public Function List(Optional pIsReadOnly As Boolean = True, Optional pWHERE As String = "") As ADODB.Recordset
On Error GoTo HELL
Dim RS As New ADODB.Recordset

SQL = "SELECT IDKonseling,tgl,NIS,NIP,Permasalahan,Penanganan,Keterangan FROM TBKonseling " & pWHERE
RS.CursorLocation = adUseClient
If pIsReadOnly Then
    RS.Open SQL, DB_CONNECTION, adOpenForwardOnly, adLockReadOnly
    Jumlah = RS.RecordCount
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


Public Function Load(pID As Long) As Integer
On Error GoTo HELL
Dim RS As New ADODB.Recordset

RS.CursorType = adOpenForwardOnly
RS.LockType = adLockReadOnly
SQL = "SELECT IDKonseling,tgl,NIS,NIP,Permasalahan,Penanganan,Keterangan FROM TBKonseling WHERE IDKonseling=" & pID & ""
Set RS = DB_CONNECTION.Execute(SQL)
With RS
    If .EOF Then
        Err.Raise BASE_ERROR, SRC_ERROR, "Data tidak ada"
    Else
        NIP = "" & !NIP
        NIS = "" & !NIS
        Permasalahan = "" & !Permasalahan
        Penanganan = "" & !Penanganan
        Keterangan = !Keterangan
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
Public Function Edit _
(IDKonseling As Long, tgl As Date, NIS As Integer, NIP As String, _
Permasalahan As String, Penanganan As String, Keterangan As String) As Integer
On Error GoTo Salah
    On Error GoTo Salah
    SQL = "UPDATE TBKonseling SET IDKonseling=" & IDKonseling & ",tgl='" & tgl & "',NIS=" & NIS & ",NIP='" & NIP & "',Permasalahan='" & Permasalahan & "',Penanganan='" & Penanganan & "',Keterangan='" & Keterangan & "' WHERE IDKonseling=" & IDKonseling & ""
    DB_CONNECTION.Execute SQL
    Edit = 1
    Exit Function
Salah:
    Edit = 0
End Function

Public Function Delete(pNIP As String) As Integer
On Error GoTo HELL

SQL = "DELETE FROM TBKonseling WHERE IDKonseling=" & IDKonseling & ""
DB_CONNECTION.Execute SQL
Delete = 1
Exit Function
HELL:
    Delete = 0
End Function

