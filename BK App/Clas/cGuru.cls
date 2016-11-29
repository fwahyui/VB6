VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGuru"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim SQL As String
Public NIP As String, Nama As String, Alamat As String, Jabatan As String

Public Function Add(NIP As String, Nama As String, Alamat As String, Jabatan As String) As Integer
    On Error GoTo Salah
    SQL = "INSERT INTO TBGuru (NIP,Nama,Alamat,Jabatan) VALUES "
    SQL = SQL & " ('" & NIP & "','" & Nama & "','" & Alamat & "','" & Jabatan & "')"
    DB_CONNECTION.Execute SQL
    Add = 1
    Exit Function
Salah:
    Add = 0
End Function

Public Function List(Optional pIsReadOnly As Boolean = True, Optional pWHERE As String = "") As ADODB.Recordset
On Error GoTo HELL
Dim RS As New ADODB.Recordset

SQL = "SELECT NIP,Nama,Alamat,Jabatan FROM TBGuru " & pWHERE
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


Public Function Load(pNIP As String) As Integer
On Error GoTo HELL
Dim RS As New ADODB.Recordset

RS.CursorType = adOpenForwardOnly
RS.LockType = adLockReadOnly
SQL = "SELECT NIP,Nama,Alamat,Jabatan FROM TbGuru WHERE NIP='" & pNIP & "'"
Set RS = DB_CONNECTION.Execute(SQL)
With RS
    If .EOF Then
        Err.Raise BASE_ERROR, SRC_ERROR, "Data tidak ada"
    Else
        NIP = "" & !NIP
        Nama = "" & !Nama
        Jabatan = "" & !Jabatan
        Alamat = "" & !Alamat
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
Public Function Edit(NIP As String, Nama As String, Alamat As String, Jabatan As String) As Integer
    On Error GoTo Salah
    On Error GoTo Salah
    SQL = "UPDATE TBGuru SET Nama='" & Nama & "',Alamat='" & Alamat & "',Jabatan='" & Jabatan & "' WHERE NIP='" & NIP & "'"
    DB_CONNECTION.Execute SQL
    Edit = 1
    Exit Function
Salah:
    Edit = 0
End Function

Public Function Delete(pNIP As String) As Integer
On Error GoTo HELL

SQL = "DELETE FROM TBGuru WHERE NIP='" & pNIP & "'"
DB_CONNECTION.Execute SQL
Delete = 1
Exit Function
HELL:
    Delete = 0
End Function


