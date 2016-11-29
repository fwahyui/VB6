VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Ckelas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim SQL As String
Public IDTahunAjaran As Integer
Public IDKelas As Integer
Public Kelas As String, Ruangan As String
Public JumlahSiswa As Integer

Public Function Add(IDTahunAjaran As Integer, IDKelas As Integer, Kelas As String, _
                Ruangan As String, JumlahSiswa As Integer) As Integer
    On Error GoTo Salah
    SQL = "INSERT INTO TBkelas (IDTahunAjaran,IDKelas,Kelas," & _
    " Ruangan,JumlahSiswa) VALUES (" & IDTahunAjaran & "," & _
    " " & IDKelas & ",'" & Kelas & "','" & Ruangan & "'," & _
    " " & JumlahSiswa & ")"
    DB_CONNECTION.Execute SQL
    Add = 1
    Exit Function
Salah:
    Add = 0
End Function

Public Function List(Optional pIsReadOnly As Boolean = True, Optional pWHERE As String = "") As ADODB.Recordset
On Error GoTo HELL
Dim RS As New ADODB.Recordset

SQL = "SELECT IDTahunAjaran,IDKelas,Kelas,Ruangan,JumlahSiswa FROM TBkelas " & pWHERE
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


Public Function Load(pIDKelas As Integer) As Integer
On Error GoTo HELL
Dim RS As New ADODB.Recordset

RS.CursorType = adOpenForwardOnly
RS.LockType = adLockReadOnly
SQL = "SELECT IDTahunAjaran,IDKelas,Kelas,Ruangan,JumlahSiswa FROM TBkelas WHERE IDKelas=" & pIDKelas & ""
Set RS = DB_CONNECTION.Execute(SQL)
With RS
    If .EOF Then
        Err.Raise BASE_ERROR, SRC_ERROR, "Data tidak ada"
    Else
        IDTahunAjaran = !IDTahunAjaran
        IDKelas = !IDKelas
        Kelas = !Kelas
        Ruangan = !Ruangan
        JumlahSiswa = !JumlahSiswa
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
Public Function Edit(IDTahunAjaran As Integer, IDKelas As Integer, Kelas As String, _
                Ruangan As String, JumlahSiswa As Integer) As Integer
    On Error GoTo Salah
    SQL = "UPDATE TBkelas SET IDTahunAjaran=" & IDTahunAjaran & ",jumlahsiswa=" & JumlahSiswa & ",Kelas='" & Kelas & "',Ruangan='" & Ruangan & "' WHERE IDKelas=" & IDKelas & ""
    DB_CONNECTION.Execute SQL
    Edit = 1
    Exit Function
Salah:
    Edit = 0
End Function

Public Function Delete(pIDKelas As Integer) As Integer
On Error GoTo HELL

SQL = "DELETE FROM TBkelas WHERE IDKelas=" & pIDKelas & ""
DB_CONNECTION.Execute SQL
Delete = 1
Exit Function
HELL:
    Delete = 0
End Function

Public Function ListCombo(Optional pWHERE As String = "") As ADODB.Recordset
On Error GoTo HELL
Dim RS As New ADODB.Recordset

SQL = "SELECT Kelas,Ruangan FROM TBkelas " & pWHERE
RS.CursorLocation = adUseClient
If pIsReadOnly Then
    RS.Open SQL, DB_CONNECTION, adOpenDynamic, adLockOptimistic
Else
    RS.Open SQL, DB_CONNECTION, adOpenDynamic, adLockOptimistic
End If

Set ListCombo = RS
Set RS = Nothing
Exit Function
HELL:
    Set RS = Nothing
    Err.Raise BASE_ERROR, SRC_ERROR, Err.Description
End Function


