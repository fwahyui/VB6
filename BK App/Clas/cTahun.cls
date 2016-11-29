VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTahun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim SQL As String
Public IDTahunAjaran As Integer, Tahunajaran As String, Aktif As Byte
Public Function Add(pTahun As String) As Integer
    On Error GoTo Salah
    SQL = "Insert Into TBTahunAjaran (tahunajaran,aktif) values ('" & pTahun & "' ," & -1 & ")"
    DB_CONNECTION.Execute SQL
    Add = 1
    If Load(pTahun) > 0 Then
        
    End If
    Exit Function
Salah:
    Add = 0
End Function
Public Function Edit(pTahunajaran As String) As Integer
    On Error GoTo Salah
    On Error GoTo Salah
    SQL = "UPDATE TBTahunAjaran SET Aktif=" & 0 & " "
    DB_CONNECTION.Execute SQL
    SQL = "UPDATE TBTahunAjaran SET Tahunajaran='" & pTahunajaran & "',Aktif=" & -1 & " WHERE Tahunajaran='" & pTahunajaran & "'"
    DB_CONNECTION.Execute SQL
    
    Edit = 1
    If Load(pTahunajaran) > 0 Then
        
    End If
    Exit Function
Salah:
    Edit = 0
End Function
Public Function Load(pTahunajaran As String) As Integer
On Error GoTo HELL
Dim RS As New ADODB.Recordset

RS.CursorType = adOpenForwardOnly
RS.LockType = adLockReadOnly
SQL = "SELECT IDTahunAjaran,Tahunajaran,Aktif FROM TBTahunAjaran WHERE Tahunajaran='" & pTahunajaran & "'"
Set RS = DB_CONNECTION.Execute(SQL)
With RS
    If .EOF Then
        Load = 0
        Exit Function
    Else
        IDTahunAjaran = !IDTahunAjaran
        Tahunajaran = !Tahunajaran
        Aktif = !Aktif
    End If
    .Close
End With
Set RS = Nothing
Load = 1
Exit Function
HELL:
    Set RS = Nothing
    Err.Raise BASE_ERROR, SRC_ERROR, Err.Description
End Function
