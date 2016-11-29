Attribute VB_Name = "Module1"
Public sql As String
Public PidTahun As Integer
Public AppMain As New DLLBK.cCon
Public xCONN As New ADODB.Connection
Public Enum ENUM_DATA_MODE      'Enum ini dipakai sebagai status dokumen
    EN_NEW                      'Dokumen baru saja dibuat
    EN_SAVED                    'Dokumen baru saja disimpan
    EN_NEW_CHANGED              'Dokumen baru mengalami perubahan
    EN_LOAD_CHANGED             'Dokumen yang dibuka mengalami perubahan
End Enum
Public strTahun As String
Private Function connectDB(ByVal pSTR_Conn As String) As Integer
On Error GoTo Hell
xCONN.CursorLocation = adUseClient
xCONN.Open pSTR_Conn
connectDB = 1

Exit Function
Hell:
    MsgBox "Koneksi ke database gagal karena:" & vbCrLf & Err.Description, vbCritical
    connectDB = 0
End Function
Private Function isUserExist(pTableName As String, pSQLInsert As String) As Integer
On Error GoTo Hell
Dim tRS As New ADODB.Recordset
tRS.Open "select * from " & pTableName, xCONN, adOpenForwardOnly, adLockReadOnly
If tRS.RecordCount > 0 Then
    isUserExist = 1
Else
    xCONN.Execute pSQLInsert
    isUserExist = 1
    MsgBox "First user created", vbInformation
End If
Set tRS = Nothing
Exit Function
Hell:
    MsgBox Err.Description, vbCritical, "Internal"
End Function

Public Function koneksi()
    koneksi = "Provider = Microsoft.Jet.OLEDB.4.0; Data Source = " & App.Path & "\db\dbbk.mdb" & ";Persist Security Info=False;Jet OLEDB:Database Password=" & "jendeladunia"
End Function
Public Function form_ditengah(ByVal Index As MDIForm, ByVal child As Form)
    Dim kiri As Integer
    Dim atas As Integer
    kiri = (Index.Width - child.Width) / 2
    atas = (Index.Height - child.Height) / 2 - 500
    child.Left = kiri
    child.Top = atas
End Function
Public Sub Main()
On Error GoTo Hell
Dim STR_Conn, sql As String
STR_Conn = "Provider = Microsoft.Jet.OLEDB.4.0; Data Source = " & App.Path & "\db\dbbk.mdb" & ";Persist Security Info=False;Jet OLEDB:Database Password=" & "jendeladunia"
sql = "insert into tbuser (kodeuser,pwd,hakakses)values('ADMIN','samadengannamanya','Administrator')"
AppMain.OpenConnection (STR_Conn)

If (connectDB(STR_Conn) > 0) Then
    If (isUserExist("TBuser", sql) > 0) Then
        Index.Show
    End If
Else
    MsgBox "Koneksi ke database GAGAL." & vbCrLf & "Periksa apakah database ada. " & vbCrLf & _
        "Atau sedang digunakan aplikasi lain", vbCritical
    End
End If
LihatTahunAktiv
Exit Sub
Hell:
    MsgBox Err.Description, vbCritical, "Internal Error"
End Sub

Public Sub LihatTahunAktiv()
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.CursorLocation = adUseClient
sql = "Select * from TBTahunAjaran where aktif = -1 "
rs.Open sql, koneksi
If Not rs.EOF Then
    PidTahun = rs!IDTahunAjaran
    strTahun = rs!tahunajaran
    Index.Caption = Index.Caption & " Tahun Ajaran (" & strTahun & ")"
Else
    frmtahunajaran.Show
End If
End Sub
