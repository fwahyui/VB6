VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Csiswa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public NIS As Long, anakke As Integer, jumlahsaudara As Integer, saudarakandung As Integer
Public saudaratiri As Integer, beratbadan As Integer, tinggibandan As Integer
Public Nama As String, alamatortu As String, pendidikanortu As String, pekerjaanortu As String, sd As String, alamatsd As String
Public Tempatlahir As String
Public tgllahir As Date
Public jk As String
Public agama As String
Public Alamat As String
Public namaortu As String
Public hubungan As String
Public agamaortu As String
Public smp As String
Public alamatsmp As String
Public nonformalsatu As String
Public alamatnonformalsatu As String
Public nonformaldua As String
Public alamatnonformaldua As String
Public nonformaltiga As String
Public alamatnonformaltiga As String
Public prestasi As String
Public prestasidua As String
Public prestasitiga As String
Public prestasiempat As String
Public prestasilima As String
Public kegiatan As String
Public kegiatan1 As String
Public kegiatan2 As String
Public penyakit1 As String
Public penyakit2 As String
Public penyakit3 As String
Public Notelp As String
Public IDkelas As Integer
'___________________________
Dim SQL As String


Public Function Delete(pNIS As Long) As Integer
On Error GoTo HELL

SQL = "DELETE FROM TBSiswa WHERE NIS=" & pNIS & ""
DB_CONNECTION.Execute SQL
Delete = 1
Exit Function
HELL:
    Delete = 0
End Function

Public Function Load(pNIS As Long) As Integer
On Error GoTo HELL
Dim RS As New ADODB.Recordset

RS.CursorType = adOpenForwardOnly
RS.LockType = adLockReadOnly
SQL = "SELECT NIS,Nama,Tempatlahir,tgllahir,jk,agama,alamat,notelp,namaortu," & _
      " hubungan,agamaortu,alamatortu,pendidikanortu,pekerjaanortu,sd," & _
      " alamatsd,smp,alamatsmp,nonformalsatu,alamatnonformalsatu,nonformaldua," & _
      " alamatnonformaldua,prestasi,prestasidua," & _
      " prestasitiga,prestasiempat,prestasilima,kegiatan,kegiatan1,kegiatan2," & _
      " anakke,jumlahsaudara,saudarakandung,saudaratiri,beratbadan,tinggibandan," & _
      " penyakit1,penyakit2,penyakit3,IDkelas FROM TBSiswa WHERE NIS=" & pNIS & ""
Set RS = DB_CONNECTION.Execute(SQL)
With RS
    If .EOF Then
        Err.Raise BASE_ERROR, SRC_ERROR, "Data tidak ada"
    Else
        NIS = !NIS
        Nama = !Nama
        Tempatlahir = "" & !Tempatlahir
        tgllahir = !tgllahir
        
        jk = "" & !jk
        agama = "" & !agama
        Alamat = "" & !Alamat
        namaortu = "" & !namaortu
        hubungan = "" & !hubungan
        agamaortu = "" & !agamaortu
        alamatortu = "" & !alamatortu
        pendidikanortu = "" & !pendidikanortu
        pekerjaanortu = "" & !pekerjaanortu
        sd = "" & !sd
        alamatsd = "" & !alamatsd
        smp = "" & !smp
        Notelp = "" & !Notelp
        alamatsmp = "" & !alamatsmp
        nonformalsatu = "" & !nonformalsatu
        alamatnonformalsatu = "" & !alamatnonformalsatu
        nonformaldua = "" & !nonformaldua
        alamatnonformaldua = "" & !alamatnonformaldua
        prestasi = "" & !prestasi
        prestasidua = "" & !prestasidua
        prestasitiga = "" & !prestasitiga
        prestasiempat = "" & !prestasiempat
        prestasilima = "" & !prestasilima
        kegiatan = "" & !kegiatan
        kegiatan1 = "" & !kegiatan1
        kegiatan2 = "" & !kegiatan2
        anakke = "" & !anakke
        jumlahsaudara = "" & !jumlahsaudara
        saudarakandung = "" & !saudarakandung
        saudaratiri = "" & !saudaratiri
        beratbadan = "" & !beratbadan
        tinggibandan = "" & !tinggibandan
        penyakit1 = "" & !penyakit1
        penyakit2 = "" & !penyakit2
        penyakit3 = "" & !penyakit3
        IDkelas = 0 & !IDkelas
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

Public Function List(Optional pIsReadOnly As Boolean = True, Optional pWHERE As String = "") As ADODB.Recordset
On Error GoTo HELL
Dim RS As New ADODB.Recordset

SQL = "SELECT NIS,Nama,Tempatlahir,tgllahir,jk,agama,alamat,Notelp,namaortu,hubungan,agamaortu,alamatortu,pendidikanortu,pekerjaanortu,sd,alamatsd,smp,alamatsmp,nonformalsatu,alamatnonformalsatu,nonformaldua,alamatnonformaldua,nonformaltiga,alamatnonformaltiga,prestasi,prestasidua,prestasitiga,prestasiempat,prestasilima,kegiatan,kegiatan1,kegiatan2,anakke,jumlahsaudara,saudarakandung,saudaratiri,beratbadan,tinggibandan,penyakit1,penyakit2,penyakit3 FROM TBSiswa" & pWHERE
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

Public Function Add(NIS As Long, Nama As String, _
Tempatlahir As String, tgllahir As Date, jk As String, agama As String, Alamat As String, Notelp As String, _
namaortu As String, hubungan As String, agamaortu As String, alamatortu As String, pendidikanortu As String, pekerjaanortu As String, _
sd As String, alamatsd As String, smp As String, alamatsmp As String, _
nonformalsatu As String, alamatnonformalsatu As String, nonformaldua As String, _
alamatnonformaldua As String, _
prestasi As String, prestasidua As String, prestasitiga As String, prestasiempat As String, _
prestasilima As String, kegiatan As String, kegiatan1 As String, kegiatan2 As String, anakke As Integer, _
jumlahsaudara As Integer, saudarakandung As Integer, saudaratiri As Integer, beratbadan As Integer, tinggibandan As Integer, penyakit1 As String, _
penyakit2 As String, penyakit3 As String, IDkelas As Integer) As Integer

    On Error GoTo Salah
    SQL = "INSERT INTO TBSiswa (NIS,Nama,Tempatlahir,tgllahir,jk,agama,"
    SQL = SQL & " alamat,Notelp,namaortu,hubungan,agamaortu,alamatortu,pendidikanortu,"
    SQL = SQL & " pekerjaanortu,sd,alamatsd,smp,alamatsmp,nonformalsatu,"
    SQL = SQL & " alamatnonformalsatu,nonformaldua,alamatnonformaldua,"
    SQL = SQL & " prestasi,prestasidua,"
    SQL = SQL & " prestasitiga,prestasiempat,prestasilima,kegiatan,kegiatan1,"
    SQL = SQL & " kegiatan2,anakke,jumlahsaudara,saudarakandung,saudaratiri,"
    SQL = SQL & " beratbadan,tinggibandan,penyakit1,penyakit2,penyakit3,IDKELAS) "
    SQL = SQL & " VALUES (" & NIS & ",'" & Nama & "','" & Tempatlahir & "',"
    SQL = SQL & " #" & Format(tgllahir, "MM/DD/YYYY") & "#,'" & jk & "','" & agama & "',"
    SQL = SQL & " '" & Alamat & "','" & Notelp & "','" & namaortu & "','" & hubungan & "',"
    SQL = SQL & " '" & agamaortu & "','" & alamatortu & "','" & pendidikanortu & "',"
    SQL = SQL & " '" & pekerjaanortu & "','" & sd & "','" & alamatsd & "','" & smp & "',"
    SQL = SQL & " '" & alamatsmp & "','" & nonformalsatu & "',"
    SQL = SQL & " '" & alamatnonformalsatu & "','" & nonformaldua & "',"
    SQL = SQL & " '" & alamatnonformaldua & "','" & prestasi & "',"
    SQL = SQL & " '" & prestasidua & "','" & prestasitiga & "',"
    SQL = SQL & " '" & prestasiempat & "','" & prestasilima & "',"
    SQL = SQL & " '" & kegiatan & "','" & kegiatan1 & "','" & kegiatan2 & "',"
    SQL = SQL & " " & anakke & "," & jumlahsaudara & "," & saudarakandung & ","
    SQL = SQL & " " & saudaratiri & "," & beratbadan & "," & tinggibandan & ","
    SQL = SQL & " '" & penyakit1 & "','" & penyakit2 & "','" & penyakit3 & "'," & IDkelas & ")"
    
    DB_CONNECTION.Execute SQL
    

    
    Add = 1
    Exit Function
Salah:
    Add = 0
End Function

Public Function Edit(NIS As Long, Nama As String, _
Tempatlahir As String, tgllahir As Date, jk As String, agama As String, Alamat As String, Notelp As String, _
namaortu As String, hubungan As String, agamaortu As String, alamatortu As String, pendidikanortu As String, pekerjaanortu As String, _
sd As String, alamatsd As String, smp As String, alamatsmp As String, _
nonformalsatu As String, alamatnonformalsatu As String, nonformaldua As String, _
alamatnonformaldua As String, _
prestasi As String, prestasidua As String, prestasitiga As String, prestasiempat As String, _
prestasilima As String, kegiatan As String, kegiatan1 As String, kegiatan2 As String, anakke As Integer, _
jumlahsaudara As Integer, saudarakandung As Integer, saudaratiri As Integer, beratbadan As Integer, tinggibandan As Integer, penyakit1 As String, _
penyakit2 As String, penyakit3 As String, IDkelas As Integer) As Integer

On Error GoTo Salah
    SQL = "UPDATE TBSiswa SET NIS=" & NIS & ",Nama='" & Nama & "'," & _
        " Tempatlahir='" & Tempatlahir & "'," & _
        " tgllahir=#" & Format(tgllahir, "MM/DD/YYYY") & "#,jk='" & jk & "'," & _
        " agama='" & agama & "',alamat='" & Alamat & "',Notelp = '" & Notelp & "'," & _
        " namaortu='" & namaortu & "',hubungan='" & hubungan & "'," & _
        " agamaortu='" & agamaortu & "',alamatortu='" & alamatortu & "'," & _
        " pendidikanortu='" & pendidikanortu & "',pekerjaanortu='" & pekerjaanortu & "'," & _
        " sd='" & sd & "',alamatsd='" & alamatsd & "',smp='" & smp & "'," & _
        " alamatsmp='" & alamatsmp & "',nonformalsatu='" & nonformalsatu & "'," & _
        " alamatnonformalsatu='" & alamatnonformalsatu & "'," & _
        " nonformaldua='" & nonformaldua & "'," & _
        " alamatnonformaldua='" & alamatnonformaldua & "'," & _
        " nonformaltiga='" & nonformaltiga & "'," & _
        " alamatnonformaltiga='" & alamatnonformaltiga & "'," & _
        " prestasi='" & prestasi & "',prestasidua='" & prestasidua & "'," & _
        " prestasitiga='" & prestasitiga & "',prestasiempat='" & prestasiempat & "'," & _
        " prestasilima='" & prestasilima & "',kegiatan='" & kegiatan & "'," & _
        " kegiatan1='" & kegiatan1 & "',kegiatan2='" & kegiatan2 & "'," & _
        " anakke = " & anakke & ",jumlahsaudara=" & jumlahsaudara & "," & _
        " saudarakandung=" & saudarakandung & ",saudaratiri=" & saudaratiri & "," & _
        " beratbadan=" & beratbadan & ",tinggibandan=" & tinggibandan & "," & _
        " penyakit1='" & penyakit1 & "',penyakit2='" & penyakit2 & "'," & _
        " penyakit3='" & penyakit3 & "', Idkelas=" & IDkelas & " WHERE NIS=" & NIS & ""
        
    DB_CONNECTION.Execute SQL
    
    NIS = NIS
    Nama = Nama
    Tempatlahir = Tempatlahir
    tgllahir = tgllahir
    jk = jk
    agama = agama
    Alamat = Alamat
    namaortu = namaortu
    hubungan = hubungan
    agamaortu = agamaortu
    alamatortu = alamatortu
    pendidikanortu = pendidikanortu
    pekerjaanortu = pekerjaanortu
    sd = sd
    alamatsd = alamatsd
    smp = smp
    Notelp = Notelp
    alamatsmp = alamatsmp
    nonformalsatu = nonformalsatu
    alamatnonformalsatu = alamatnonformalsatu
    nonformaldua = nonformaldua
    alamatnonformaldua = alamatnonformaldua
    nonformaltiga = nonformaltiga
    alamatnonformaltiga = alamatnonformaltiga
    prestasi = prestasi
    prestasidua = prestasidua
    prestasitiga = prestasitiga
    prestasiempat = prestasiempat
    prestasilima = prestasilima
    kegiatan = kegiatan
    kegiatan1 = kegiatan1
    kegiatan2 = kegiatan2
    anakke = anakke
    jumlahsaudara = jumlahsaudara
    saudarakandung = saudarakandung
    saudaratiri = saudaratiri
    beratbadan = beratbadan
    tinggibandan = tinggibandan
    penyakit1 = penyakit1
    penyakit2 = penyakit2
    penyakit3 = penyakit3
    IDkelas = IDkelas
    Edit = 1
    Exit Function
Salah:
    Edit = 0
End Function

