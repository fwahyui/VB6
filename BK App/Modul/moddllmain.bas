Attribute VB_Name = "moddllmain"
Option Explicit
Public DB_CONNECTION As ADODB.Connection

Public Function SQLFix(SQL As String) As String
    On Error GoTo ErrorHandler
    Dim ReturnString As String
    Dim EndOfString As Boolean
    Dim n As Long
    Dim LengthOfString As Long
    Dim StringPosition As Long

    EndOfString = False
    n = 1
    ReturnString = SQL
    LengthOfString = Len(SQL)

    Do While Not EndOfString
        StringPosition = InStr(n, ReturnString, "'")

        If (StringPosition <> 0) Then
            ReturnString = Left$(ReturnString, StringPosition) & _
                    Mid$(ReturnString, StringPosition, LengthOfString)
        Else
            EndOfString = True
        End If

        n = StringPosition + 2
    Loop

    SQLFix = ReturnString

    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function koneksi() As String
    koneksi = "Provider = Microsoft.Jet.OLEDB.4.0; Data Source = " & App.Path & "\db\dbbk.mdb" & ";Persist Security Info=False;Jet OLEDB:Database Password=" & "jendeladunia"
End Function







