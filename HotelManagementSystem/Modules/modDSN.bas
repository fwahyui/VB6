Attribute VB_Name = "modDSN"
Private Const ODBC_ADD_SYS_DSN = 4        ' Add data source
Private Const ODBC_CONFIG_DSN = 2         ' Configure (edit) data Source
Private Const ODBC_REMOVE_DSN = 3         ' Remove data source
Private Const vbAPINull As Long = vbNull  ' NULL Pointer


'Function Declare
Private Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" _
          (ByVal hwndParent As Long, ByVal fRequest As Long, _
          ByVal lpszDriver As String, ByVal lpszAttributes As String) _
          As Long

Public Sub createDSN()
    'Creating the DSN

    #If Win32 Then
          Dim intRet As Long
    #Else
          Dim intRet As Integer
    #End If

    Dim strDriver As String
    Dim strAttributes As String

    strDriver = "Microsoft Access Driver (*.mdb)"

    strAttributes = strAttributes & "DESCRIPTION=" & "Hotel DSN " & Chr$(0)
    strAttributes = strAttributes & "DSN=" & "Hotel" & Chr$(0)
    strAttributes = strAttributes & "PWD=" & "jaypee" & Chr$(0)
    strAttributes = strAttributes & "UID=" & "admin" & Chr$(0)
    strAttributes = strAttributes & "DBQ=" & DBPath & Chr$(0)

    intRet = SQLConfigDataSource(vbAPINull, ODBC_ADD_SYS_DSN, _
    strDriver, strAttributes)

    ' DSN created
End Sub



