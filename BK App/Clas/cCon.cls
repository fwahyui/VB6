VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public DBConnRian As ADODB.Connection


Public Property Get DBConnection() As ADODB.Connection
    On Error GoTo HELL
    
    Set DBConnection = DB_CONNECTION
    Set DBConnRian = DBConnection
    
    Exit Property
HELL:
    Err.Raise Err.Number, Err.Source, Err.Description
End Property


Public Sub OpenConnection(ByVal pConnectionString As String)
    On Error GoTo HELL

    Set DB_CONNECTION = New ADODB.Connection
    With DB_CONNECTION
        .Mode = adModeShareDenyNone
        .ConnectionString = pConnectionString
        .ConnectionTimeout = 30
        .Open
    End With
    
    Exit Sub
HELL:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub BeginTrans()
    DB_CONNECTION.BeginTrans
End Sub

Public Sub CommitTrans()
    DB_CONNECTION.CommitTrans
End Sub

Public Sub RollBackTrans()
    DB_CONNECTION.RollBackTrans
End Sub

Public Sub DisConnect()
    On Error GoTo HELL
    
    DB_CONNECTION.Close
    Set DB_CONNECTION = Nothing
    
    Exit Sub
HELL:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub





