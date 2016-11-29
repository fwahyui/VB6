Attribute VB_Name = "Module3"
'i know this can be cut down
'but to do that you will have to remove errors
'-----------------------------------------------
'Program strings
'datasource is were the information is coming from
'key1 tells it to add or subtract to the asc code
Option Explicit

Public Function encrypted(key1 As String, DataSource As String) As String

Dim letter As String
Dim data As String
Dim x As Integer

For x = 1 To Len(DataSource)
letter = Asc(Mid(DataSource, x, 1))
On Error GoTo errval
data = data & Chr(Val(letter) + Val(key1))
Next x

If data = "" Then
encrypted = DataSource
ElseIf data <> "" Then
encrypted = data
End If
Exit Function
errval:
MsgBox "Error: Val" & letter & " + " & key1, vbCritical, "Error Val" & letter & " + " & key1
encrypted = DataSource
End Function
Public Function decrypted(key1 As String, DataSource As String) As String
Dim letter As String
Dim data As String
Dim x As Integer

For x = 1 To Len(DataSource)
letter = Asc(Mid(DataSource, x, 1))
On Error GoTo errval
data = data & Chr(Val(letter) - Val(key1))
Next x

If data = "" Then
MsgBox "Error: Val" & letter & " + " & key1, vbCritical, "Error Val" & letter & " + " & key1
decrypted = DataSource
ElseIf data <> "" Then
decrypted = data
End If
Exit Function
errval:
MsgBox "Error: Val" & letter & " - " & key1, vbCritical, "Error Val" & letter & " - " & key1
decrypted = DataSource
End Function
