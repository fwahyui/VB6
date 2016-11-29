Attribute VB_Name = "modFunction"
Option Explicit

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'Function used to format recordset
Public Function FormatRS(ByVal srcField As Field, Optional AllowNewLine As Boolean) As String
    Dim strRet As String
    
    With srcField
        If AllowNewLine = True Then
            strRet = srcField
        Else
            strRet = Replace(srcField, vbCrLf, " ", , , vbTextCompare)
        End If
        
        'If srcField.Type = adCurrency Or srcField.Type = adDouble Then
        If srcField.Type = adCurrency Then
            strRet = Format$(srcField, "#,##0.00")
        ElseIf srcField.Type = adDate Then
            strRet = Format$(srcField, "MMM-dd-yyyy")
        Else
            strRet = srcField
        End If
    End With
    
    FormatRS = strRet
    
    strRet = vbNullString
End Function

'Function that will format return a generated id
Public Function GenerateID(ByVal srcNo As String, ByVal src1stStr As String, ByVal src2ndStr As String) As String
    If Len(src2ndStr) <= Len(srcNo) Then
        GenerateID = src1stStr & srcNo
    Else
        GenerateID = src1stStr & Left$(src2ndStr, Len(src2ndStr) - Len(srcNo)) & srcNo
    End If
End Function

'Function used to check if the record exit or not.
Public Function isRecordExist(ByVal sTable As String, ByVal sField As String, ByVal sStr As String, Optional isString As Boolean) As Boolean
    Dim RS As New Recordset

    RS.CursorLocation = adUseClient
    If isString = False Then
        RS.Open "Select * From " & sTable & " Where " & sField & " = " & sStr, CN, adOpenStatic, adLockOptimistic
    Else
        RS.Open "Select * From " & sTable & " Where " & sField & " = '" & sStr & "'", CN, adOpenStatic, adLockOptimistic
    End If
    If RS.RecordCount < 1 Then
        isRecordExist = False
    Else
        isRecordExist = True
    End If
    Set RS = Nothing
End Function

'Function used to check if the Ascii is a number or not (return 0 if number)
Public Function isNumber(ByVal sKeyAscii) As Integer
    If Not ((sKeyAscii >= 48 And sKeyAscii <= 57) Or sKeyAscii = 8 Or sKeyAscii = 46) Then
        isNumber = 0
    Else
        isNumber = sKeyAscii
    End If
End Function

'Function used to check if the record exist in Flex grid
Public Function isRecExistInFlex(ByVal srcFlexGrd As MSHFlexGrid, ByVal srcWhatCol As Integer, ByVal srcFindWhat As String) As Boolean
    isRecExistInFlex = False
    Dim i As Long
    For i = 1 To srcFlexGrd.Rows - 1
        If srcFlexGrd.TextMatrix(i, srcWhatCol) = srcFindWhat Then isRecExistInFlex = True: Exit For
    Next i
    i = 0
End Function

'Function used to check if the record exist in Flex grid
Public Function getFlexPos(ByVal srcFlexGrd As MSHFlexGrid, ByVal srcWhatCol As Integer, ByVal srcFindWhat As String) As Integer
    Dim R As Long, ret As Integer
    
    ret = -1 'Means not found
    For R = 0 To srcFlexGrd.Rows - 1
        If srcFlexGrd.TextMatrix(R, srcWhatCol) = srcFindWhat Then ret = R: Exit For
    Next R
    
    getFlexPos = ret
    R = 0: ret = 0
End Function

'Function used to left split user fields
Public Function LeftSplitUF(ByVal srcUF As String) As String
    If srcUF = "*~~~~~*" Then LeftSplitUF = "": Exit Function
    Dim i As Integer
    Dim t As String
    For i = 1 To Len(srcUF)
        If Mid$(srcUF, i, 7) = "*~~~~~*" Then
            Exit For
        Else
            t = t & Mid$(srcUF, i, 1)
        End If
    Next i
    LeftSplitUF = t
    i = 0
    t = ""
End Function

'Function used to right split user fields
Public Function RightSplitUF(ByVal srcUF As String) As String
    If srcUF = "*~~~~~*" Then RightSplitUF = "": Exit Function
    Dim i As Integer
    Dim t As String
    For i = (InStr(1, srcUF, "*~~~~~*", vbTextCompare) + 7) To Len(srcUF)
        t = t & Mid$(srcUF, i, 1)
    Next i
    RightSplitUF = t
    i = 0
    t = ""
End Function

'Function that return true if the control is empty
Public Function is_empty(ByRef sText As Variant, Optional UseTagValue As Boolean) As Boolean
    On Error Resume Next
    If sText.Text = "" Then
        is_empty = True
        If UseTagValue = True Then
            MsgBox "The field '" & sText.Tag & "' is required.Please check it!", vbExclamation
        Else
            MsgBox "The field is required.Please check it!", vbExclamation
        End If
        sText.SetFocus
    Else
        is_empty = False
    End If
End Function

'Function used to change the yes/no value
Public Function changeYNValue(ByVal srcStr As String) As String
    Select Case srcStr
        Case "Y": changeYNValue = "1"
        Case "N": changeYNValue = "0"
        Case "1": changeYNValue = "Y"
        Case "0": changeYNValue = "N"
    End Select
End Function

'Function used to change the true/false value
Public Function changeTFValue(ByVal srcStr As String) As String
    Select Case srcStr
        Case "True": changeTFValue = "1"
        Case "False": changeTFValue = "0"
        Case "1": changeTFValue = "True"
        Case "0": changeTFValue = "False"
    End Select
End Function

'Function that return true if the control is numeric
Public Function is_numeric(ByRef sText As String) As Boolean
    If IsNumeric(sText) = False Then
        is_numeric = False
        MsgBox "The field required a numeric input.Please check it!", vbExclamation
    Else
        is_numeric = True
    End If
End Function

'Function that return the value of a certain field
Public Function getValueAt(ByVal srcSQL As String, ByVal whichField As String) As String
    Dim RS As New Recordset
    
    RS.CursorLocation = adUseClient
    RS.Open srcSQL, CN, adOpenStatic, adLockReadOnly
    If RS.RecordCount > 0 Then getValueAt = RS.Fields(whichField)
    
    Set RS = Nothing
End Function

'Convert string to number
'I create this istead of val() co'z val return incorrect value
'ex. Try to see the output of val("3,800")
'It did not support characters like , and etc.
Public Function toNumber(ByVal srcCurrency As String, Optional RetZeroIfNegative As Boolean) As Double
    If srcCurrency = "" Then
        toNumber = 0
    Else
        Dim retValue As Double
        If InStr(1, srcCurrency, ",") > 0 Then
            retValue = Val(Replace(srcCurrency, ",", "", , , vbTextCompare))
        Else
            retValue = Val(srcCurrency)
        End If
        If RetZeroIfNegative = True Then
            If retValue < 1 Then retValue = 0
        End If
        toNumber = retValue
        retValue = 0
    End If
End Function

'Function that return the count of the rows in the table
Public Function getRecordCount(ByVal srcTable As String, Optional srcCondition As String, Optional isFormatted As Boolean) As String
    If srcCondition <> "" Then srcCondition = " " & srcCondition
    Dim RS As New Recordset
    
    RS.CursorLocation = adUseClient
    RS.Open "SELECT COUNT(PK) as TCount FROM " & srcTable & srcCondition, CN, adOpenStatic, adLockReadOnly
    If isFormatted = True Then
        getRecordCount = Format$(RS![TCount], "#,##0")
    Else
        getRecordCount = RS![TCount]
    End If
    Set RS = Nothing
End Function

'Function that will return a currenct format
Public Function toMoney(ByVal srcCurr As String) As String
   toMoney = Format$(IIf(Trim(srcCurr) = "", 0, srcCurr), "#,##0.00")
End Function

'Function used to determine if the object has been set
Public Function isObjectSet(srcObject As Object) As Boolean
    On Error GoTo err
    'I use tag because almost all controls have this
    srcObject.Tag = srcObject.Tag
    isObjectSet = True
    
    Exit Function
err:
    isObjectSet = False
End Function

'Function used to get the end day number of a cetain month
Public Function getEndDay(ByVal srcDate As Date) As Byte
    Dim h1 As String
    h1 = Format(srcDate, "mm")
    On Error GoTo err
    Select Case h1
        Case Is = "01": getEndDay = 31
        Case Is = "02": getEndDay = Day(h1 & "/29/" & Format(srcDate, "yy"))
        Case Is = "03": getEndDay = 31
        Case Is = "04": getEndDay = 30
        Case Is = "05": getEndDay = 31
        Case Is = "06": getEndDay = 30
        Case Is = "07": getEndDay = 31
        Case Is = "08": getEndDay = 31
        Case Is = "09": getEndDay = 30
        Case Is = "10": getEndDay = 31
        Case Is = "11": getEndDay = 30
        Case Is = "12": getEndDay = 31
    End Select
    h1 = ""
    Exit Function
err:
        If err.Number = 13 Then getEndDay = 28: h1 = "" 'Day if encounter not a left-year
End Function

Public Function getUnitID(ByVal sUnit As String) As Long
  Dim RS As New ADODB.Recordset
  Dim sql As String
  
  sql = "SELECT UnitID From Unit WHERE (((Unit)='" & Replace(sUnit, "'", "''") & "'))"
  RS.Open sql, CN, adOpenDynamic, adLockOptimistic
  
  If Not RS.EOF Then
    getUnitID = RS!UnitID
  Else
    getUnitID = 0
  End If
   
  RS.Close
  Set RS = Nothing
End Function

Function GetINI(strMain As String, strSub As String) As String
    Dim strBuffer As String
    Dim lngLen As Long
    Dim lngRet As Long
    
    strBuffer = Space(100)
    lngLen = Len(strBuffer)
    lngRet = GetPrivateProfileString(strMain, strSub, vbNullString, strBuffer, lngLen, App.Path & "\config.txt")
    GetINI = Left(strBuffer, lngRet)
End Function

'Function to determine user's permission
Public Function allowOpen(frmForm As String, ByRef User As String) As Boolean
    Dim RS As New Recordset
    Dim srcSQL As String
    
    srcSQL = "SELECT * FROM qry_User WHERE Form = '" & frmForm & "' AND UserID = '" & User & "'"
    
    RS.CursorLocation = adUseClient
    RS.Open srcSQL, CN, adOpenStatic, adLockReadOnly
    If RS.RecordCount > 0 Then allowOpen = True
    
    Set RS = Nothing
End Function

Public Function ChangePermission(ByVal UserPermID As Long, ByVal bNewPermission As Boolean) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    ChangePermission = False
    
    sSQL = "SELECT *" & _
            " From [User Permission]" & _
            " WHERE UserPermissionID=" & UserPermID
    
    vRS.Open sSQL, CN, adOpenStatic, adLockOptimistic
    
    On Error GoTo RAE
    
    vRS.MoveFirst
    vRS.Fields("AllowOpen").Value = bNewPermission
    vRS.Update
    
    ChangePermission = True
    
RAE:
    Set vRS = Nothing
End Function

