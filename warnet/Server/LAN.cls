VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LAN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim PCLIST As String 'buffer to hold pc's names
Private Type NETRESOURCE
   dwScope As Long
   dwType As Long
   dwDisplayType As Long
   dwUsage As Long
   lpLocalName As Long
   lpRemoteName As Long
   lpComment As Long
   lpProvider As Long
End Type

Private Declare Function WNetOpenEnum Lib "mpr.dll" Alias _
  "WNetOpenEnumA" (ByVal dwScope As Long, ByVal dwType As Long, _
  ByVal dwUsage As Long, lpNetResource As Any, lphEnum As Long) As Long

Private Declare Function WNetEnumResource Lib "mpr.dll" Alias _
  "WNetEnumResourceA" (ByVal hEnum As Long, lpcCount As Long, _
  ByVal lpBuffer As Long, lpBufferSize As Long) As Long

Private Declare Function WNetCloseEnum Lib "mpr.dll" _
   (ByVal hEnum As Long) As Long

Private Const RESOURCE_CONNECTED = &H1
Private Const RESOURCE_GLOBALNET = &H2
Private Const RESOURCE_REMEMBERED = &H3

Private Const RESOURCETYPE_ANY = &H0
Private Const RESOURCETYPE_DISK = &H1
Private Const RESOURCETYPE_PRINT = &H2
Private Const RESOURCETYPE_UNKNOWN = &HFFFF

Private Const RESOURCEUSAGE_CONNECTABLE = &H1
Private Const RESOURCEUSAGE_CONTAINER = &H2
Private Const RESOURCEUSAGE_RESERVED = &H80000000

Private Const GMEM_FIXED = &H0
Private Const GMEM_ZEROINIT = &H40
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Private Declare Function GlobalAlloc Lib "kernel32" _
  (ByVal wFlags As Long, ByVal dwBytes As Long) As Long

Private Declare Function GlobalFree Lib "kernel32" _
  (ByVal hMem As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias _
  "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, _
   ByVal cbCopy As Long)
   
Private Declare Function CopyPointer2String Lib _
  "kernel32" Alias "lstrcpyA" (ByVal NewString As _
  String, ByVal OldString As Long) As Long

Private Function DoNetEnum()
Dim hEnum As Long, lpBuff As Long, NR As NETRESOURCE
Dim cbBuff As Long, cCount As Long
Dim P As Long, res As Long, i As Long

On Error Resume Next
If Err.Number > 0 Then Exit Function

On Error GoTo ErrorHandler

NR.lpRemoteName = 0

cbBuff = 1024 * 31
cCount = &HFFFFFFFF

res = WNetOpenEnum(RESOURCE_GLOBALNET, _
  RESOURCETYPE_ANY, 0, NR, hEnum)

If res = 0 Then

   lpBuff = GlobalAlloc(GPTR, cbBuff)

   res = WNetEnumResource(hEnum, cCount, lpBuff, cbBuff)
   If res = 0 Then
      P = lpBuff
      For i = 1 To cCount
         CopyMemory NR, ByVal P, LenB(NR)
         'list.AddItem PointerToString(NR.lpRemoteName)
         DoNetEnum2 NR
         P = P + LenB(NR)
      Next i
      End If


ErrorHandler:
On Error Resume Next
   If lpBuff <> 0 Then GlobalFree (lpBuff)
   WNetCloseEnum (hEnum)

End If

End Function

Private Function PointerToString(P As Long) As String

   Dim s As String
   s = String(65535, Chr$(0))
   CopyPointer2String s, P
   PointerToString = Left(s, InStr(s, Chr$(0)) - 1)

End Function

Private Sub DoNetEnum2(NR As NETRESOURCE)

   Dim hEnum As Long, lpBuff As Long
   Dim cbBuff As Long, cCount As Long
   Dim P As Long, res As Long, i As Long

   cbBuff = 1024 * 31
   cCount = &HFFFFFFFF

   res = WNetOpenEnum(RESOURCE_GLOBALNET, _
     RESOURCETYPE_ANY, 0, NR, hEnum)
   If res = 0 Then

      lpBuff = GlobalAlloc(GPTR, cbBuff)
      res = WNetEnumResource(hEnum, cCount, lpBuff, cbBuff)
      If res = 0 Then
         P = lpBuff
         For i = 1 To cCount
            CopyMemory NR, ByVal P, LenB(NR)
            Dim st As String
            Select Case NR.dwDisplayType
                Case &H1
                    st = "Domain"
                Case &H2
                    st = "Server"
                Case &H3
                    st = "Share"
                Case &H4
                    st = "File"
                Case &H5
                    st = "Groups"
                Case &H6
                    st = "Protocol Categories"
            End Select
            If LCase(st) <> "domain" Then
            PCLIST = PCLIST & "||" & Replace(PointerToString(NR.lpRemoteName), "\", "") '& " is  a : " & st
            End If
            DoEvents
            If Not NR.dwDisplayType = 2 Then DoNetEnum2 NR
            P = P + LenB(NR)
         Next i
      End If

      If lpBuff <> 0 Then GlobalFree (lpBuff)
      WNetCloseEnum (hEnum)

   End If
PCLIST = stripDelimiter(PCLIST)
End Sub
Property Get GetPCList() As String
GetPCList = PCLIST
End Property
Private Function stripDelimiter(ByVal s As String) As String
If Left(s, 2) = "||" Then s = Right(s, Len(s) - 2)
stripDelimiter = s
End Function
Private Sub Class_Initialize()
DoNetEnum
End Sub
