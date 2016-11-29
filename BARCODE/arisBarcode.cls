VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "arisBarcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Function Code128(zObj As Object, zBarH As Double, BarText As String, Optional ByVal HasCaption As Boolean = False)
    Bar128 zObj, zBarH, BarText, HasCaption
End Function

Function Code39(zObj As Object, zBarH As Double, BarText As String, Optional zWithCheckSum As Boolean = False, Optional ByVal HasCaption As Boolean = False)
    Bar39 zObj, zBarH, BarText, zWithCheckSum, HasCaption
End Function

Function EAN13(zObj As Object, zBarH As Integer, BarText As String, Optional ByVal HasCaption As Boolean = False)
    BarEAN "EAN13", zObj, zBarH, BarText, HasCaption
End Function
Function EAN8(zObj As Object, zBarH As Integer, BarText As String, Optional ByVal HasCaption As Boolean = False)
    BarEAN "EAN8", zObj, zBarH, BarText, HasCaption
End Function
Function PDF417(zObj As Object, BarText As String)
    GenPDF417 zObj, BarText
End Function

