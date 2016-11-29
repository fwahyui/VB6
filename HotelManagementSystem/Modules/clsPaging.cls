VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPaging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_PAGE_CURRENT              As Long 'Store current page
Private m_PAGE_PREV                 As Long 'Store the previous page
Private m_PAGE_NEXT                 As Long 'Store the next page
Private m_PAGE_TOTAL                As Long 'Store the total page count

Private m_Recordset                 As Recordset 'Reference recordset where we will perform the paging

Private m_PageBy                    As Long  'Store the how many record display per page
Private m_PageStart                 As Long 'Store the starting record position
Private m_PageEnd                   As Long 'Store the ending record position

Private PageInformation             As String 'Store the page information

'Return the current position
Public Property Get CurrentPosition() As Long
    MsgBox m_PAGE_CURRENT
    CurrentPosition = m_PAGE_CURRENT
End Property
'Set the current position
Public Property Let CurrentPosition(ByVal whichPage As Long)
    m_PAGE_CURRENT = whichPage
    PageRecordset
End Property


'Procedure used to refresh pages
Public Sub Refresh()
    PageRecordset
End Sub

'Start the class and get all the input needed
Public Sub Start(ByRef srcRecordset, ByVal srcPageBy As Long)
Set m_Recordset = srcRecordset
    m_Recordset.CacheSize = srcPageBy
    m_PageBy = srcPageBy
End Sub

Private Sub Class_Terminate()
    'Clear all variables
    m_PAGE_CURRENT = 0
    m_PAGE_NEXT = 0
    m_PAGE_PREV = 0
    m_PAGE_TOTAL = 0
    
    m_PageBy = 0
    m_PageStart = 0
    m_PageEnd = 0
    
    PageInformation = vbNullString
    
    Set m_Recordset = Nothing
End Sub

'Procedure for paging the recordset
Private Sub PageRecordset()
    If m_Recordset.RecordCount < 1 Then
        m_PAGE_NEXT = 0
        m_PAGE_PREV = 0
        m_PAGE_TOTAL = 1
        
        m_PageStart = 0
        m_PageEnd = 0
        
        PageInformation = "0 - 0 of 0"
        
        Exit Sub
    End If
    'Initialize the paging variables
    m_PAGE_TOTAL = Fix(m_Recordset.RecordCount / m_PageBy)
    If InStr(1, (m_Recordset.RecordCount / m_PageBy), ".") > 0 Then
        m_PAGE_TOTAL = m_PAGE_TOTAL + 1
    End If
    
    If m_PAGE_TOTAL = 0 Then m_PAGE_TOTAL = m_PAGE_TOTAL + 1
    
    'Page the records
    m_PAGE_NEXT = m_PAGE_CURRENT + 1
    m_PAGE_PREV = m_PAGE_CURRENT - 1
    
    If m_PAGE_TOTAL = 1 Then
        m_PageStart = 1
        m_PageEnd = m_Recordset.RecordCount
    ElseIf m_PAGE_CURRENT = 1 And m_PAGE_TOTAL > 1 Then
        m_PageStart = 1
        m_PageEnd = m_PageBy
    ElseIf m_PAGE_CURRENT = m_PAGE_TOTAL And m_PAGE_CURRENT > 1 Then
        m_PageStart = ((m_PAGE_CURRENT - 1) * m_PageBy) + 1
        m_PageEnd = m_Recordset.RecordCount
    Else
        m_PageStart = ((m_PAGE_CURRENT - 1) * m_PageBy) + 1
        m_PageEnd = (m_PAGE_NEXT - 1) * m_PageBy
    End If
    'Set the page information
    PageInformation = m_PageStart & " - " & m_PageEnd & " of " & m_Recordset.RecordCount
End Sub

'Return the current page
Public Function PAGE_CURRENT() As Long
    PAGE_CURRENT = m_PAGE_CURRENT
End Function
'Return the previous page
Public Function PAGE_PREVIOUS() As Long
    PAGE_PREVIOUS = m_PAGE_PREV
End Function
'Return the next page
Public Function PAGE_NEXT() As Long
    PAGE_NEXT = m_PAGE_NEXT
End Function
'Return the page total
Public Function PAGE_TOTAL() As Long
    PAGE_TOTAL = m_PAGE_TOTAL
End Function

'Return the starting record position
Public Function PageStart() As Long
    PageStart = m_PageStart
End Function
'Return the ending record position
Public Function PageEnd() As Long
    PageEnd = m_PageEnd
End Function
'Return the page information
Public Function PageInfo() As String
    PageInfo = PageInformation
End Function




