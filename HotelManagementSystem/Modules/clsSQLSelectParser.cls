VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSQLSelectParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Variables for SQL statement
Private m_Fields              As String
Private m_Tables              As String
Private m_wCondition          As String
Private m_GroupOrder          As String
Private m_hCondition          As String
Private m_SortOrder           As String

'Variables for storing saved statement
Private def_Fields            As String
Private def_Tables            As String
Private def_wCondition        As String
Private def_GroupOrder        As String
Private def_hCondition        As String
Private def_SortOrder         As String

'Return the fields
Public Property Get Fields() As String
    Fields = m_Fields
End Property

'Set the fields
Public Property Let Fields(ByVal srcFields As String)
    m_Fields = srcFields
End Property

'Return the tables
Public Property Get Tables() As String
    Tables = m_Tables
End Property

'Set the tables
Public Property Let Tables(ByVal srcTables As String)
    m_Tables = srcTables
End Property

'Return the where condition
Public Property Get wCondition() As String
    wCondition = m_wCondition
End Property

'Set the where condition
Public Property Let wCondition(ByVal srcwCondition As String)
    m_wCondition = srcwCondition
End Property

'Return the group order
Public Property Get GroupOrder() As String
    GroupOrder = m_GroupOrder
End Property

'Set the group order
Public Property Let GroupOrder(ByVal srcGroupOrder As String)
    m_GroupOrder = srcGroupOrder
End Property

'Return the having condition
Public Property Get hCondition() As String
    hCondition = m_hCondition
End Property

'Set the having condition
Public Property Let hCondition(ByVal srchCondition As String)
    m_hCondition = srchCondition
End Property

'Return the sort order
Public Property Get SortOrder() As String
    SortOrder = m_SortOrder
End Property

'Set the sort order
Public Property Let SortOrder(ByVal srcSortOrder As String)
    m_SortOrder = srcSortOrder
End Property

'Return the SQL statement
Public Property Get SQLStatement() As String
    'Required statement
    SQLStatement = "SELECT " & m_Fields & " FROM " & m_Tables
    'Optional statement
    If m_wCondition <> "" Then SQLStatement = SQLStatement & " WHERE " & m_wCondition
    If m_GroupOrder <> "" Then SQLStatement = SQLStatement & " GROUP BY " & m_GroupOrder
    If m_hCondition <> "" Then SQLStatement = SQLStatement & " HAVING " & m_hCondition
    If m_SortOrder <> "" Then SQLStatement = SQLStatement & " ORDER BY " & m_SortOrder
    Debug.Print SQLStatement
End Property

'Save the statement
Public Sub SaveStatement()
    def_Fields = m_Fields
    def_Tables = m_Tables
    def_wCondition = m_wCondition
    def_GroupOrder = m_GroupOrder
    def_hCondition = m_hCondition
    def_SortOrder = m_SortOrder
End Sub

'Restore the statement
Public Sub RestoreStatement()
    m_Fields = def_Fields
    m_Tables = def_Tables
    m_wCondition = def_wCondition
    m_GroupOrder = def_GroupOrder
    m_hCondition = def_hCondition
    m_SortOrder = def_SortOrder
End Sub

Private Sub Class_Terminate()
    'Clear used variables
    m_Fields = vbNullString
    m_Tables = vbNullString
    m_wCondition = vbNullString
    m_GroupOrder = vbNullString
    m_hCondition = vbNullString
    m_SortOrder = vbNullString
    
    def_Fields = vbNullString
    def_Tables = vbNullString
    def_wCondition = vbNullString
    def_GroupOrder = vbNullString
    def_hCondition = vbNullString
    def_SortOrder = vbNullString
End Sub


