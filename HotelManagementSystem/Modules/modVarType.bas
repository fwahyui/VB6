Attribute VB_Name = "modVarType"

'Variable structure for user
Public Type USER_INFO
    USER_PK As Long
    USER_NAME As String
    USER_ISADMIN As Boolean
End Type

'Enumerator for form state
Public Enum FormState
    adStateAddMode = 0
    adStateEditMode = 1
    adStatePopupMode = 2
    adStateViewMode = 3
End Enum

Public Type BUSINESS_INFO
    BUSINESS_NAME As String
    BUSINESS_ADDRESS As String
    BUSINESS_CONTACT_INFO As String
End Type

