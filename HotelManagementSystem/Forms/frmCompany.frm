VERSION 5.00
Object = "{9EDDC69F-10E8-4DE7-9648-EC8A45F005C0}#1.0#0"; "b8Controls4.ocx"
Begin VB.Form frmCompany 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Company"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6405
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   6405
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   0
      Left            =   1680
      MaxLength       =   100
      TabIndex        =   0
      Tag             =   "Company"
      Top             =   180
      Width           =   2415
   End
   Begin VB.CommandButton cmdUsrHistory 
      Caption         =   "Modification History"
      Height          =   315
      Left            =   765
      TabIndex        =   10
      Top             =   3690
      Width           =   1680
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4605
      TabIndex        =   9
      Top             =   3675
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   3165
      TabIndex        =   8
      Top             =   3675
      Width           =   1335
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   2
      Left            =   1680
      MaxLength       =   100
      TabIndex        =   2
      Top             =   975
      Width           =   2415
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   1
      Left            =   1680
      MaxLength       =   100
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   3
      Left            =   1680
      MaxLength       =   100
      TabIndex        =   3
      Top             =   1365
      Width           =   3840
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   5
      Left            =   1680
      MaxLength       =   100
      TabIndex        =   5
      Top             =   2115
      Width           =   2415
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   4
      Left            =   1680
      MaxLength       =   100
      TabIndex        =   4
      Top             =   1740
      Width           =   2415
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   7
      Left            =   1680
      MaxLength       =   100
      TabIndex        =   7
      Top             =   2835
      Width           =   2415
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   6
      Left            =   1680
      MaxLength       =   100
      TabIndex        =   6
      Top             =   2460
      Width           =   2415
   End
   Begin b8Controls4.b8Line b8Line1 
      Height          =   30
      Left            =   495
      TabIndex        =   11
      Top             =   3450
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   53
      BorderColor1    =   11325655
      BorderColor2    =   16185592
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Company"
      Height          =   240
      Index           =   7
      Left            =   480
      TabIndex        =   19
      Top             =   180
      Width           =   1140
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Lastname"
      Height          =   240
      Index           =   1
      Left            =   480
      TabIndex        =   18
      Top             =   600
      Width           =   1140
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Firstname"
      Height          =   240
      Index           =   0
      Left            =   480
      TabIndex        =   17
      Top             =   975
      Width           =   1140
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Address"
      Height          =   240
      Index           =   2
      Left            =   480
      TabIndex        =   16
      Top             =   1365
      Width           =   1140
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Email Address"
      Height          =   240
      Index           =   3
      Left            =   480
      TabIndex        =   15
      Top             =   1740
      Width           =   1140
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Tel No"
      Height          =   240
      Index           =   4
      Left            =   480
      TabIndex        =   14
      Top             =   2115
      Width           =   1140
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Mobile No"
      Height          =   240
      Index           =   5
      Left            =   480
      TabIndex        =   13
      Top             =   2460
      Width           =   1140
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Fax No"
      Height          =   240
      Index           =   6
      Left            =   480
      TabIndex        =   12
      Top             =   2835
      Width           =   1140
   End
End
Attribute VB_Name = "frmCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public State                As FormState 'Variable used to determine on how the form used
Public PK                   As Long 'Variable used to get what record is going to edit
Public srcText              As TextBox 'Used in pop-up mode

Dim HaveAction              As Boolean 'Variable used to detect if the user perform some action
Dim RS                      As New Recordset

Private Sub DisplayForEditing()
    On Error GoTo err
    
    With RS
        txtEntry(0).Text = .Fields("Company")
        txtEntry(1).Text = .Fields("LastName")
        txtEntry(2).Text = .Fields("FirstName")
        txtEntry(3).Text = .Fields("Address")
        txtEntry(4).Text = .Fields("EmailAddress")
        txtEntry(5).Text = .Fields("TelNo")
        txtEntry(6).Text = .Fields("MobileNo")
        txtEntry(7).Text = .Fields("FaxNo")
    End With
    
    Exit Sub
err:
        If err.Number = 94 Then Resume Next
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub ResetFields()
    clearText Me
    
    txtEntry(0).SetFocus
End Sub

Private Sub cmdSave_Click()
    If is_empty(txtEntry(0), True) = True Then Exit Sub
    
    If State = adStateAddMode Then
        RS.AddNew
        RS.Fields("CompanyID") = PK
        RS.Fields("DateAdded") = Now
        RS.Fields("AddedByFK") = CurrUser.USER_PK
    Else
        RS.Fields("DateModified") = Now
        RS.Fields("LastUserFK") = CurrUser.USER_PK
    End If
    'Phill 2:12
    With RS
        .Fields("Company") = txtEntry(0).Text
        .Fields("LastName") = txtEntry(1).Text
        .Fields("FirstName") = txtEntry(2).Text
        .Fields("Address") = txtEntry(3).Text
        .Fields("EmailAddress") = txtEntry(4).Text
        .Fields("TelNo") = txtEntry(5).Text
        .Fields("MobileNo") = txtEntry(6).Text
        .Fields("FaxNo") = txtEntry(7).Text
        
        .Update
    End With
    
    HaveAction = True
    
    If State = adStateAddMode Then
        MsgBox "New record has been successfully saved.", vbInformation
        If MsgBox("Do you want to add another new record?", vbQuestion + vbYesNo) = vbYes Then
            ResetFields
            PK = getIndex("Company")
         Else
            Unload Me
        End If
    Else
        MsgBox "Changes in  record has been successfully saved.", vbInformation
        Unload Me
    End If
End Sub

Private Sub cmdUsrHistory_Click()
    On Error Resume Next
    Dim tDate1 As String
    Dim tDate2 As String
    Dim tUser1 As String
    Dim tUser2 As String
    
    tDate1 = Format$(RS.Fields("DateAdded"), "MMM-dd-yyyy HH:MM AMPM")
    tDate2 = Format$(RS.Fields("DateModified"), "MMM-dd-yyyy HH:MM AMPM")
    
    tUser1 = getValueAt("SELECT PK,CompleteName FROM tbl_SM_Users WHERE PK = " & RS.Fields("AddedByFK"), "CompleteName")
    tUser2 = getValueAt("SELECT PK,CompleteName FROM tbl_SM_Users WHERE PK = " & RS.Fields("LastUserFK"), "CompleteName")
    
    MsgBox "Date Added: " & tDate1 & vbCrLf & _
           "Added By: " & tUser1 & vbCrLf & _
           "" & vbCrLf & _
           "Last Modified: " & tDate2 & vbCrLf & _
           "Modified By: " & tUser2, vbInformation, "Modification History"
           
    tDate1 = vbNullString
    tDate2 = vbNullString
    tUser1 = vbNullString
    tUser2 = vbNullString
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    RS.CursorLocation = adUseClient
    RS.Open "SELECT * FROM Company WHERE CompanyID = " & PK, CN, adOpenStatic, adLockOptimistic
    'Check the form state
    If State = adStateAddMode Then
        Caption = "Create New Entry"
        cmdUsrHistory.Enabled = False
        PK = getIndex("Company")
    Else
        Caption = "Edit Entry"
        DisplayForEditing
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        If State = adStateAddMode Or adStateEditMode Then
            frmCompanyList.RefreshRecords
        End If
    End If
    
    Set frmCompany = Nothing
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
    HLText txtEntry(Index)
End Sub

