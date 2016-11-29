VERSION 5.00
Object = "{9EDDC69F-10E8-4DE7-9648-EC8A45F005C0}#1.0#0"; "b8Controls4.ocx"
Begin VB.Form frmUsers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Users"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5610
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   5610
   StartUpPosition =   2  'CenterScreen
   Begin b8Controls4.b8Line b8Line1 
      Height          =   30
      Left            =   90
      TabIndex        =   10
      Top             =   1770
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   53
      BorderColor1    =   11325655
      BorderColor2    =   16185592
   End
   Begin VB.CommandButton cmdUsrHistory 
      Caption         =   "Modification History"
      Height          =   315
      Left            =   210
      TabIndex        =   6
      Top             =   1995
      Width           =   1680
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4200
      TabIndex        =   5
      Top             =   1995
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   2760
      TabIndex        =   4
      Top             =   1995
      Width           =   1335
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   1
      Left            =   1335
      MaxLength       =   100
      TabIndex        =   3
      Tag             =   "Password"
      Top             =   495
      Width           =   2415
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   0
      Left            =   1335
      MaxLength       =   100
      TabIndex        =   2
      Tag             =   "Username"
      Top             =   120
      Width           =   2415
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   2
      Left            =   1335
      MaxLength       =   100
      TabIndex        =   1
      Tag             =   "Complete Name"
      Top             =   945
      Width           =   3840
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Admin"
      Height          =   240
      Left            =   1335
      TabIndex        =   0
      Top             =   1320
      Width           =   990
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Username"
      Height          =   240
      Index           =   1
      Left            =   285
      TabIndex        =   9
      Top             =   120
      Width           =   990
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Password"
      Height          =   240
      Index           =   0
      Left            =   360
      TabIndex        =   8
      Top             =   495
      Width           =   915
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Complete Name"
      Height          =   240
      Index           =   2
      Left            =   135
      TabIndex        =   7
      Top             =   945
      Width           =   1140
   End
End
Attribute VB_Name = "frmUsers"
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
        txtEntry(0).Text = .Fields("UserID")
        txtEntry(1).Text = Enc.DecryptString(.Fields("Password"))
        
        txtEntry(2).Text = .Fields("CompleteName")
        Check1.Value = changeYNValue(.Fields("Admin"))
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
    Check1.Value = 0
    txtEntry(0).SetFocus
End Sub

Private Sub cmdSave_Click()
    If is_empty(txtEntry(0), True) = True Then Exit Sub
    If is_empty(txtEntry(1), True) = True Then Exit Sub
    If is_empty(txtEntry(2), True) = True Then Exit Sub
    
    If State = adStateAddMode Then
        RS.AddNew
        RS.Fields("PK") = PK
        RS.Fields("DateAdded") = Now
        RS.Fields("AddedByFK") = CurrUser.USER_PK
        
        Call AddPermission
    Else
        RS.Fields("DateModified") = Now
        RS.Fields("LastUserFK") = CurrUser.USER_PK
    End If
    'Phill 2:12
    With RS
        .Fields("UserID") = txtEntry(0).Text
        .Fields("Password") = Enc.EncryptString(txtEntry(1).Text)
        .Fields("CompleteName") = txtEntry(2).Text
        .Fields("Admin") = changeYNValue(Check1.Value)
        .Update
    End With
    
    HaveAction = True
    
    If State = adStateAddMode Then
        MsgBox "New record has been successfully saved.", vbInformation
        If MsgBox("Do you want to add another new record?", vbQuestion + vbYesNo) = vbYes Then
            ResetFields
            PK = getIndex("tbl_SM_Users")
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

Private Sub Form_Load()
    RS.CursorLocation = adUseClient
    RS.Open "SELECT * FROM Users WHERE PK = " & PK, CN, adOpenStatic, adLockOptimistic
    'Check the form state
    If State = adStateAddMode Then
        Caption = "Create New Entry"
        cmdUsrHistory.Enabled = False
        PK = getIndex("Users")
    Else
        Caption = "Edit Entry"
        DisplayForEditing
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        If State = adStateAddMode Or adStateEditMode Then
            frmUsersList.CommandPass 5
        End If
    End If
    
    Set frmUsers = Nothing
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
    HLText txtEntry(Index)
End Sub

Public Sub AddPermission()
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "INSERT INTO [User Permission] ( UserID, FormID ) " _
            & "SELECT '" & Me.txtEntry(0).Text & "', Form.FormID " _
            & "FROM Form"

    CN.Execute sSQL
    
    Exit Sub
    
RAE:
    Set vRS = Nothing
End Sub


