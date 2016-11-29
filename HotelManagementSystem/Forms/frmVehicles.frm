VERSION 5.00
Object = "*\A..\..\b8Controls4\b8Controls4.vbp"
Begin VB.Form frmVehicles 
   Caption         =   "Vehicles"
   ClientHeight    =   1620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5610
   LinkTopic       =   "Form2"
   ScaleHeight     =   1620
   ScaleWidth      =   5610
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEntry 
      Height          =   285
      Left            =   1710
      MaxLength       =   100
      TabIndex        =   3
      Tag             =   "Category"
      Top             =   150
      Width           =   3030
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   2370
      TabIndex        =   2
      Top             =   1065
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   3810
      TabIndex        =   1
      Top             =   1065
      Width           =   1335
   End
   Begin VB.CommandButton cmdUsrHistory 
      Caption         =   "Modification History"
      Height          =   315
      Left            =   270
      TabIndex        =   0
      Top             =   1065
      Width           =   1680
   End
   Begin b8Controls4.b8Line b8Line1 
      Height          =   30
      Left            =   90
      TabIndex        =   4
      Top             =   780
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   53
      BorderColor1    =   11325655
      BorderColor2    =   16185592
      BorderStyle1    =   0
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Vehicle"
      Height          =   240
      Left            =   390
      TabIndex        =   5
      Top             =   150
      Width           =   1215
   End
End
Attribute VB_Name = "frmVehicles"
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
        txtEntry.Text = .Fields("Vehicle")
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
    txtEntry.SetFocus
End Sub

Private Sub cmdSave_Click()
    If State = adStateAddMode Then
        RS.AddNew
        RS.Fields("DateAdded") = Now
        RS.Fields("AddedByFK") = CurrUser.USER_PK
    Else
        RS.Fields("DateModified") = Now
        RS.Fields("LastUserFK") = CurrUser.USER_PK
    End If
    'Phill 2:12
    With RS
        .Fields("Vehicle") = txtEntry.Text
    
        .Update
    End With
    
    HaveAction = True
    
    If State = adStateAddMode Then
        MsgBox "New record has been successfully saved.", vbInformation
        If MsgBox("Do you want to add another new record?", vbQuestion + vbYesNo) = vbYes Then
            ResetFields
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
    RS.Open "SELECT * FROM [Vehicles] WHERE VehicleID = " & PK, CN, adOpenStatic, adLockOptimistic
    'Check the form state
    If State = adStateAddMode Then
        Caption = "Create New Entry"
        cmdUsrHistory.Enabled = False
    Else
        Caption = "Edit Entry"
        DisplayForEditing
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        If State = adStateAddMode Or adStateEditMode Then
            frmVehiclesList.RefreshRecords
        End If
    End If
    
    Set frmVehicles = Nothing
End Sub


