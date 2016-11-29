VERSION 5.00
Begin VB.Form frmAccounts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Account"
   ClientHeight    =   3915
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   2910
      TabIndex        =   14
      Top             =   3225
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4380
      TabIndex        =   13
      Top             =   3225
      Width           =   1335
   End
   Begin VB.CommandButton cmdUsrHistory 
      Caption         =   "Modification History"
      Height          =   315
      Left            =   60
      TabIndex        =   12
      Top             =   3300
      Width           =   1680
   End
   Begin VB.TextBox txtPassword 
      DataField       =   "Password"
      DataMember      =   "Employees"
      DataSource      =   "deHotel"
      Height          =   285
      Left            =   2370
      TabIndex        =   11
      Top             =   2475
      Width           =   3375
   End
   Begin VB.TextBox txtUserName 
      DataField       =   "UserName"
      DataMember      =   "Employees"
      DataSource      =   "deHotel"
      Height          =   285
      Left            =   2370
      TabIndex        =   9
      Top             =   2085
      Width           =   1650
   End
   Begin VB.TextBox txtSecLevel 
      DataField       =   "SecLevel"
      DataMember      =   "Employees"
      DataSource      =   "deHotel"
      Height          =   285
      Left            =   2370
      TabIndex        =   7
      Top             =   1710
      Width           =   165
   End
   Begin VB.TextBox txtPosition 
      DataField       =   "Position"
      DataMember      =   "Employees"
      DataSource      =   "deHotel"
      Height          =   285
      Left            =   2370
      TabIndex        =   5
      Top             =   1335
      Width           =   3375
   End
   Begin VB.TextBox txtName 
      DataField       =   "Name"
      DataMember      =   "Employees"
      DataSource      =   "deHotel"
      Height          =   285
      Left            =   2370
      TabIndex        =   3
      Top             =   945
      Width           =   3375
   End
   Begin VB.TextBox txtEmployeeCode 
      DataField       =   "EmployeeCode"
      DataMember      =   "Employees"
      DataSource      =   "deHotel"
      Height          =   285
      Left            =   2370
      TabIndex        =   1
      Top             =   570
      Width           =   3300
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   255
      Index           =   5
      Left            =   525
      TabIndex        =   10
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "UserName:"
      Height          =   255
      Index           =   4
      Left            =   525
      TabIndex        =   8
      Top             =   2130
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "SecLevel:"
      Height          =   255
      Index           =   3
      Left            =   525
      TabIndex        =   6
      Top             =   1755
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Position:"
      Height          =   255
      Index           =   2
      Left            =   525
      TabIndex        =   4
      Top             =   1380
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   255
      Index           =   1
      Left            =   525
      TabIndex        =   2
      Top             =   990
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "EmployeeCode:"
      Height          =   255
      Index           =   0
      Left            =   525
      TabIndex        =   0
      Top             =   615
      Width           =   1815
   End
End
Attribute VB_Name = "frmAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public State                As FormState 'Variable used to determine on how the form used
Public PK                   As Long 'Variable used to get what record is going to edit

Dim HaveAction              As Boolean 'Variable used to detect if the user perform some action
Dim RS                      As New Recordset

Private Sub DisplayForEditing()
    On Error GoTo err
    Dim rsClients As New Recordset
    
    rsClients.CursorLocation = adUseClient
    rsClients.Open "SELECT * FROM qry_Clients WHERE ClientID = " & PK, CN, adOpenStatic, adLockOptimistic
    
    With rsClients
        txtEntry(1).Text = .Fields("Company")
        dcCategory.BoundText = .Fields![CategoryID]
        txtEntry(2).Text = .Fields("Tin")
        txtEntry(3).Text = .Fields("OwnersName")
        txtEntry(4).Text = .Fields("Address")
        dcCity.BoundText = .Fields![CityID]
        txtEntry(6).Text = .Fields("PurchaserName")
        txtEntry(7).Text = .Fields("Mobile")
        txtEntry(8).Text = .Fields("Landline")
        txtEntry(9).Text = .Fields("Fax")
        txtEntry(14).Text = .Fields("CreditTerm")
        txtEntry(15).Text = .Fields("CreditLimit")
        chkBlackListed.Value = IIf(.Fields("BlackListed") = True, 1, 0)
        txtEntry(16).Text = .Fields("Remarks")
    End With
    
    'Display the details
    Dim rsClientBank As New Recordset

    cIRowCount = 0
    
    rsClientBank.CursorLocation = adUseClient
    rsClientBank.Open "SELECT * FROM qry_Clients_Bank WHERE ClientID=" & PK, CN, adOpenStatic, adLockOptimistic
    
    If rsClientBank.RecordCount > 0 Then
        rsClientBank.MoveFirst
        While Not rsClientBank.EOF
          cIRowCount = cIRowCount + 1     'increment
            With Grid
                If .Rows = 2 And .TextMatrix(1, 5) = "" Then
                    .TextMatrix(1, 1) = rsClientBank![Bank]
                    .TextMatrix(1, 2) = rsClientBank![Branch]
                    .TextMatrix(1, 3) = rsClientBank![AccountNo]
                    .TextMatrix(1, 4) = rsClientBank![AccountName]
                    .TextMatrix(1, 5) = rsClientBank![BankID]
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = rsClientBank![Bank]
                    .TextMatrix(.Rows - 1, 2) = rsClientBank![Branch]
                    .TextMatrix(.Rows - 1, 3) = rsClientBank![AccountNo]
                    .TextMatrix(.Rows - 1, 4) = rsClientBank![AccountName]
                    .TextMatrix(.Rows - 1, 5) = rsClientBank![BankID]
                End If
            End With
            rsClientBank.MoveNext
        Wend
        Grid.Row = 1
        Grid.ColSel = 5
        'Set fixed cols
        If State = adStateEditMode Then
            Grid.FixedRows = Grid.Row: 'Grid.SelectionMode = flexSelectionFree
            Grid.FixedCols = 1
        End If
    End If

    rsClientBank.Close
    'Clear variables
    Set rsClientBank = Nothing
        
    'txtEntry(1).SetFocus
    Exit Sub
err:
    If err.Number = 94 Then Resume Next
    
    prompt_err err, Name, "DisplayForEditing"
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub ResetFields()
  clearText Me
  
  txtEntry(15).Text = "0.00"
  txtEntry(1).SetFocus
End Sub

Private Sub cmdSave_Click()
    On Error GoTo err

    If Trim(txtEntry(1).Text) = "" Then Exit Sub
    
    CN.BeginTrans

    If State = adStateAddMode Or State = adStatePopupMode Then
        RS.AddNew
        
        RS.Fields("ClientID") = PK
        RS.Fields("AddedByFK") = CurrUser.USER_PK
    Else
        RS.Fields("DateModified") = Now
        RS.Fields("LastUserFK") = CurrUser.USER_PK
    End If
    
    With RS
      .Fields("Company") = txtEntry(1).Text
      .Fields("CategoryID") = dcCategory.BoundText
      .Fields("Tin") = txtEntry(2).Text
      .Fields("OwnersName") = txtEntry(3).Text
      .Fields("Address") = txtEntry(4).Text
      .Fields("CityID") = dcCity.BoundText
      .Fields("PurchaserName") = txtEntry(6).Text
      .Fields("Mobile") = txtEntry(7).Text
      .Fields("Landline") = txtEntry(8).Text
      .Fields("Fax") = txtEntry(9).Text
      .Fields("CreditTerm") = toNumber(txtEntry(14).Text)
      .Fields("CreditLimit") = toNumber(txtEntry(15).Text)
      .Fields("BlackListed") = IIf(chkBlackListed.Value = 1, True, False)
      .Fields("Remarks") = txtEntry(16).Text
       
      .Update
    End With

    Dim rsClientBank As New Recordset

    rsClientBank.CursorLocation = adUseClient
    rsClientBank.Open "SELECT * FROM Clients_Bank WHERE ClientID=" & PK, CN, adOpenStatic, adLockOptimistic
    
    DeleteItems
    
    Dim c As Integer
    
    With Grid
        'Save the details of the records
        For c = 1 To cIRowCount
            .Row = c
            If State = adStateAddMode Or State = adStatePopupMode Then
AddNew:
                'Add qty received in Local Purchase Details
                rsClientBank.AddNew

                rsClientBank![ClientID] = PK
                rsClientBank![BankID] = toNumber(.TextMatrix(c, 5))
                rsClientBank![AccountNo] = .TextMatrix(c, 3)
                rsClientBank![AccountName] = .TextMatrix(c, 4)

                rsClientBank.Update
            ElseIf State = adStateEditMode Then
                rsClientBank.Filter = "BankID = " & toNumber(.TextMatrix(c, 5))
            
                If rsClientBank.RecordCount = 0 Then GoTo AddNew

                rsClientBank![ClientID] = PK
                rsClientBank![BankID] = toNumber(.TextMatrix(c, 5))
                rsClientBank![AccountNo] = .TextMatrix(c, 3)
                rsClientBank![AccountName] = .TextMatrix(c, 4)

                rsClientBank.Update
            End If

        Next c
    End With

    'Clear variables
    c = 0
    Set rsClientBank = Nothing
    
    CN.CommitTrans

    HaveAction = True
    
    If State = adStateAddMode Then
        MsgBox "New record has been successfully saved.", vbInformation
        If MsgBox("Do you want to add another new record?", vbQuestion + vbYesNo) = vbYes Then
            ResetFields
         Else
            Unload Me
        End If
    ElseIf State = adStatePopupMode Then
        MsgBox "New record has been successfully saved.", vbInformation
        Unload Me
    Else
        MsgBox "Changes in  record has been successfully saved.", vbInformation
        Unload Me
    End If

    Exit Sub

err:
    CN.RollbackTrans
    prompt_err err, Name, "cmdSave_Click"
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdUsrHistory_Click()
    On Error Resume Next
    Dim tDate1 As String
    Dim tDate2 As String
    Dim tUser1 As String
    Dim tUser2 As String
    
    tDate1 = Format$(RS.Fields("DateAdded"), "MMM-dd-yyyy HH:MM AMPM")
    tDate2 = Format$(RS.Fields("DateModified"), "MMM-dd-yyyy HH:MM AMPM")
    
    tUser1 = getValueAt("SELECT PK,CompleteName FROM Users WHERE PK = " & RS.Fields("AddedByFK"), "CompleteName")
    tUser2 = getValueAt("SELECT PK,CompleteName FROM Users WHERE PK = " & RS.Fields("LastUserFK"), "CompleteName")
    
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
    If KeyAscii = 13 And blnRemarks = False Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    RS.CursorLocation = adUseClient
    RS.Open "SELECT * FROM Clients WHERE ClientID = " & PK, CN, adOpenStatic, adLockOptimistic
        
    rsClientBank.CursorLocation = adUseClient
    rsClientBank.Open "SELECT * FROM qry_Clients_Bank WHERE ClientID = " & PK, CN, adOpenStatic, adLockOptimistic
    
    bind_dc "SELECT * FROM Clients_Category", "Category", dcCategory, "CategoryID", True
    bind_dc "SELECT * FROM Cities", "City", dcCity, "CityID", True
   
    'Check the form state
    If State = adStateAddMode Or State = adStatePopupMode Then
        Caption = "Create New Entry"
        cmdUsrHistory.Enabled = False
        
        GeneratePK
    Else
        Caption = "Edit Entry"
        DisplayForEditing
        cmdPH.Enabled = True
    End If

End Sub

'Procedure used to generate PK
Private Sub GeneratePK()
    PK = getIndex("Clients")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        If State = adStateAddMode Or State = adStateEditMode Then
            frmCustomers.RefreshRecords
        ElseIf State = adStatePopupMode Then
            srcText.Text = txtEntry(0).Text
            srcText.Tag = PK
            On Error Resume Next
            srcTextAdd.Text = RS![DisplayAddr]
            srcTextCP.Text = txtEntry(6).Text
            'srcTextDisc.Text = toNumber(cmdDisc.Text)
        End If
    End If
    
    Set frmCustomersAE = Nothing
End Sub

Private Sub ResetEntry()
    txtBranch.Text = ""
    txtAcctNo.Text = ""
    txtAcctName.Text = ""
End Sub
