VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Records"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6870
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   6870
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbFields 
      Height          =   315
      ItemData        =   "frmSearch.frx":0A02
      Left            =   1800
      List            =   "frmSearch.frx":0A04
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   4995
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   5520
      TabIndex        =   13
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   4200
      TabIndex        =   12
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   " Condition "
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   6615
      Begin VB.TextBox txtFilter 
         Height          =   285
         Index           =   0
         Left            =   3120
         TabIndex        =   3
         Top             =   360
         Width           =   3255
      End
      Begin VB.TextBox txtFilter 
         Height          =   285
         Index           =   1
         Left            =   3120
         TabIndex        =   9
         Top             =   1080
         Width           =   3255
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   285
         Index           =   0
         Left            =   3120
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "MMM-dd-yyyy"
         Format          =   77594627
         CurrentDate     =   38207
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Or"
         Height          =   255
         Left            =   1560
         TabIndex        =   7
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "And"
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   720
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.ComboBox cmbOperation 
         Height          =   315
         Index           =   1
         ItemData        =   "frmSearch.frx":0A06
         Left            =   240
         List            =   "frmSearch.frx":0A22
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1080
         Width           =   2470
      End
      Begin VB.ComboBox cmbOperation 
         Height          =   315
         Index           =   0
         ItemData        =   "frmSearch.frx":0AAA
         Left            =   240
         List            =   "frmSearch.frx":0AC6
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   2470
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   285
         Index           =   1
         Left            =   5040
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "MMM-dd-yyyy"
         Format          =   77594627
         CurrentDate     =   38207
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   285
         Index           =   2
         Left            =   3120
         TabIndex        =   10
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "MMM-dd-yyyy"
         Format          =   77594627
         CurrentDate     =   38207
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   285
         Index           =   3
         Left            =   5040
         TabIndex        =   11
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "MMM-dd-yyyy"
         Format          =   77594627
         CurrentDate     =   38207
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "And"
         Height          =   255
         Left            =   4560
         TabIndex        =   16
         Top             =   1110
         Width           =   375
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "And"
         Height          =   255
         Left            =   4560
         TabIndex        =   15
         Top             =   390
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   2760
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   2760
         Stretch         =   -1  'True
         Top             =   360
         Width           =   240
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Records Where?"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''*****************************************************************
'' File Name:
'' Purpose:
'' Required Files:
''
'' Programmer: Philip V. Naparan   E-mail: philipnaparan@yahoo.com
'' Date Created:
'' Last Modified:
'' Modified By:
'' Credits: NONE, ALL CODES ARE CODED BY Philip V. Naparan
''*****************************************************************

Option Explicit


Public srcColumnHeaders As ColumnHeaders 'Source column headers
Public srcNoOfCol As Long
Public srcForm As Form 'Source form

Private Sub cmbOperation_Click(Index As Integer)
    If Index = 0 Then
        If cmbOperation(Index).ListIndex = 7 Then
            dtpDate(0).Visible = True
            dtpDate(1).Visible = True
            txtFilter(0).Visible = False
        Else
            txtFilter(0).Visible = True
            dtpDate(0).Visible = False
            dtpDate(1).Visible = False
        End If
    Else
        If cmbOperation(Index).ListIndex = 7 Then
            dtpDate(2).Visible = True
            dtpDate(3).Visible = True
            txtFilter(1).Visible = False
        Else
            txtFilter(1).Visible = True
            dtpDate(2).Visible = False
            dtpDate(3).Visible = False
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    'Verify
    If cmbOperation(0).ListIndex <> 7 Then If txtFilter(0).Text = "" Then txtFilter(0).SetFocus: Exit Sub
    
    On Error GoTo err
    Dim strFilter As String
    'Initialize the fields
    strFilter = Replace(cmbFields.Text, "/", "") 'ex. City/Town for tblCustomer
    strFilter = Replace(cmbFields.Text, " ", "")
    strFilter = "[" & strFilter & "]"
    'Initialize the operation used
    'First operation
    Select Case cmbOperation(0).ListIndex
        Case 0: strFilter = strFilter & " LIKE '%" & txtFilter(0).Text & "%'"
        Case 1: strFilter = strFilter & " = '" & txtFilter(0).Text & "'"
        Case 2: strFilter = strFilter & " <> '" & txtFilter(0).Text & "'"
        Case 3: strFilter = strFilter & " > '" & txtFilter(0).Text & "'"
        Case 4: strFilter = strFilter & " >= '" & txtFilter(0).Text & "'"
        Case 5: strFilter = strFilter & " < '" & txtFilter(0).Text & "'"
        Case 6: strFilter = strFilter & " <= '" & txtFilter(0).Text & "'"
        Case 7: strFilter = strFilter & " BETWEEN #" & dtpDate(0).Value & "# AND #" & dtpDate(1).Value & "#"
    End Select
    If cmbOperation(1).Text <> "" Then
        '-Second operation
        If Option1.Value = True Then
            strFilter = strFilter & " AND "
        Else
            strFilter = strFilter & " OR "
        End If
        
        Select Case cmbOperation(1).ListIndex
            Case 0: strFilter = strFilter & " LIKE '%" & txtFilter(1).Text & "%'"
            Case 1: strFilter = strFilter & " = '" & txtFilter(1).Text & "'"
            Case 2: strFilter = strFilter & " <> '" & txtFilter(1).Text & "'"
            Case 3: strFilter = strFilter & " > '" & txtFilter(1).Text & "'"
            Case 4: strFilter = strFilter & " >= '" & txtFilter(1).Text & "'"
            Case 5: strFilter = strFilter & " < '" & txtFilter(1).Text & "'"
            Case 6: strFilter = strFilter & " <= '" & txtFilter(1).Text & "'"
            Case 7: strFilter = strFilter & " BETWEEN #" & dtpDate(2).Value & "# AND #" & dtpDate(3).Value & "#"
        End Select
    End If
        
    'InputBox "", , strFilter
    'Pass the condition to filtered records
    srcForm.FilterRecord strFilter
    'Clear used variables
    strFilter = vbNullString
    
    Unload Me
    Exit Sub
err:
        If err.Number = -2147352571 Then
            MsgBox "Invalid search operation.", vbExclamation
            Unload Me
        ElseIf err.Number = 3001 Then
            Resume Next
        Else
            prompt_err err, "frmFilter", "cmdOk_Click"
        End If
End Sub

Private Sub Form_Load()
    'Initialize values
    dtpDate(0).Value = Date
    dtpDate(1).Value = Date
    dtpDate(2).Value = Date
    dtpDate(3).Value = Date
    'Set the images for the controls
    With mdiMain
        Image1.Picture = .i16x16.ListImages(7).Picture
        Image2.Picture = .i16x16.ListImages(7).Picture
    End With
    
    Dim i As Integer
    If srcNoOfCol = 0 Then srcNoOfCol = srcColumnHeaders.Count
    
    For i = 1 To srcNoOfCol
        If srcColumnHeaders(i).Text <> "" Then cmbFields.AddItem srcColumnHeaders(i).Text
    Next i
    i = 0
    
    cmbFields.ListIndex = 0
    cmbOperation(0).ListIndex = 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmSearch = Nothing
End Sub

Private Sub txtFilter_GotFocus(Index As Integer)
    HLText txtFilter(Index)
End Sub
