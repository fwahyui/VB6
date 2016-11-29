VERSION 5.00
Object = "{9EDDC69F-10E8-4DE7-9648-EC8A45F005C0}#1.0#0"; "b8Controls4.ocx"
Begin VB.Form frmUserPermission 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Permission"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   6270
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   465
      Left            =   4590
      TabIndex        =   1
      Top             =   5940
      Width           =   1245
   End
   Begin VB.ComboBox cboAllowOpen 
      Height          =   315
      ItemData        =   "frmUserPermission.frx":0000
      Left            =   1410
      List            =   "frmUserPermission.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1890
      Visible         =   0   'False
      Width           =   1455
   End
   Begin b8Controls4.LynxGrid3 listProdPack 
      Height          =   5565
      Left            =   330
      TabIndex        =   2
      Top             =   240
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   9816
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorBkg    =   16056319
      BackColorSel    =   8438015
      ForeColorSel    =   0
      GridColor       =   11136767
      BorderStyle     =   0
      FocusRectColor  =   33023
      AllowUserResizing=   4
      Editable        =   -1  'True
      Striped         =   -1  'True
      SBackColor1     =   16056319
      SBackColor2     =   14940667
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00808080&
      Height          =   4905
      Left            =   330
      Top             =   210
      Width           =   5475
   End
End
Attribute VB_Name = "frmUserPermission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strUser As String

Private Sub cboAllowOpen_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        listProdPack.ToggleEdit
        listProdPack.Refresh
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
'    PaintGrad bgHeader, &HEDEBE9, &HFFFFFF, 0
    
    'set list column
    With listProdPack
        .AddColumn "Form Description", 230
        .AddColumn "User Permission ID", 0
        .AddColumn "Allow Open", 80
        .BindControl 2, cboAllowOpen, lgBCLeft Or lgBCTop Or lgBCWidth
        .RowHeightMin = 21
'        .ImageList = ilList
    End With

    RefreshProdPack strUser
End Sub

Public Sub RefreshProdPack(ByVal UserID As String)
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim li As Long
    
    
    listProdPack.Redraw = False
    listProdPack.Clear
    
    sSQL = "SELECT [User Permission].UserPermissionID, [User Permission].UserID, Form.Description, [User Permission].AllowOpen " _
            & "FROM Form INNER JOIN [User Permission] ON Form.FormID = [User Permission].FormID " _
            & "WHERE UserID = '" & strUser & "'"

    vRS.Open sSQL, CN, adOpenStatic, adLockPessimistic
    
    'fill
    vRS.MoveFirst
    While vRS.EOF = False
        With listProdPack
            li = .AddItem(vRS.Fields("Description"))
            .ItemImage(li) = 1
            .CellText(li, 1) = vRS.Fields("UserPermissionID")
            .CellText(li, 2) = vRS.Fields("AllowOpen")
        End With
        vRS.MoveNext
    Wend
    
RAE:
    Set vRS = Nothing
    listProdPack.Redraw = True
    listProdPack.Refresh
End Sub

Private Sub listProdPack_Click()
    If listProdPack.RowCount > 0 Then
        If listProdPack.Col = 2 Then
            listProdPack.EditCell listProdPack.Row, listProdPack.Col
        End If
    End If
End Sub

Private Sub listProdPack_RequestEdit(Row As Long, Col As Long, Cancel As Boolean)
    Select Case Col
        Case 2
            cboAllowOpen.Text = listProdPack.CellText(Row, Col)
    End Select
End Sub

Private Sub listProdPack_RequestUpdate(Row As Long, Col As Long, NewValue As String, Cancel As Boolean)
    'default
    Cancel = True
    
    'validate
'    If (NewValue <> "True" Or NewValue <> "False") Then
'        MsgBox "Please enter valid value. It must True or False", vbExclamation
'        Exit Sub
'    End If
    
'    NewValue = FormatNumber(GetTxtVal(NewValue), 2)
    NewValue = cboAllowOpen.Text
    
    If ChangePermission(listProdPack.CellText(Row, 1), NewValue) Then
        listProdPack.CellText(Row, Col) = NewValue
        listProdPack.Refresh
    Else
        'WriteErrorLog Me.Name, "listProdPack_RequestUpdate", "Failed on: 'modRSSup.SetSupBegAP(GetTxtVal(listProdPack.celltext(Row, 1)), GetTxtVal(NewValue)) = True'"
        'refresh list
        RefreshProdPack strUser
    End If
    
    Cancel = False
End Sub


