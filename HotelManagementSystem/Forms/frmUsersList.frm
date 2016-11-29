VERSION 5.00
Object = "{9EDDC69F-10E8-4DE7-9648-EC8A45F005C0}#1.0#0"; "b8Controls4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmUsersList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Users"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8835
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   8835
   StartUpPosition =   2  'CenterScreen
   Begin b8Controls4.b8Line b8Line1 
      Height          =   30
      Left            =   30
      TabIndex        =   2
      Top             =   420
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   53
      BorderColor1    =   11325655
      BorderColor2    =   16185592
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   635
      ButtonWidth     =   2037
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "i16x16"
      DisabledImageList=   "i16x16"
      HotImageList    =   "i16x16"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Find"
            Object.ToolTipText     =   "F1"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            Object.ToolTipText     =   "F2"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "F3"
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit"
            Object.ToolTipText     =   "F4"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Object.ToolTipText     =   "F5"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh"
            Object.ToolTipText     =   "F6"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Permission"
            Object.ToolTipText     =   "F7"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   5520
      Top             =   3645
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsersList.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsersList.frx":0A12
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsersList.frx":1424
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsersList.frx":1E36
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsersList.frx":2848
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsersList.frx":325A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   4860
      Left            =   60
      TabIndex        =   1
      Top             =   540
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   8573
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "User Name"
         Object.Width           =   3598
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Complete Name"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Admin"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmUsersList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS As New Recordset

Private Sub Form_Load()
    'Set the graphics needed
    'Set the graphics for the controls
    With mdiMain
        'For listview
        Set lvList.SmallIcons = .i16x16
        Set lvList.Icons = .i16x16
    End With

    RefreshRecords
End Sub

Private Sub RefreshRecords()
    Me.Enabled = False
    If RS.State = adStateOpen Then RS.Close
    RS.Open "SELECT UserID,CompleteName,Admin,PK FROM Users ORDER BY UserID ASC", CN, adOpenStatic, adLockOptimistic
    FillListView lvList, RS, 3, 2, False, True, "PK"
    Me.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmUsersList = Nothing
End Sub

Private Sub lvList_DblClick()
    CommandPass 3
End Sub

Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1: CommandPass 1
        Case 3: CommandPass 2
        Case 5: CommandPass 3
        Case 7: CommandPass 4
        Case 9: CommandPass 5
        Case 11: CommandPass 6
        Case 13: CommandPass 7
    End Select
End Sub

Public Sub CommandPass(ByVal IntCmd As Integer)
On Error GoTo err
    Select Case IntCmd
        'Find
        Case 1
            If lvList.ListItems.Count < 1 Then MsgBox "No record to search.", vbExclamation: Exit Sub
'            With frmFind
'                Set .srcListView = lvList
'                .Show vbModal
'            End With
        'New
        Case 2
            frmUsers.State = adStateAddMode
            frmUsers.Show vbModal
        'Edit
        Case 3
            If lvList.ListItems.Count > 0 Then
                If isRecordExist("Users", "PK", CLng(lvList.SelectedItem.Tag)) = False Then
                    MsgBox "This record has been removed by other user.Click 'OK' button to refresh the records.", vbExclamation
                    RefreshRecords
                    Exit Sub
                Else
                    With frmUsers
                        .State = adStateEditMode
                        .PK = CLng(lvList.SelectedItem.Tag)
                        .Show vbModal
                    End With
                End If
            End If
        'Delete
        Case 4
            If CLng(lvList.SelectedItem.Tag) = CurrUser.USER_PK Then
                MsgBox "You cannot remove your own record because you currently using it.", vbExclamation
                Exit Sub
            Else
                If lvList.ListItems.Count > 0 Then
                    If isRecordExist("Users", "PK", CLng(lvList.SelectedItem.Tag)) = False Then
                        MsgBox "This record has been removed by other user.Click 'OK' button to refresh the records.", vbExclamation
                        RefreshRecords
                        Exit Sub
                    Else
                        Dim ANS As Integer
                        ANS = MsgBox("Are you sure you want to delete the selected record?" & vbCrLf & vbCrLf & "WARNING: You cannot undo this operation.", vbCritical + vbYesNo, "Confirm Record Delete")
                        Me.MousePointer = vbHourglass
                        If ANS = vbYes Then
                            DelRecwSQL "Users", "PK", "", True, CLng(lvList.SelectedItem.Tag)
                            RefreshRecords
                            MsgBox "Record has been successfully deleted.", vbInformation, "Confirm"
                        End If
                        ANS = 0
                        Me.MousePointer = vbDefault
                    End If
                Else
                    MsgBox "No record to delete.", vbExclamation
                End If
            End If
        'Reload
        Case 5: RefreshRecords
        Case 6:
            With frmUserPermission
                .strUser = Me.lvList.SelectedItem.Text
                .Show 1
            End With
        'Close
        Case 7: Unload Me
    End Select
    Exit Sub
err:
    If err.Number = -2147467259 Then
        MsgBox "You cannot delete this record because it was used by other records! If you want to delete this record" & vbCrLf & _
               "you will first have to delete or change the records that currenly used this record as shown bellow." & vbCrLf & vbCrLf & _
               err.Description, , "Delete Operation Failed!"
        Me.MousePointer = vbDefault
    End If
End Sub
