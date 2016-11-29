VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmReports 
   Caption         =   "Form2"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9390
   LinkTopic       =   "Form2"
   ScaleHeight     =   7995
   ScaleWidth      =   9390
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CR 
      Height          =   3915
      Left            =   1470
      TabIndex        =   0
      Top             =   1260
      Width           =   6405
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strReport        As String
Public PK               As String
Public strYear          As String
Public blnPaid          As Boolean
Public strWhere         As String

Dim mTest As CRAXDRT.Application
Dim mReport As CRAXDRT.Report
Dim SubReport As CRAXDRT.Report
Dim mParam As CRAXDRT.ParameterFieldDefinitions

Public Sub CommandPass(ByVal srcPerformWhat As String)
    Select Case srcPerformWhat
        Case "Close"
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
On Error GoTo err_Form_Load
    Dim mSubRep
    
    Set mTest = New CRAXDRT.Application
    Set mReport = New CRAXDRT.Report
    
    Select Case strReport
        Case "Folio"
            Set mReport = mTest.OpenReport(App.Path & "\Reports\rpt_Folio.rpt")

            mReport.RecordSelectionFormula = strWhere

            Set mParam = mReport.ParameterFields
        
            mParam.Item(1).AddCurrentValue CurrBiz.BUSINESS_NAME
            mParam.Item(2).AddCurrentValue CurrBiz.BUSINESS_ADDRESS
            mParam.Item(3).AddCurrentValue CurrBiz.BUSINESS_CONTACT_INFO
        Case "Account Receivable"
            Set mReport = mTest.OpenReport(App.Path & "\Reports\rpt_Account_Receivable.rpt")

            If frmRPTAccRec.dcCompany.Text <> "" Then
                mReport.RecordSelectionFormula = "{qry_Account_Receivable.Company} = '" & frmRPTAccRec.dcCompany.Text & "'"
            End If

            Set mParam = mReport.ParameterFields
        
            mParam.Item(1).AddCurrentValue CurrBiz.BUSINESS_NAME
            mParam.Item(2).AddCurrentValue CurrBiz.BUSINESS_ADDRESS
            mParam.Item(3).AddCurrentValue CurrBiz.BUSINESS_CONTACT_INFO
        Case "Other Charges"
            Set mReport = mTest.OpenReport(App.Path & "\Reports\rpt_Other_Charges.rpt")

            mReport.RecordSelectionFormula = strWhere

            Set mParam = mReport.ParameterFields
        
            mParam.Item(1).AddCurrentValue CurrBiz.BUSINESS_NAME
            mParam.Item(2).AddCurrentValue CurrBiz.BUSINESS_ADDRESS
            mParam.Item(3).AddCurrentValue CurrBiz.BUSINESS_CONTACT_INFO
        Case "CheckIn Guest"
            Set mReport = mTest.OpenReport(App.Path & "\Reports\rpt_CheckIn_Guest.rpt")

            Set mParam = mReport.ParameterFields
        
            mParam.Item(1).AddCurrentValue CurrBiz.BUSINESS_NAME
            mParam.Item(2).AddCurrentValue CurrBiz.BUSINESS_ADDRESS
            mParam.Item(3).AddCurrentValue CurrBiz.BUSINESS_CONTACT_INFO
        Case "Check Out"
            Set mReport = mTest.OpenReport(App.Path & "\Reports\rpt_CheckOut.rpt")

            mReport.RecordSelectionFormula = strWhere

            Set mParam = mReport.ParameterFields
        
            mParam.Item(1).AddCurrentValue CurrBiz.BUSINESS_NAME
            mParam.Item(2).AddCurrentValue CurrBiz.BUSINESS_ADDRESS
            mParam.Item(3).AddCurrentValue CurrBiz.BUSINESS_CONTACT_INFO
        Case "Due Reservation"
            Set mReport = mTest.OpenReport(App.Path & "\Reports\rpt_Due_Reservation.rpt")

            mReport.RecordSelectionFormula = strWhere

            Set mParam = mReport.ParameterFields
        
            mParam.Item(1).AddCurrentValue CurrBiz.BUSINESS_NAME
            mParam.Item(2).AddCurrentValue CurrBiz.BUSINESS_ADDRESS
            mParam.Item(3).AddCurrentValue CurrBiz.BUSINESS_CONTACT_INFO
        Case "Guest List"
            Set mReport = mTest.OpenReport(App.Path & "\Reports\rpt_Guest_List.rpt")

            Set mParam = mReport.ParameterFields
        
            mParam.Item(1).AddCurrentValue CurrBiz.BUSINESS_NAME
            mParam.Item(2).AddCurrentValue CurrBiz.BUSINESS_ADDRESS
            mParam.Item(3).AddCurrentValue CurrBiz.BUSINESS_CONTACT_INFO
        Case "Room History"
            Set mReport = mTest.OpenReport(App.Path & "\Reports\rpt_Room_History.rpt")

            mReport.RecordSelectionFormula = strWhere

            Set mParam = mReport.ParameterFields
        
            mParam.Item(1).AddCurrentValue CurrBiz.BUSINESS_NAME
            mParam.Item(2).AddCurrentValue CurrBiz.BUSINESS_ADDRESS
            mParam.Item(3).AddCurrentValue CurrBiz.BUSINESS_CONTACT_INFO
        Case "Reservation"
            Set mReport = mTest.OpenReport(App.Path & "\Reports\rpt_Reservation.rpt")

            mReport.RecordSelectionFormula = strWhere
        End Select
    
    Screen.MousePointer = vbHourglass
    CR.ReportSource = mReport
    CR.ViewReport
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
err_Form_Load:
    prompt_err err, Name, "Form_Load"
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    With CR
        .Top = 0
        .Left = 0
        .Height = ScaleHeight
        .Width = ScaleWidth
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mdiMain.RemoveChild Me.Name
    
    Set frmReports = Nothing
End Sub
