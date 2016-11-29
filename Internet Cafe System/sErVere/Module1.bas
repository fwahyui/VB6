Attribute VB_Name = "modMain"
Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Public Conn As ADODB.Connection
'Client Logs Dbase
Public Rst As ADODB.Recordset
'BackUp Dbase
Public SvrDbRst As ADODB.Recordset
'Lock/Unlock Client Monitoring Dbase
Public Mon_Rst As ADODB.Recordset
'KWH DBase
Public KwhMon_Rst As ADODB.Recordset

Public SvrStrConn As String
Public NumberComps As Byte
Public Export As Boolean
Public SavePath As String
Public SetPass As Boolean
Public Config As Boolean
Public TypeTrans As Boolean 'OPEN=True; LIMITED=False
Public RENTAL_RATE As Integer
Public INTERNET_RATE As Integer
Public MIN_AMT As Byte
Public TrayPass As Boolean

'Temporary computer Data
Public tmpService As Byte
Public tmpPreviousElapse As Long
Public tmpCurrentElapse As Long
Public tmpStartLog As String
Public tmpEndLog As String
Public tmpName As String
Public tmpPreviousAmount As Integer
Public tmpCurrentAmount As Integer
Public tmpAmount_Limited As Integer
Public tmpInitialStart As String
Public tmpConnected As Byte
Public tmpAccount As Byte
Public tmpExceeded As Boolean
Public tmpComputerNumber As Integer

'Excel
Public oXL As Excel.Application
Public oWB As Excel.Workbook
Public oSheet As Excel.Worksheet
Public oRng As Excel.Range

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long

'Ini Class
'Public INI As cINI

#If Win16 Then
  Public Type RECT
    Left As Integer
    Top As Integer
    Right As Integer
    Bottom As Integer
  End Type
#Else
  Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
  End Type
#End If

#If Win16 Then
    Declare Sub GetWindowRect Lib "User" (ByVal hwnd As Integer, lpRect As RECT)
    Declare Function GetDC Lib "User" (ByVal hwnd As Integer) As Integer
    Declare Function ReleaseDC Lib "User" (ByVal hwnd As Integer, ByVal hdc As Integer) As Integer
    Declare Sub SetBkColor Lib "GDI" (ByVal hdc As Integer, ByVal crColor As Long)
    Declare Sub Rectangle Lib "GDI" (ByVal hdc As Integer, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
    Declare Function CreateSolidBrush Lib "GDI" (ByVal crColor As Long) As Integer
    Declare Function SelectObject Lib "GDI" (ByVal hdc As Integer, ByVal hObject As Integer) As Integer
    Declare Sub DeleteObject Lib "GDI" (ByVal hObject As Integer)
#Else
    Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
    Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
    Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
    Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
    Declare Function SelectObject Lib "user32" (ByVal hdc As Long, ByVal hObject As Long) As Long
    Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
#End If

Public Sub ImplodeForm(f As Form, Direction As Integer, Movement As Integer, ModalState As Integer)
    Dim myRect As RECT
    Dim formWidth%, formHeight%, i%, X%, Y%, cX%, cY%
    Dim TheScreen As Long
    Dim Brush As Long
    
    GetWindowRect f.hwnd, myRect
    formWidth = (myRect.Right - myRect.Left)
    formHeight = myRect.Bottom - myRect.Top
    TheScreen = GetDC(0)
    Brush = CreateSolidBrush(f.BackColor)
    
        For i = Movement To 1 Step -1
        cX = formWidth * (i / Movement)
        cY = formHeight * (i / Movement)
        X = myRect.Left + (formWidth - cX) / 2
        Y = myRect.Top + (formHeight - cY) / 2
        Rectangle TheScreen, X, Y, X + cX, Y + cY
    Next i
    
    X = ReleaseDC(0, TheScreen)
    DeleteObject (Brush)
        
End Sub

Public Sub ExplodeForm(f As Form, Movement As Integer)
  Dim myRect As RECT
  Dim formWidth%, formHeight%, i%, X%, Y%, cX%, cY%
  Dim TheScreen As Long
  Dim Brush As Long
  
  GetWindowRect f.hwnd, myRect
  formWidth = (myRect.Right - myRect.Left)
  formHeight = myRect.Bottom - myRect.Top
  TheScreen = GetDC(0)
  Brush = CreateSolidBrush(f.BackColor)
  
  For i = 1 To Movement
      cX = formWidth * (i / Movement)
      cY = formHeight * (i / Movement)
      X = myRect.Left + (formWidth - cX) / 2
      Y = myRect.Top + (formHeight - cY) / 2
      Rectangle TheScreen, X, Y, X + cX, Y + cY
  Next i
  
  X = ReleaseDC(0, TheScreen)
  DeleteObject (Brush)
  
End Sub

Public Function MoveForm(TheForm As Form)
    Dim Ret
    ReleaseCapture
    SendMessage TheForm.hwnd, &HA1, 2, 0&
End Function

Public Sub Open_Connection()
  Set Conn = New ADODB.Connection
  With Conn
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Caferitos.mdb;Jet OLEDB:Database Password=GallanosA;"
    .CommandTimeout = 0
    .CursorLocation = adUseClient
    .Open
  End With
End Sub

Sub Main()
  Dim i As Byte
  
  frmSplash.lblStatus.Caption = "Loading Personal Web Server..."
  Load frmWebServer
  'filters to run only one app itself
  'If App.PrevInstance = True Then End

  frmSplash.Bar.Value = 20
  
  'set INI class
  'Set INI = New cINI
  
  Open_Connection
  'Backup Database
  Set SvrDbRst = New ADODB.Recordset
  SvrStrConn = "SELECT * FROM tblBackUp"
  frmSplash.lblStatus.Caption = "Loading Backup DBase..."
  DoEvents
  SvrDbRst.Open SvrStrConn, Conn, adOpenDynamic, adLockOptimistic
  
  frmSplash.imgScreen2.Picture = frmSplash.imgScreen.Picture
  frmSplash.Bar.Value = 50
  
  'Client Logs Database
  Set Rst = New ADODB.Recordset
  SvrStrConn = "SELECT * FROM tblLogs ORDER by ID"
  frmSplash.lblStatus.Caption = "Loading Main DBase..."
  DoEvents
  Rst.Open SvrStrConn, Conn, adOpenDynamic, adLockOptimistic
  
  frmSplash.imgScreen.Picture = LoadPicture("")
  frmSplash.Bar.Value = 75
  
  'Database of Lock/Unlock of the Clients
  Set Mon_Rst = New ADODB.Recordset
  SvrStrConn = "SELECT * FROM tblClientMon"
  frmSplash.lblStatus.Caption = "Loading Lock/Unlock DBase..."
  DoEvents
  Mon_Rst.Open SvrStrConn, Conn, adOpenDynamic, adLockOptimistic
  
  frmSplash.Bar.Value = 90
  
  'KWH Dbase
  Set KwhMon_Rst = New ADODB.Recordset
  SvrStrConn = "SELECT * FROM tblKWH"
  frmSplash.lblStatus.Caption = "Loading KWH DBase..."
  DoEvents
  KwhMon_Rst.Open SvrStrConn, Conn, adOpenDynamic, adLockOptimistic
  
  frmSplash.Bar.Value = 100
  
  Sleep 1000
  
  Rst.MoveFirst
  INTERNET_RATE = Rst!Month
  RENTAL_RATE = Rst!Day
  MIN_AMT = Rst!Year
  NumberComps = Rst!CompNum
  
'  frmMain.Show
  Check_Stagnant
  frmMain.tmrResume.Enabled = True
  Unload frmSplash
  
End Sub

Public Sub Check_Stagnant()
On Error Resume Next
Dim i As Byte
  
  With SvrDbRst
    For i = 1 To NumberComps
      .MoveFirst
      .Find "ComNum LIKE " & i, 1, adSearchForward
      If !StartLog <> "" And !Unlock = False Then
        frmMain.lvMain.ListItems(Val(!ComNum)).ListSubItems(1).Text = !Name 'Name
        'Account Type
        If !Account = v_Open Then 'Open
          frmMain.lvMain.ListItems(Val(!ComNum)).ListSubItems(2).Text = "OPEN"
          frmMain.lvMain.ListItems(Val(!ComNum)).ListSubItems(7).Text = FormatNumber(!Amt, 2) 'Amount
        Else 'Limited
          frmMain.lvMain.ListItems(Val(!ComNum)).ListSubItems(2).Text = "LIMITED"
          frmMain.lvMain.ListItems(Val(!ComNum)).ListSubItems(7).Text = FormatNumber(!AmtLimited, 2) 'Amount
        End If
        
        frmMain.lvMain.ListItems(Val(!ComNum)).ListSubItems(6).Text = Formatter(!Elapse) 'Elapse
        
        frmMain.lvMain.ListItems(Val(!ComNum)).ListSubItems(4).Text = FormatDateTime(!StartLog, vbLongTime) 'Start Log
        'Service Type
        If !Service = v_Internet Then
          frmMain.lvMain.ListItems(Val(!ComNum)).ListSubItems(3).Text = "Internet"
        Else
          frmMain.lvMain.ListItems(Val(!ComNum)).ListSubItems(3).Text = "Gms/Rntl"
        End If
      End If
    Next i
  End With
End Sub

Public Function Formatter2(B As Integer) As String
  Select Case Len(Str(B))
  Case 1
    Formatter2 = "000" & Trim(Str(B))
  Case 2
    Formatter2 = "00" & Trim(Str(B))
  Case 3
    Formatter2 = "0" & Trim(Str(B))
  Case Else
    Formatter2 = Trim(Str(B))
  End Select
End Function

Public Function Excel_Daily() As Boolean
Dim Ctr As Integer
Dim i As Integer
Dim from As Long
  
  Ctr = 0
  Set oWB = oXL.Workbooks.Add(App.Path & "\Template\Report.xlt")
  Set oSheet = oWB.ActiveSheet
  
  Rst.MoveFirst
  For i = 0 To 10000
  
Ulit:
    Rst.Find "Year LIKE " & frmDailyReport.Cal.Year, 1, adSearchForward
    If Rst.EOF = True Then Exit For
    If Rst!Month = frmDailyReport.Cal.Month And Rst!Day = frmDailyReport.Cal.Day Then
      oSheet.Cells(5 + Ctr, 2).Value = Str(Rst!Month) & "/" & Str(Rst!Day) & "/" & Right(Str(Rst!Year), 2) 'Date
      oSheet.Cells(5 + Ctr, 3).Value = Rst!CompNum 'Computer #
      If Rst!Service = v_Internet Then
        oSheet.Cells(5 + Ctr, 4).Value = "Internet" 'Internet
      Else
        oSheet.Cells(5 + Ctr, 4).Value = "Games/Rental" 'Internet
      End If
      oSheet.Cells(5 + Ctr, 5).Value = Rst!StartLog 'login
      oSheet.Cells(5 + Ctr, 6).Value = Rst!EndLog 'Logout
      oSheet.Cells(5 + Ctr, 7).Value = Rst!Elapse 'Duration
      oSheet.Cells(5 + Ctr, 8).Value = Rst!Amt 'Amount
      Ctr = Ctr + 1
    Else
      GoTo Ulit
    End If
  Next i
    
  If Ctr = 0 Then
    MsgBox "No records found!"
    Excel_Daily = False
    Exit Function
  Else
    oSheet.Cells(3, 5).Value = Format(frmDailyReport.Cal.Value, "dddddd")
    oSheet.Cells(5 + Ctr, 7).Value = "________________________________"
    oSheet.Cells(6 + Ctr, 6).Value = "TOTAL"
    oSheet.Cells(6 + Ctr, 8).Value = "=SUM(H5:H" & Trim(Str(Ctr + 4)) & ")"
  End If
  
  KwhMon_Rst.MoveFirst
  Do Until KwhMon_Rst.EOF = True
    KwhMon_Rst.Find "Month LIKE " & Format(frmDailyReport.Cal.Value, "m"), 1, adSearchForward
    If KwhMon_Rst.EOF = False Then
      If KwhMon_Rst!Day = Val(Format(frmDailyReport.Cal.Value, "d")) _
      And KwhMon_Rst!Year = Val(Format(frmDailyReport.Cal.Value, "yyyy")) Then
        KwhMon_Rst.MovePrevious
        from = KwhMon_Rst!KwhRead
        KwhMon_Rst.MoveNext
        oSheet.Cells(7 + Ctr, 3).Value = "From"
        oSheet.Cells(8 + Ctr, 3).Value = from
        oSheet.Cells(7 + Ctr, 4).Value = "To"
        oSheet.Cells(8 + Ctr, 4).Value = KwhMon_Rst!KwhRead
        oSheet.Cells(7 + Ctr, 5).Value = "Used"
        oSheet.Cells(8 + Ctr, 5).Value = "<" & KwhMon_Rst!KwhRead - from & ">"
        oSheet.Cells(7 + Ctr, 2).Value = "Kwh"
        oSheet.Cells(8 + Ctr, 2).Value = "Reading"
        Exit Do
      End If
    End If
  Loop
  
  If Export = True Then ' Export
    oSheet.SaveAs SavePath
  Else 'Print
    'oXL.Visible = True
    oSheet.PrintOut
  End If
  Excel_Daily = True
End Function

Public Function Excel_Monthly(mnth As Integer, yr As Integer) As Boolean
Dim Ctr As Integer
Dim i As Integer
Dim DayTotal As Integer
Dim DayPrm As Byte
Dim ExitFor As Boolean
Dim Logs As Integer
Dim from As Long

  Ctr = 0
  DayPrm = 40
  DayTotal = 0
  Logs = 0
  ExitFor = False
  Set oWB = oXL.Workbooks.Add(App.Path & "\Template\Report2.xlt")
  Set oSheet = oWB.ActiveSheet
    
  Rst.MoveFirst
  KwhMon_Rst.MoveFirst
  
  For i = 0 To 10000
Ulit:
    Rst.Find "Year LIKE " & yr, 1, adSearchForward
    If Rst.EOF = True Then
      ExitFor = True
      GoTo Iron
    End If
    If Rst!Month = mnth Then
      If DayPrm <> Rst!Day Then
        If DayPrm <> 40 Then
Iron:
          Do Until KwhMon_Rst.EOF = True
            KwhMon_Rst.Find "Month LIKE " & mnth, 1, adSearchForward
            If KwhMon_Rst.EOF = False Then
              If KwhMon_Rst!Day = DayPrm And KwhMon_Rst!Year = yr Then
                KwhMon_Rst.MovePrevious
                from = KwhMon_Rst!KwhRead
                KwhMon_Rst.MoveNext
                oSheet.Cells(5 + Ctr, 4).Value = from & "-" & KwhMon_Rst!KwhRead 'from-to kwh reading
                oSheet.Cells(5 + Ctr, 5).Value = KwhMon_Rst!KwhRead - from 'kwh used
                Exit Do
              End If
            End If
          Loop
       
          oSheet.Cells(5 + Ctr, 2).Value = DayPrm & "  " & Format(mnth & "/" & DayPrm & "/" & yr, "dddd") 'Day
          oSheet.Cells(5 + Ctr, 3).Value = Logs 'No of Logs
          oSheet.Cells(5 + Ctr, 6).Value = DayTotal 'TotalAmt
          Ctr = Ctr + 1
          If ExitFor = True Then
            Exit For
          End If
        End If
        DayPrm = Rst!Day
        DayTotal = Rst!Amt
        Logs = 1
      Else
        DayTotal = DayTotal + Rst!Amt
        Logs = Logs + 1
      End If
    Else
      GoTo Ulit
    End If
  Next i
    
  If Ctr = 0 Then
    MsgBox "No records found!"
    Excel_Monthly = False
    Exit Function
  Else
    oSheet.Cells(3, 2).Value = frmMonthlyReport.cmbMonth.Text & " " & frmMonthlyReport.cmbYear.Text
    oSheet.Cells(5 + Ctr, 5).Value = "____________________________________________"
    oSheet.Cells(6 + Ctr, 4).Value = "TOTALS"
    oSheet.Cells(6 + Ctr, 5).Value = "=SUM(E5:E" & Trim(Str(Ctr + 4)) & ")" 'Total kwh used for the month
    oSheet.Cells(6 + Ctr, 6).Value = "=SUM(F5:F" & Trim(Str(Ctr + 4)) & ")" 'Total Gross Income
  End If
        
  If Export = True Then ' Export
    oWB.SaveAs SavePath
  Else 'Print
    'oXL.Visible = True
    oSheet.PrintOut
  End If
  Excel_Monthly = True
End Function

Public Function Formatter(B As Long) As String
  If (B Mod 60) < 30 Then
    Formatter = Str(FormatNumber(B / 60, 0)) & " hrs " & (B Mod 60) & " mins"
  Else
    Formatter = Str(FormatNumber(B / 60, 0) - 1) & " hrs " & (B Mod 60) & " mins"
  End If
End Function

Public Function AddSpace(expr As String, tot_lenght As Byte) As String
  AddSpace = Space(tot_lenght - Len(Trim(expr))) & expr
End Function

Public Sub SynchronizeFlash()
Dim i As Byte
  For i = 1 To frmMain.Client.UBound - 1
    frmMain.Client(i).SendData "SZ" & "@"
  Next i
End Sub

Public Sub Remove_BackUp_Data(ComputerNmbr As Byte)
  SvrDbRst.MoveFirst
  SvrDbRst.Find "ComNum LIKE " & ComputerNmbr, 1, adSearchForward
  
  SvrDbRst!Name = Null
  SvrDbRst!StartLog = Null
  SvrDbRst!Elapse = 0
  SvrDbRst!Amt = 0
  SvrDbRst!Account = 0
  SvrDbRst!Service = 0
  SvrDbRst!AmtLimited = 0
  SvrDbRst!Unlock = False
  SvrDbRst.Update
End Sub

Public Sub Remove_Grid_Data(ComputerNmbr As Byte)
Dim Ctr As Byte
  'Clear Data Grid
  For Ctr = 1 To 7
    frmMain.lvMain.ListItems(ComputerNmbr).ListSubItems(Ctr).Text = ""
  Next Ctr
End Sub
