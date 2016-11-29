VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MWARNET - CALCULATOR"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4215
   Icon            =   "calculator.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "calculator.frx":FA8A
   ScaleHeight     =   4455
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEqual 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3120
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3000
      Width           =   450
   End
   Begin VB.CommandButton cmdOP 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   1
      Left            =   2640
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3000
      Width           =   450
   End
   Begin VB.CommandButton cmdDot 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2160
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3000
      Width           =   450
   End
   Begin VB.CommandButton cmdPlusMinus 
      Caption         =   "+/-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1680
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3000
      Width           =   450
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   0
      Left            =   1200
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   3000
      Width           =   450
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   1
      Left            =   1200
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2520
      Width           =   450
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   2
      Left            =   1680
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2520
      Width           =   450
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   3
      Left            =   2160
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2520
      Width           =   450
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   4
      Left            =   1200
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2040
      Width           =   450
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   5
      Left            =   1680
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2040
      Width           =   450
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   6
      Left            =   2160
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2040
      Width           =   450
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   7
      Left            =   1200
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1590
      Width           =   450
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   8
      Left            =   1680
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1590
      Width           =   450
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   9
      Left            =   2160
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1590
      Width           =   450
   End
   Begin VB.CommandButton cmdOP 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   2
      Left            =   2640
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2520
      Width           =   450
   End
   Begin VB.CommandButton cmdOP 
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   3
      Left            =   2640
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2040
      Width           =   450
   End
   Begin VB.CommandButton cmdOP 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   4
      Left            =   2640
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1590
      Width           =   450
   End
   Begin VB.CommandButton cmdC 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2160
      TabIndex        =   12
      Top             =   3600
      Width           =   1410
   End
   Begin VB.CommandButton cmdCE 
      Caption         =   "CE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1680
      TabIndex        =   9
      Top             =   3600
      Width           =   450
   End
   Begin VB.CommandButton cmdBS 
      Caption         =   "Backspace"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   600
      TabIndex        =   8
      Top             =   3600
      Width           =   1050
   End
   Begin VB.CommandButton cmdMEM 
      Caption         =   "MC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   0
      Left            =   600
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1560
      Width           =   450
   End
   Begin VB.CommandButton cmdMEM 
      Caption         =   "MR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   1
      Left            =   600
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2040
      Width           =   450
   End
   Begin VB.CommandButton cmdMEM 
      Caption         =   "MS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   2
      Left            =   600
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2520
      Width           =   450
   End
   Begin VB.CommandButton cmdMEM 
      Caption         =   "M+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   3
      Left            =   600
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3000
      Width           =   450
   End
   Begin VB.CommandButton cmdPercent 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2040
      Width           =   450
   End
   Begin VB.CommandButton cmdOneDivideX 
      Caption         =   "1/x"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2520
      Width           =   450
   End
   Begin VB.CommandButton btnFocusEqual 
      Caption         =   "btnTakeFocusFor ="
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton cmdSQRT 
      Caption         =   "sqrt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1590
      Width           =   450
   End
   Begin VB.Label lblOutput 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   600
      TabIndex        =   28
      Top             =   1080
      Width           =   2985
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MWARNET 2 - FREEWARE EDITION
'COPYRIGHT(C) 2007 MTechnologi Bali Indonesia
'Programed by A.A.Ngr.Manik Artawan
'e-mail : gungmanik@telkom.net
'---------------------------------------------
'THANK YOu FOR DOWNLOAD THIS SMALL APPLICATION
'---------------------------------------------


Dim nLastNum As Double, nResult As Double
Dim bOp As Boolean
Dim nOp As Integer
Dim bEqual As Boolean
Dim nMEM As Double, bMEM As Boolean
Dim bWasError As Boolean

Private Sub btnFocusC_Click()
Call cmdC_Click
End Sub
Private Sub btnFocusEqual_Click()
Call cmdEqual_Click
End Sub
Private Sub cmdC_Click()
bWasError = False
nLastNum = 0
nResult = 0
bOp = False
nOp = 0
bEqual = False
lblOutput.Caption = "0"

End Sub
Private Sub cmdCE_Click()
If bWasError Then
    Call cmdC_Click
End If
bWasError = False
lblOutput.Caption = 0
nLastNum = 0
End Sub
Private Sub cmdDot_Click()
If bWasError Then
    Beep
    Exit Sub
End If
If bOp = True Then
    lblOutput.Caption = ""
    nLastNum = 0 '!'
End If
If InStr(lblOutput.Caption, ".") = 0 And lblOutput.Caption <> "" Then
    lblOutput.Caption = lblOutput.Caption & "."
    ElseIf lblOutput.Caption = "" Then
        lblOutput.Caption = "0." & lblOutput.Caption
        Else
        Beep
End If
bOp = False
End Sub
Private Sub cmdEqual_Click()
On Error GoTo CheckIfOverFlowErr
If bWasError Then
    Beep
    Exit Sub
End If
If bOp = True And bEqual = False Then
nLastNum = nResult
End If
bEqual = True
bOp = False
Call cmdOP_Click(nOp)
bEqual = True

Exit Sub

CheckIfOverFlowErr:

If Err.Number = 6 Then
    lblOutput.Caption = "Value is over max calculation limit."
    bWasError = True
End If

End Sub
Private Sub cmdMEM_Click(Index As Integer)

If bWasError Then
    Beep
    Exit Sub
End If

Select Case Index


    Case 0 'MC
        
       nMEM = 0
       lblMem.Caption = ""
        
    Case 1 'MR
        
            lblOutput.Caption = nMEM
        
    Case 2 'MS
        If bOp Then  'If the last input was an operator
            Exit Sub 'Do nothing
        End If
        
        If CDbl(lblOutput.Caption) <> 0 Then
            nMEM = CDbl(lblOutput.Caption)
            lblMem = "M"
        End If
        
    Case 3 'M+
        If bOp Then  'If the last input was an operator
            Exit Sub 'Do nothing
        End If
    
        If CDbl(lblOutput.Caption) <> 0 Then
        
                nMEM = nMEM + CDbl(lblOutput.Caption)
                    lblMem.Caption = "M"
                    
            End If
            
            
End Select

bMEM = True

End Sub
Private Sub cmdNum_Click(Index As Integer)

If bWasError Then
    Beep
    Exit Sub
End If

If bOp Or bMEM Then
    lblOutput.Caption = ""
    ElseIf lblOutput.Caption = "0" Then
        lblOutput.Caption = ""
End If

If bEqual Then
            nLastNum = 0
            nResult = 0
Else
    nLastNum = 0
End If




lblOutput.Caption = lblOutput.Caption & cmdNum(Index).Caption

nLastNum = CDbl(lblOutput.Caption)
bOp = False
bMEM = False
bEqual = False

btnFocusEqual.SetFocus

End Sub

Private Sub cmdOneDivideX_Click()

If bWasError Then
    Beep
    Exit Sub
End If

nLastNum = CDbl(lblOutput.Caption)

If nLastNum = 0 Then
    lblOutput.Caption = "Cannot divide by zero."
    bWasError = True
    Exit Sub
End If

nLastNum = 1 / nLastNum

bOp = True
bEqual = True

lblOutput.Caption = nLastNum

CheckPlusMinusDot (True)

End Sub

Private Sub cmdOP_Click(Index As Integer)

On Error GoTo CheckIfOverFlowErr 'Check if we passed the Max Double Var Value's


If bWasError Then
    Beep
    Exit Sub
End If

If bOp = True And bEqual = False Then 'If the user has dbl clicked on an operator
 nOp = Index 'Remember the last operatore
 Exit Sub 'And..Exit sub
    ElseIf bEqual = True And bOp = True Then 'A MAJOR exeption! - If the user has pressed equal but Before it , pressed on an Operator (i.e - "3","+","=" ...)
        If nOp = 1 Or nOp = 2 Then 'So if the operator was "+ or - " ,
            nLastNum = 0 'Reset last num because we Don't Want to calculate Twice(First when OP was pressed and Second time when Equal was pressed)
                Else
                    nLastNum = 1 'If the last operator was "*" or "\" DO th Same(i.e - reset the last number) BUT ,don't put a Zero on it! because it will cause an Error when_
                    'the Next Calculation will take place(i.e - (38 * 0)+ 1=1 ->a wronge calculation   BUT  (38+0)+1=39 ->a correct calculation)
        End If
End If


If nOp = 0 Then
    nResult = CDbl(lblOutput.Caption)
End If

Select Case nOp

    Case 1 '+
        nResult = nResult + nLastNum
    Case 2 '-
        nResult = nResult - nLastNum
    Case 3 '*
        nResult = nResult * nLastNum
    Case 4 '/
        
        If nLastNum = 0 Then
            lblOutput.Caption = "Cannot divide by zero."
            bWasError = True
            Exit Sub
            Else
                nResult = nResult / nLastNum
        End If
        
End Select

nOp = Index
bOp = True
bEqual = False
lblOutput.Caption = nResult

If Left$(lblOutput.Caption, 1) = "." Then
    lblOutput.Caption = "0" & nResult
End If

btnFocusEqual.SetFocus

Exit Sub
CheckIfOverFlowErr:

If Err.Number = 6 Then
    lblOutput.Caption = "Value is over max calculation limit."
    bWasError = True
End If

End Sub

Private Sub cmdPercent_Click()

If bWasError Then
    Beep
    Exit Sub
End If

nLastNum = nResult * (CDbl(lblOutput.Caption) / 100)
lblOutput.Caption = nLastNum

End Sub

Private Sub cmdPlusMinus_Click()

If bWasError Then
    Beep
    Exit Sub
End If

If bOp And Not bEqual Then

    lblOutput.Caption = "0"
    Else
        lblOutput.Caption = CDbl(lblOutput.Caption) * (-1)
End If


CheckPlusMinusDot (True)

End Sub
Private Sub CheckPlusMinusDot(bAfterCalculation As Boolean)

lblOutput.Caption = Replace$(lblOutput.Caption, "-.", "-0.")

If Left$(lblOutput.Caption, 1) = "." Then
    lblOutput.Caption = Replace$(lblOutput.Caption, ".", "0.")
End If

If bAfterCalculation Then
    nLastNum = CDbl(lblOutput.Caption)
End If

End Sub
Private Sub CmdBS_Click()

If bWasError Then
    Beep
    Exit Sub
End If

If bEqual Or bMEM Or bOp Then
    Beep
    Exit Sub
End If

Static nBSCount As Integer
If (Len(lblOutput.Caption) > 1 And CDbl(lblOutput.Caption) > 0) Or (CDbl(lblOutput.Caption) < 0 And Len(lblOutput.Caption) > 2) Then
    lblOutput.Caption = Left$(lblOutput.Caption, Len(lblOutput.Caption) - 1)
    nBSCount = nBSCount + 1
    Else
    Beep
    lblOutput.Caption = 0
End If
nLastNum = CDbl(lblOutput.Caption)
End Sub
Private Sub cmdSQRT_Click()
If bWasError Then
    Beep
    Exit Sub
End If
nLastNum = CDbl(lblOutput.Caption)
If nLastNum < 0 Then
    lblOutput.Caption = "Invalid input for function."
    bWasError = True
    nLastNum = 0
    Exit Sub
    Else
        nLastNum = nLastNum ^ 0.5
        lblOutput.Caption = nLastNum
        bEqual = True
        bOp = True
End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 8 'Backspace
        Call CmdBS_Click
    Case 27 'Escape
        Call cmdC_Click
    Case 46 'Del
        Call cmdCE_Click
End Select
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 48 To 57 'Nums
        Call cmdNum_Click(Chr$(KeyAscii))
    Case 47
        Call cmdOP_Click(4)
    Case 42 '*
        Call cmdOP_Click(3)
    Case 45 '-
        Call cmdOP_Click(2)
    Case 43 '+
        Call cmdOP_Click(1)
    Case 46 'Dot
        Call cmdDot_Click
    Case 37 'Percent
        Call cmdPercent_Click
    End Select
End Sub

Private Sub lblOutput_Change()
If Len(lblOutput.Caption) > 36 Then 'Limit of input is - 33 numbers max
    Beep
    lblOutput.Caption = Left$(lblOutput.Caption, 36)
    End If
CheckPlusMinusDot (False)
End Sub
