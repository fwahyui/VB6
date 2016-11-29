Attribute VB_Name = "modFuncChild"
Option Explicit

Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

' Rectangle
Private Type RECT
   Left As Long     ' Left of the rectangle
   Top As Long      ' Top of the rectangle
   Right As Long    ' Right of the rectangle
   Bottom As Long   ' Bottom of the rectangle
End Type


Public Sub LoadForm(ByRef CFrm As Form)
    
    Dim R As RECT
    
    CFrm.Visible = False
    CFrm.WindowState = vbNormal

    GetClientRect mdiMain.hWnd, R
    
    'set client size
    'right
    If mdiMain.b8SBC.Visible = True Then
        R.Right = R.Right - (mdiMain.b8SBC.Width / Screen.TwipsPerPixelX)
    End If
    'bottom
    R.Bottom = R.Bottom - ((mdiMain.b8CW.Height / Screen.TwipsPerPixelY) + (mdiMain.bgHeader.Height / Screen.TwipsPerPixelY)) - R.Top
    
    mdiMain.b8CW.LoadChildWindow mdiMain.hWnd, CFrm.hWnd, CFrm.Name, CFrm.Caption, R.Top, R.Left, R.Right, R.Bottom

    CFrm.Visible = True
    CFrm.Show
    CFrm.SetFocus
    
    ResizeMdiChildForm CFrm
End Sub

Public Sub ResizeMdiChildForm(ByRef CFrm As Form)

    Dim R As RECT
    
    GetClientRect mdiMain.hWnd, R
    
    'set client size
    'right
    If mdiMain.b8SBC.Visible = True Then
        R.Right = R.Right - (mdiMain.b8SBC.Width / Screen.TwipsPerPixelX)
    End If
    'bottom
    R.Bottom = R.Bottom - ((mdiMain.b8CW.Height / Screen.TwipsPerPixelY) + (mdiMain.bgHeader.Height / Screen.TwipsPerPixelY)) - R.Top

    mdiMain.b8CW.ResizeClientWin CFrm.hWnd, R.Top, R.Left, R.Right, R.Bottom

End Sub

Public Sub ActivateMDIChildForm(ByVal sFormName As String)
    
    Dim frm As Form
    
    For Each frm In Forms
        If frm.Name = sFormName Then
            'activate form
            ResizeMdiChildForm frm
            frm.Visible = True
            frm.Show
            frm.SetFocus
            'set tab active window
            mdiMain.b8CW.SetActiveWindow sFormName
            Exit For
        End If
    Next
    
    
    Set frm = Nothing
End Sub


