VERSION 5.00
Begin VB.UserControl SkinButton 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1275
   ScaleHeight     =   495
   ScaleWidth      =   1275
   Begin VB.PictureBox pic_Buttons 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   420
      Left            =   0
      Picture         =   "ctrl_SkinableButton.ctx":0000
      ScaleHeight     =   360
      ScaleWidth      =   4320
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   4380
   End
   Begin VB.PictureBox pic_Normal 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   1335
      TabIndex        =   1
      Top             =   0
      Width           =   1335
      Begin VB.Label lbl_Normal 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caption"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   660
      End
   End
   Begin VB.PictureBox pic_MouseMove 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   1335
      TabIndex        =   5
      Top             =   0
      Width           =   1335
      Begin VB.Label lbl_MouseMove 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caption"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   660
      End
   End
   Begin VB.PictureBox Pic_Down 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   1335
      TabIndex        =   3
      Top             =   0
      Width           =   1335
      Begin VB.Label lbl_down 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caption"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   660
      End
   End
End
Attribute VB_Name = "SkinButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Skin Button By Arbie Sarkissian

Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020

Const DefCaption = "Caption"
Const DefForeColor = 0
Const DefEnabled = 1

Dim v_sSkinPath As String
Dim v_sCaption As String
Dim v_oForeColor As OLE_COLOR
Dim v_bEnabled As Boolean

Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."

Public Sub LoadSkin()
    Dim v_lRtn As Long
    Dim v_iCenterImgFrequency As Integer
    Dim v_iLoop As Integer

    With UserControl
        
        .pic_Normal.Cls
        .pic_MouseMove.Cls
        .Pic_Down.Cls
        
        .lbl_Normal.Caption = ""
        .lbl_MouseMove.Caption = ""
        .lbl_down.Caption = ""
        
        .Pic_Down.Visible = True
        .pic_MouseMove.Visible = True
        .pic_Normal.Visible = True
        
        .pic_Normal.Width = .Width
        .pic_MouseMove.Width = .Width
        .Pic_Down.Width = .Width
        .pic_Normal.Height = 360
        .pic_MouseMove.Height = 360
        .Pic_Down.Height = 360
        
        v_lRtn = BitBlt(.pic_Normal.hDC, 0, 0, 15, 24, .pic_Buttons.hDC, 0, 0, SRCCOPY)
        v_lRtn = BitBlt(.pic_MouseMove.hDC, 0, 0, 15, 24, .pic_Buttons.hDC, 72, 0, SRCCOPY)
        v_lRtn = BitBlt(.Pic_Down.hDC, 0, 0, 15, 24, .pic_Buttons.hDC, 144, 0, SRCCOPY)
        
        v_iCenterImgFrequency = Abs((.Width / Screen.TwipsPerPixelX) / 15)
        If v_iCenterImgFrequency > 0 Then
            For v_iLoop = 1 To v_iCenterImgFrequency
                v_lRtn = BitBlt(.pic_Normal.hDC, v_iLoop * 15, 0, 15, 24, .pic_Buttons.hDC, 15, 0, SRCCOPY) 'normal
                v_lRtn = BitBlt(.pic_MouseMove.hDC, v_iLoop * 15, 0, 15, 24, .pic_Buttons.hDC, 83, 0, SRCCOPY) 'mouse move
                v_lRtn = BitBlt(.Pic_Down.hDC, v_iLoop * 15, 0, 15, 24, .pic_Buttons.hDC, 159, 0, SRCCOPY) 'mouse Down
            Next v_iLoop
        End If
        v_lRtn = BitBlt(.pic_Normal.hDC, (.Width / Screen.TwipsPerPixelX) - 16, 0, 16, 24, .pic_Buttons.hDC, 55, 0, SRCCOPY) 'normal
        v_lRtn = BitBlt(.pic_MouseMove.hDC, (.Width / Screen.TwipsPerPixelX) - 16, 0, 16, 24, .pic_Buttons.hDC, 128, 0, SRCCOPY) 'mouse move
        v_lRtn = BitBlt(.Pic_Down.hDC, (.Width / Screen.TwipsPerPixelX) - 16, 0, 16, 24, .pic_Buttons.hDC, 202, 0, SRCCOPY) 'mousedown
         
        Pic_Down.Refresh
        pic_MouseMove.Refresh
        pic_Normal.Refresh
        
        .lbl_Normal.Left = .Width / 2
        .lbl_Normal.Top = 60
        .lbl_MouseMove.Left = .Width / 2
        .lbl_MouseMove.Top = 60
        .lbl_down.Left = .Width / 2
        .lbl_down.Top = 60
        
        .lbl_Normal.Caption = v_sCaption
        .lbl_MouseMove.Caption = v_sCaption
        .lbl_down.Caption = v_sCaption
        .lbl_down.ForeColor = v_oForeColor
        .lbl_MouseMove.ForeColor = v_oForeColor
        .lbl_Normal.ForeColor = v_oForeColor
        Refresh
    End With
End Sub

Public Sub Refresh()
    Select Case True
    Case Pic_Down.Visible = True
      Pic_Down.Visible = False
    Case pic_MouseMove.Visible = True
      pic_MouseMove.Visible = False
    Case pic_Normal.Visible = False
      pic_Normal.Visible = True
    End Select
End Sub

Public Property Get SkinPath() As String
    SkinPath = v_sSkinPath
End Property

Public Property Let SkinPath(ByVal m_SkinPath As String)
    v_sSkinPath = m_SkinPath
    PropertyChanged "SkinPath"
End Property

Public Property Get Caption() As String
    Caption = v_sCaption
End Property

Public Property Let Caption(ByVal m_Caption As String)
    v_sCaption = m_Caption
    PropertyChanged "Caption"
    lbl_Normal.Caption = m_Caption
    lbl_MouseMove.Caption = m_Caption
    lbl_down.Caption = m_Caption
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = v_oForeColor
End Property

Public Property Let ForeColor(ByVal m_ForeColor As OLE_COLOR)
    v_oForeColor = m_ForeColor
    PropertyChanged "ForeColor"
    lbl_Normal.ForeColor = m_ForeColor
    lbl_MouseMove.ForeColor = m_ForeColor
    lbl_down.ForeColor = m_ForeColor
End Property

Public Property Get Enabled() As Boolean
    Enabled = v_bEnabled
End Property

Public Property Let Enabled(ByVal m_Enabled As Boolean)
    v_bEnabled = m_Enabled
    PropertyChanged "Enabled"
    
    lbl_Normal.Enabled = v_bEnabled
    pic_Normal.Enabled = v_bEnabled
    lbl_MouseMove.Enabled = v_bEnabled
    pic_MouseMove.Enabled = v_bEnabled
    lbl_down.Enabled = v_bEnabled
    Pic_Down.Enabled = v_bEnabled
End Property

Private Sub pic_Button_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lbl_Caption_MouseUp Button, Shift, X, Y
End Sub

Private Sub lbl_MouseMove_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  pic_MouseMove_MouseDown Button, Shift, X, Y
End Sub

Private Sub lbl_MouseMove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub lbl_MouseMove_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  pic_MouseMove_MouseUp Button, Shift, X, Y
End Sub

Private Sub lbl_Normal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  pic_Normal_MouseMove Button, Shift, X, Y
End Sub

Private Sub pic_MouseMove_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  pic_MouseMove.Visible = False
  Pic_Down.Visible = True
  RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub pic_MouseMove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub pic_MouseMove_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  pic_MouseMove.Visible = True
  Pic_Down.Visible = False
  RaiseEvent MouseUp(Button, Shift, X, Y)
  RaiseEvent Click
End Sub

Private Sub pic_Normal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  pic_Normal.Visible = False
  pic_MouseMove.Visible = True
  RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_InitProperties()
    v_sSkinPath = App.Path & "\Skins\Titanium"
    v_sCaption = DefCaption
    v_oForeColor = DefForeColor
    v_bEnabled = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    v_sSkinPath = PropBag.ReadProperty("SkinPath", App.Path & "\Skins\Titanium")
    v_sCaption = PropBag.ReadProperty("Caption", DefCaption)
    v_oForeColor = PropBag.ReadProperty("ForeColor", DefForeColor)
    Call LoadSkin
          
    v_bEnabled = PropBag.ReadProperty("Enabled", DefEnabled)
    If v_bEnabled = True Then
        Call Refresh
    Else
        UserControl.lbl_Normal.Enabled = False
        UserControl.pic_Normal.Enabled = False
    End If
End Sub

Private Sub UserControl_Resize()
    LoadSkin
    UserControl.Height = pic_Normal.Height
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("SkinPath", v_sSkinPath, App.Path & "\Skins\Titanium")
    Call PropBag.WriteProperty("Caption", v_sCaption, DefCaption)
    Call PropBag.WriteProperty("ForeColor", v_oForeColor, DefForeColor)
    Call PropBag.WriteProperty("Enabled", v_bEnabled, DefEnabled)
End Sub
