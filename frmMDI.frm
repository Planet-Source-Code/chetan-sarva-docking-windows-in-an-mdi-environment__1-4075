VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9960
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   1500
      Left            =   0
      ScaleHeight     =   1470
      ScaleWidth      =   9930
      TabIndex        =   0
      Top             =   5385
      Visible         =   0   'False
      Width           =   9960
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal H%, ByVal hb%, ByVal X%, ByVal Y%, ByVal Cx%, ByVal Cy%, ByVal f%) As Integer
Private Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long)
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const WS_THICKFRAME = &H40000
Private Const GWL_STYLE = (-16)
Private Const SWP_DEFERERASE = &H2000
Private Const SWP_DRAWFRAME = &H20
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOCOPYBITS = &H100
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOREPOSITION = &H200
Private Const SWP_NOSENDCHANGING = &H400
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_FLAGS = SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME

Private Sub MDIForm_Load()

    Form2.Show
    Form1.Show
    Call SetWindowLong(Picture1.hWnd, GWL_STYLE, GetWindowLong(Picture1.hWnd, GWL_STYLE) Or WS_THICKFRAME)
    Call SetWindowPos(Picture1.hWnd, Me.hWnd, 0, 0, 0, 0, SWP_FLAGS)

End Sub

Private Sub Picture1_Resize()

    If Picture1.Visible = True Then Form1.Move -4 * Screen.TwipsPerPixelX, -2 * Screen.TwipsPerPixelY, Picture1.ScaleWidth + (8 * Screen.TwipsPerPixelX) - 35, Picture1.ScaleHeight + (8 * Screen.TwipsPerPixelY) - 35

End Sub
