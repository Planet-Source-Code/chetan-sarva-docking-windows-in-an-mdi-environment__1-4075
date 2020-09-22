VERSION 5.00
Object = "{307C5043-76B3-11CE-BF00-0080AD0EF894}#1.0#0"; "MSGHOO32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   2580
   ClientLeft      =   840
   ClientTop       =   930
   ClientWidth     =   3360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   3360
   ShowInTaskbar   =   0   'False
   Begin MsghookLib.Msghook MsgHook 
      Left            =   2820
      Top             =   2040
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Move me around towards the edges of the forms.. Pretty neat huh?"
      Height          =   915
      Left            =   300
      TabIndex        =   0
      Top             =   240
      Width           =   2715
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Chetan Sarva
' csarva@ic.sunysb.edu
'
' This code is a variation of the code posted by
' Steve of http://www.vbtutor.com/ on Planet
' Source Code. This code requires the
' MsgHook 32 OCX available here:
' http://www.mvps.org/vb/code/msghook.zip
' or possibly from somewhere on the Mabry
' site as well.
'
' If you use this code in any way, shape or
' form, please include all the above lines
' somewhere at the top of your application.
' It is only fair to give credit where credit is due.

' ####################################
' Declared Functions, Constants, and Types
' ####################################

'API Declares
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As Rect, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As Rect) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As Any) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As Rect, ByVal lLeft As Long, ByVal lTop As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetROP2 Lib "gdi32" (ByVal hDC As Long, ByVal nDrawMode As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal H%, ByVal hb%, ByVal x%, ByVal y%, ByVal Cx%, ByVal Cy%, ByVal f%) As Integer
Private Declare Function SetWindowWord Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

' Constants

Private Const WM_NCACTIVATE = &H86
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOPMOST = -1
Private Const WM_SYSCOMMAND = &H112
Private Const VK_LBUTTON = &H1
Private Const PS_SOLID = 0
Private Const R2_NOTXORPEN = 10
Private Const BLACK_PEN = 7

' Types

Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

' #############
' User Variables
' #############

'Public variables used elsewhere to set values
' for this form's position and size.
Dim lFloatingWidth As Long
Dim lFloatingHeight As Long
Dim lFloatingLeft As Long
Dim lFloatingTop As Long

Dim bMoving As Boolean

'Private variables used to track moving/sizing etc.
Public bDocked As Boolean
Public lDockedWidth As Long
Public lDockedHeight As Long

Dim fLeft As Long ' Hold the form's x coordinate for the Form_Moved event
Dim fWidth As Long ' Hold the form width for the Form_Moved event
Dim fHeight As Long ' Hold the form height for the Form_Moved event

Dim dropZone As Integer ' Size of the drop area

Dim TitleBarHeight As Integer ' Hold the height of the titlebar of our mdi form
Dim dockParent As Long ' Hold the hwnd of the mdi parent window
'

Sub Calc_Bottom(ptx As Long)

    fLeft = 0
    fWidth = MDIForm1.ScaleWidth / Screen.TwipsPerPixelX
    fHeight = dropZone / Screen.TwipsPerPixelY
            
End Sub

Sub Calc_Default()

    fLeft = 0
    fWidth = lFloatingWidth / Screen.TwipsPerPixelX
    fHeight = lFloatingHeight / Screen.TwipsPerPixelY
            
End Sub

Sub Calc_Left(ptx As Long)

    fLeft = ptx - (dropZone / Screen.TwipsPerPixelX) / 2
    fWidth = dropZone / Screen.TwipsPerPixelX
    fHeight = MDIForm1.ScaleHeight / Screen.TwipsPerPixelY
            
End Sub

Sub Calc_Right(ptx As Long)

    fLeft = ptx - (dropZone / Screen.TwipsPerPixelX) / 2
    fWidth = dropZone / Screen.TwipsPerPixelX
    fHeight = MDIForm1.ScaleHeight / Screen.TwipsPerPixelY
            
End Sub

Sub Calc_Top(ptx As Long)

    fLeft = 0
    fWidth = MDIForm1.ScaleWidth / Screen.TwipsPerPixelX
    fHeight = dropZone / Screen.TwipsPerPixelY
            
End Sub

Sub Dock_Bottom()

    MDIForm1.Picture1.Align = 2
    MDIForm1.Picture1.Height = dropZone
    lDockedWidth = MDIForm1.Picture1.ScaleWidth + (8 * Screen.TwipsPerPixelX)
    lDockedHeight = MDIForm1.Picture1.ScaleHeight + (8 * Screen.TwipsPerPixelY)
    bDocked = True
    Call SetParent(Me.hWnd, MDIForm1!Picture1.hWnd)
    Me.Move -4 * Screen.TwipsPerPixelX, -2 * Screen.TwipsPerPixelY, lDockedWidth - 35, lDockedHeight - 35
    MDIForm1!Picture1.Visible = True
    
End Sub

Sub Dock_Left()

    MDIForm1.Picture1.Align = 3
    MDIForm1.Picture1.Width = dropZone
    lDockedWidth = MDIForm1.Picture1.ScaleWidth + (8 * Screen.TwipsPerPixelX)
    lDockedHeight = MDIForm1.Picture1.ScaleHeight + (8 * Screen.TwipsPerPixelY)
    bDocked = True
    Call SetParent(Me.hWnd, MDIForm1!Picture1.hWnd)
    Me.Move -4 * Screen.TwipsPerPixelX, -2 * Screen.TwipsPerPixelY, lDockedWidth - 35, lDockedHeight - 35
    MDIForm1!Picture1.Visible = True
    
End Sub

Sub Dock_Right()

    MDIForm1.Picture1.Align = 4
    MDIForm1.Picture1.Width = dropZone
    lDockedWidth = MDIForm1.Picture1.ScaleWidth + (8 * Screen.TwipsPerPixelX)
    lDockedHeight = MDIForm1.Picture1.ScaleHeight + (8 * Screen.TwipsPerPixelY)
    bDocked = True
    Call SetParent(Me.hWnd, MDIForm1!Picture1.hWnd)
    Me.Move -4 * Screen.TwipsPerPixelX, -2 * Screen.TwipsPerPixelY, lDockedWidth - 35, lDockedHeight - 35
    MDIForm1!Picture1.Visible = True
    
End Sub

Sub Dock_Top()

    MDIForm1.Picture1.Align = 1
    MDIForm1.Picture1.Height = dropZone
    lDockedWidth = MDIForm1.Picture1.ScaleWidth + (8 * Screen.TwipsPerPixelX)
    lDockedHeight = MDIForm1.Picture1.ScaleHeight + (8 * Screen.TwipsPerPixelY)
    bDocked = True
    Call SetParent(Me.hWnd, MDIForm1!Picture1.hWnd)
    Me.Move -4 * Screen.TwipsPerPixelX, -2 * Screen.TwipsPerPixelY, lDockedWidth - 35, lDockedHeight - 35
    MDIForm1!Picture1.Visible = True
    
End Sub


Private Sub Form_Dropped(ptx As Long, pty As Long)
    
    Dim formRect As Rect
    Dim mdiRect As Rect
    Dim picRect As Rect
    Dim leftDock As Rect
    Dim rightDock As Rect
    Dim topDock As Rect
    Dim botDock As Rect

    'If over Picture1 on MDIForm1 which we are using as a Dock, set parent
    'of this form to Picture1, and position it at -4,-4 pixels, otherwise
    'set this Form's parent to the desktop and postion it at Left,Top
    'We dont need to size the form, as the DragForm control will have done
    'this for us.
    'For the purposes of this example, we only dock if the top left corner
    'of this form is within the area bounded by Picture1
    
    ' Get the screen based coordinates of our MDI Form
    GetWindowRect MDIForm1.hWnd, mdiRect
    GetWindowRect MDIForm1.Picture1.hWnd, picRect
    
    ' Set up the drop zone regions. These will be used for
    ' check to see if the form is to be docked or not.
    
    With leftDock
        .Left = mdiRect.Left + 4
        .Top = mdiRect.Top + TitleBarHeight '(mdiRect.Bottom - mdiRect.Top - MDIForm1.ScaleHeight \ Screen.TwipsPerPixelY) - 8
        .Right = .Left + dropZone \ Screen.TwipsPerPixelX
        .Bottom = .Top + MDIForm1.ScaleHeight \ Screen.TwipsPerPixelY + 4
    End With
    
    With rightDock
        .Left = mdiRect.Right - dropZone \ Screen.TwipsPerPixelX - 4
        .Top = mdiRect.Top + TitleBarHeight '(mdiRect.Bottom - mdiRect.Top - MDIForm1.ScaleHeight \ Screen.TwipsPerPixelY) - 8
        .Right = mdiRect.Right - 4
        .Bottom = .Top + MDIForm1.ScaleHeight \ Screen.TwipsPerPixelY + 4
    End With
    
    With topDock
        .Left = mdiRect.Left + dropZone \ Screen.TwipsPerPixelX + 4 + 1
        .Top = mdiRect.Top + TitleBarHeight '(mdiRect.Bottom - mdiRect.Top - MDIForm1.ScaleHeight \ Screen.TwipsPerPixelY) - 8
        .Right = rightDock.Right - dropZone \ Screen.TwipsPerPixelX
        .Bottom = .Top + dropZone \ Screen.TwipsPerPixelX
    End With
    
    With botDock
        .Left = mdiRect.Left + dropZone \ Screen.TwipsPerPixelX + 4 + 1
        .Top = mdiRect.Bottom - dropZone \ Screen.TwipsPerPixelX - 4
        .Right = rightDock.Right - dropZone \ Screen.TwipsPerPixelX
        .Bottom = mdiRect.Bottom - 4
    End With
    
    'See if the top/left corner of this form is in Picture1's screen rectangle
    'As we have set RepositionForm to false, we are responsible for positioning the form
    If (Not bDocked) Then
        If PtInRect(leftDock, ptx, pty) Then
            Dock_Left
        ElseIf PtInRect(rightDock, ptx, pty) Then
            Dock_Right
        ElseIf PtInRect(topDock, ptx, pty) Then
            Dock_Top
        ElseIf PtInRect(botDock, ptx, pty) Then
            Dock_Bottom
        End If
        
    Else
        Select Case (MDIForm1.Picture1.Align)
            Case 3
                If (PtInRect(picRect, ptx, pty)) Then
                    Dock_Left
                Else
                
                    If PtInRect(rightDock, ptx, pty) Then
                        Dock_Right
                    ElseIf PtInRect(topDock, ptx, pty) Then
                        Dock_Top
                    ElseIf PtInRect(botDock, ptx, pty) Then
                        Dock_Bottom
                    Else
                        UnDock
                    End If
                    
                End If ' (PtInRect(picRect, ptx, pty))
        
            Case 4
                If (PtInRect(picRect, ptx, pty)) Then
                    Dock_Right
                Else
                                
                    If PtInRect(leftDock, ptx, pty) Then
                        Dock_Left
                    ElseIf PtInRect(topDock, ptx, pty) Then
                        Dock_Top
                    ElseIf PtInRect(botDock, ptx, pty) Then
                        Dock_Bottom
                    Else
                        UnDock
                    End If
                
                End If ' (PtInRect(picRect, ptx, pty))
                
            Case 1
                If (PtInRect(picRect, ptx, pty)) Then
                    Dock_Top
                Else
                                
                    If PtInRect(leftDock, ptx, pty) Then
                        Dock_Left
                    ElseIf PtInRect(rightDock, ptx, pty) Then
                        Dock_Right
                    ElseIf PtInRect(botDock, ptx, pty) Then
                        Dock_Bottom
                    Else
                        UnDock
                    End If
                
                End If ' (PtInRect(picRect, ptx, pty))
                
            Case 2
                If (PtInRect(picRect, ptx, pty)) Then
                    Dock_Bottom
                Else
                    
                    If PtInRect(leftDock, ptx, pty) Then
                        Dock_Left
                    ElseIf PtInRect(rightDock, ptx, pty) Then
                        Dock_Right
                    ElseIf PtInRect(topDock, ptx, pty) Then
                        Dock_Top
                    Else
                        UnDock
                    End If
                    
                End If ' (PtInRect(picRect, ptx, pty))
                
        End Select ' Case (MDIForm1.Picture1.Align)

    End If ' (Not bDocked
    ' Reset the moving flag and store the form dimensions
    bMoving = False
    StoreFormDimensions

End Sub

Private Sub StoreFormDimensions()

   'Store the height/width values
    If Not bMoving Then
        If bDocked Then
            lDockedWidth = Me.Width
            lDockedHeight = Me.Height
        Else
            lFloatingLeft = Me.Left
            lFloatingTop = Me.Top
            lFloatingWidth = Me.Width
            lFloatingHeight = Me.Height
        End If
    End If
    
End Sub


Private Sub DrawDragRectangle(ByVal x As Long, ByVal y As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal lWidth As Long)

    'Draw a rectangle using the Win32 API

    Dim hDC As Long
    Dim hPen As Long
    hPen = CreatePen(PS_SOLID, lWidth, &HE0E0E0)
    hDC = GetDC(0)
    Call SelectObject(hDC, hPen)
    Call SetROP2(hDC, R2_NOTXORPEN)
    Call Rectangle(hDC, x, y, X1, Y1)
    Call SelectObject(hDC, GetStockObject(BLACK_PEN))
    Call DeleteObject(hPen)
    Call SelectObject(hDC, hPen)
    Call ReleaseDC(0, hDC)
    
End Sub

Private Sub DrawRect(rct As Rect)

    With rct

        'Draw a rectangle using the Win32 API
    
        Dim hDC As Long
        Dim hPen As Long
        hPen = CreatePen(PS_SOLID, 3, &HE0E0E0)
        hDC = GetDC(0)
        Call SelectObject(hDC, hPen)
        Call SetROP2(hDC, R2_NOTXORPEN)
        Call Rectangle(hDC, .Left, .Top, .Right, .Bottom)
        Call SelectObject(hDC, GetStockObject(BLACK_PEN))
        Call DeleteObject(hPen)
        Call SelectObject(hDC, hPen)
        Call ReleaseDC(0, hDC)
    
    End With
    
End Sub


Private Sub Form_Moved(ptx As Long, pty As Long)

    Dim formRect As Rect
    Dim mdiRect As Rect
    Dim picRect As Rect
    Dim leftDock As Rect
    Dim rightDock As Rect
    Dim topDock As Rect
    Dim botDock As Rect
    
    'Set the moving flag so we dont store the wrong dimensions
    bMoving = True
    
    'If over Picture1 on MDIForm1 which we are using as a Dock, change the width to that of
    'Picture1, else change it to the 'floating width and height
    'For the purposes of this example, we only dock if the top left corner
    'of this form is within the area bounded by Picture1
    
    ' Get the screen based coordinates of our MDI Form
    GetWindowRect MDIForm1.hWnd, mdiRect
    
    ' Get the screen based coordinates of our PictureBox
    GetWindowRect MDIForm1.Picture1.hWnd, picRect
    
    ' Set up the drop zone regions. These will be used for
    ' check to see if the form is to be docked or not.
    
    With leftDock
        .Left = mdiRect.Left + 4
        .Top = mdiRect.Top + TitleBarHeight '(mdiRect.Bottom - mdiRect.Top - MDIForm1.ScaleHeight \ Screen.TwipsPerPixelY) - 8
        .Right = .Left + dropZone \ Screen.TwipsPerPixelX
        .Bottom = .Top + MDIForm1.ScaleHeight \ Screen.TwipsPerPixelY + 4
    End With
    
    With rightDock
        .Left = mdiRect.Right - dropZone \ Screen.TwipsPerPixelX - 4
        .Top = mdiRect.Top + TitleBarHeight '(mdiRect.Bottom - mdiRect.Top - MDIForm1.ScaleHeight \ Screen.TwipsPerPixelY) - 8
        .Right = mdiRect.Right - 4
        .Bottom = .Top + MDIForm1.ScaleHeight \ Screen.TwipsPerPixelY + 4
    End With
    
    With topDock
        .Left = mdiRect.Left + dropZone \ Screen.TwipsPerPixelX + 4 + 1
        .Top = mdiRect.Top + TitleBarHeight '(mdiRect.Bottom - mdiRect.Top - MDIForm1.ScaleHeight \ Screen.TwipsPerPixelY) - 8
        .Right = rightDock.Right - dropZone \ Screen.TwipsPerPixelX
        .Bottom = .Top + dropZone \ Screen.TwipsPerPixelX
    End With
    
    With botDock
        .Left = mdiRect.Left + dropZone \ Screen.TwipsPerPixelX + 4 + 1
        .Top = mdiRect.Bottom - dropZone \ Screen.TwipsPerPixelX - 4
        .Right = rightDock.Right - dropZone \ Screen.TwipsPerPixelX
        .Bottom = mdiRect.Bottom - 4
    End With

    'DrawRect leftDock
    'DrawRect rightDock
    'DrawRect topDock
    'DrawRect botDock
    
    'Debug.Print "a) "; mdiRect.Top; " <--> "; mdiRect.Bottom
    'Debug.Print "b) "; topDock.Top; " <-->"; topDock.Bottom
    
    'See if the top/left corner of this form is in Picture1's screen rectangle
    
    If (Not bDocked) Then
        If PtInRect(leftDock, ptx, pty) Then
            Calc_Left ptx
        ElseIf PtInRect(rightDock, ptx, pty) Then
            Calc_Right ptx
        ElseIf PtInRect(topDock, ptx, pty) Then
            Calc_Top ptx
        ElseIf PtInRect(botDock, ptx, pty) Then
            Calc_Bottom ptx
        Else
            Calc_Default
        End If
        
    Else
        Select Case (MDIForm1.Picture1.Align)
            Case 3
                If (PtInRect(picRect, ptx, pty)) Then
                    Calc_Left ptx
                Else
                
                    If PtInRect(rightDock, ptx, pty) Then
                        Calc_Right ptx
                    ElseIf PtInRect(topDock, ptx, pty) Then
                        Calc_Top ptx
                    ElseIf PtInRect(botDock, ptx, pty) Then
                        Calc_Bottom ptx
                    Else
                        Calc_Default
                    End If
                    
                End If ' (PtInRect(picRect, ptx, pty))
        
            Case 4
                If (PtInRect(picRect, ptx, pty)) Then
                    Calc_Right ptx
                Else
                                
                    If PtInRect(leftDock, ptx, pty) Then
                        Calc_Left ptx
                    ElseIf PtInRect(topDock, ptx, pty) Then
                        Calc_Top ptx
                    ElseIf PtInRect(botDock, ptx, pty) Then
                        Calc_Bottom ptx
                    Else
                        Calc_Default
                    End If
                
                End If ' (PtInRect(picRect, ptx, pty))
                
            Case 1
                If (PtInRect(picRect, ptx, pty)) Then
                    Calc_Top ptx
                Else
                                
                    If PtInRect(leftDock, ptx, pty) Then
                        Calc_Left ptx
                    ElseIf PtInRect(rightDock, ptx, pty) Then
                        Calc_Right ptx
                    ElseIf PtInRect(botDock, ptx, pty) Then
                        Calc_Bottom ptx
                    Else
                        Calc_Default
                    End If
                
                End If ' (PtInRect(picRect, ptx, pty))
                
            Case 2
                If (PtInRect(picRect, ptx, pty)) Then
                    Calc_Bottom ptx
                Else
                    
                    If PtInRect(leftDock, ptx, pty) Then
                        Calc_Left ptx
                    ElseIf PtInRect(rightDock, ptx, pty) Then
                        Calc_Right ptx
                    ElseIf PtInRect(topDock, ptx, pty) Then
                        Calc_Top ptx
                    Else
                        Calc_Default
                    End If
                    
                End If ' (PtInRect(picRect, ptx, pty))
                
        End Select ' Case (MDIForm1.Picture1.Align)

    End If ' (Not bDocked)


    
End Sub

Sub UnDock()

    ' If it was docked before, undock it
    Call SetParent(Me.hWnd, dockParent)
    Me.Visible = False
    bDocked = False
    MDIForm1!Picture1.Visible = False
    Me.Visible = True
    Call SendMessage(Me.hWnd, WM_NCACTIVATE, 1, 0)
        
End Sub

Private Sub Form_Load()

    Dim mdiRect As Rect
    
    ' Initialize the drop area to a default value
    dropZone = 1500
    
    ' Get the screen based coordinates of our MDI Form
    GetWindowRect MDIForm1.hWnd, mdiRect
    ' Calculate the height of the titlebar and store it in a
    ' variable. We'll need this for later when a form is docked.
    TitleBarHeight = (mdiRect.Bottom - mdiRect.Top - MDIForm1.ScaleHeight \ Screen.TwipsPerPixelY) - 8
    
    ' The hWnd of the window that is the actual parent
    ' of our form is different from MDIForm1.hWnd
    ' Get it, and store it in a variable for later use
    dockParent = GetParent(Me.hWnd)
    
    ' Initialize the positions/sizes of this form
    lFloatingLeft = Me.Left
    lFloatingTop = Me.Top
    lFloatingWidth = Me.Width
    lFloatingHeight = Me.Height
    
    ' Subclass this form
    With MsgHook
        .HwndHook = Me.hWnd
        .Message(WM_SYSCOMMAND) = True
    End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If bDocked Then
    bDocked = False
    SetParent Me.hWnd, MDIForm1.hWnd
End If

End Sub

Private Sub Form_Resize()

' Update the stored Values
 If Me.WindowState <> vbMinimized Then StoreFormDimensions
        
End Sub

Private Sub MsgHook_Message(ByVal msg As Long, ByVal wp As Long, ByVal lp As Long, result As Long)

'Debug.Print GetWinMsgStr(msg)
Debug.Print wp

Select Case msg
    Case WM_SYSCOMMAND
        
        ' User is dragging the form. Simulate it with API
        If (wp = 61458) Then
            
            ' Local variables
            Dim pt As POINTAPI ' Current mouse location
            Dim ptPrev As POINTAPI ' Previous mouse location
            Dim mdiRect As Rect ' MDI Rectangle
            Dim objRect As Rect ' Form Rectangle
            Dim DragRect As Rect ' New Rectangle
            Dim lBorderWidth As Long ' Width of the border of the form (normally 3)
            Dim lObjWidth As Long ' Width of our form
            Dim lObjHeight As Long ' Height of our form
            Dim lXOffset As Long ' X Offset from mouse location
            Dim lYOffset As Long ' Y Offset from mouse location
            Dim bMoved As Boolean ' Did the form move?
            
            ' Get the rectangles for our MDI and our Form
            GetWindowRect MDIForm1.hWnd, mdiRect
            GetWindowRect Me.hWnd, objRect
            
            ' Determine the height and width of our form
            lObjWidth = objRect.Right - objRect.Left
            lObjHeight = objRect.Bottom - objRect.Top
            
            ' Get the location of our cursor and...
            GetCursorPos pt
            ' ... store it
            ptPrev.x = pt.x
            ptPrev.y = pt.y
            
            ' Determine offsets
            lXOffset = pt.x - objRect.Left
            lYOffset = pt.y - objRect.Top
            
            ' Create inital rectangle for drawing form's "edges"
            With DragRect
                .Left = pt.x - lXOffset
                .Top = pt.y - lYOffset
                .Right = .Left + lObjWidth
                .Bottom = .Top + lObjHeight
            End With
            
            ' Width of the form's border - used for showing
            ' the user that form is being dragged
            lBorderWidth = 3
            ' Draw the rectangle on the screen
            DrawDragRectangle DragRect.Left, DragRect.Top, DragRect.Right, DragRect.Bottom, lBorderWidth
            
            ' Check if the form is being moved
            Do While GetKeyState(VK_LBUTTON) < 0
                GetCursorPos pt
                
                If (pt.x <> ptPrev.x Or pt.y <> ptPrev.y) Then
                   
                   If pt.x < mdiRect.Left + 6 Then
                        pt.x = mdiRect.Left + 6
                    ElseIf pt.x > mdiRect.Right - 6 Then
                        pt.x = mdiRect.Right - 6
                    End If
                                        
                    If pt.y < mdiRect.Top + TitleBarHeight + 3 Then
                        pt.y = mdiRect.Top + TitleBarHeight + 3
                    ElseIf pt.y > mdiRect.Bottom - 6 Then
                        pt.y = mdiRect.Bottom - 6
                    End If
                    
                    SetCursorPos pt.x, pt.y
                    
                    ptPrev.x = pt.x
                    ptPrev.y = pt.y
                    
                    ' Erase the previous drag rectangle, if it exists, by drawing on top of it
                    DrawDragRectangle DragRect.Left, DragRect.Top, DragRect.Right, DragRect.Bottom, lBorderWidth
                    
                    ' Fire the Form_Moved event
                    Call Form_Moved(pt.x, pt.y)
                    
                    ' Adjust the height/width
                    With DragRect
                        If fLeft > 0 Then .Left = fLeft Else .Left = pt.x - lXOffset
                        .Top = pt.y - lYOffset
                        .Right = .Left + fWidth
                        .Bottom = .Top + fHeight
                    End With
                    
                    ' Draw the rectagle again at it's new position and dimensions
                    DrawDragRectangle DragRect.Left, DragRect.Top, DragRect.Right, DragRect.Bottom, lBorderWidth
                    bMoved = True
                End If ' (pt.X <> ptPrev.X Or pt.Y <> ptPrev.Y)
                
                DoEvents
            Loop ' While GetKeyState(VK_LBUTTON) < 0
            
            ' Erase the previous drag rectangle if any
            DrawDragRectangle DragRect.Left, DragRect.Top, DragRect.Right, DragRect.Bottom, lBorderWidth
            
            ' User let go of the left mouse button.
            ' If it is in a new location, fire Form_Dropped
            If (bMoved) Then
                
                Call Form_Dropped(pt.x, pt.y)
                
                ' If the form isn't docked, move it to it's new location
                If Not bDocked Then MoveWindow Me.hWnd, DragRect.Left - mdiRect.Left - 6, DragRect.Top - mdiRect.Top + 6 - (mdiRect.Bottom - mdiRect.Top - MDIForm1.ScaleHeight \ Screen.TwipsPerPixelY), DragRect.Right - DragRect.Left, DragRect.Bottom - DragRect.Top, True
            
            End If ' (bMoved)
        
        ' User clicked the close button. First we have to
        ' set the parent back to the MDI or it will lock up.
        ElseIf wp = 61536 Then
            Call SetParent(Me.hWnd, dockParent)
            Me.Visible = False
            MDIForm1!Picture1.Visible = False
            Unload Me
            
        ' If we got another wparam, return control to the window
        Else
            result = MsgHook.InvokeWindowProc(msg, wp, lp)
            
        End If ' (wp = 61458)
        
    ' On all other messages, return control the window
    Case Else
        result = MsgHook.InvokeWindowProc(msg, wp, lp)
            
End Select
    
End Sub
