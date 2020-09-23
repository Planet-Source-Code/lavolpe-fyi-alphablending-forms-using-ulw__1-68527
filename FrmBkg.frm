VERSION 5.00
Begin VB.Form frmBkg 
   BorderStyle     =   0  'None
   ClientHeight    =   1080
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   1410
   ControlBox      =   0   'False
   LinkTopic       =   "frmBkg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   1410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmBkg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetCapture Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Const MOUSEEVENTF_LEFTUP As Long = &H4
Private Declare Sub mouse_event Lib "user32.dll" (ByVal dwFlags As Long, ByVal dX As Long, ByVal dY As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Function SetFocusAPI Lib "user32.dll" Alias "SetFocus" (ByVal hwnd As Long) As Long

' Open ULW_Readme.txt.  It is only a few paragraphs and may help understand what/why.

' Trying to offer a friend some advice on the UpdateLayeredWindow and SetLayeredWindowAttributes
' APIs, I found myself needing to understand it a bit more. Therefore, I whipped together
' a simple demo and thought it might be worth sharing.

' REQUIRES WINDOWS 2000, XP or VISTA


Private Const WS_EX_LAYERED As Long = &H80000
Private Const GWL_EXSTYLE As Long = -20
Private Const ULW_ALPHA As Long = &H2
Private Const ULW_COLORKEY As Long = &H1
Private Const WM_NCLBUTTONDOWN As Long = &HA1
Private Const HTCAPTION As Long = 2
Private Const AC_SRC_ALPHA As Long = &H1
Private Const AC_SRC_OVER As Long = &H0
Private Const GWL_STYLE As Long = -16
Private Const WS_BORDER As Long = &H800000

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type Size
    cX As Long
    cY As Long
End Type

Private Declare Function UpdateLayeredWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal hdcDst As Long, ByRef pptDst As Any, ByRef psize As Any, ByVal hdcSrc As Long, ByRef pptSrc As Any, ByVal crKey As Long, ByRef pblend As Long, ByVal dwFlags As Long) As Long
' modified above API parameters
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SetWindowRgn Lib "user32.dll" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private mButton As Integer          ' see mouse_move & mouse_up
Private mMousePoints As POINTAPI
Private cComposite As c32bppDIB


Private Sub Form_DblClick()
    Unload Me   ' offer another way to close the form
End Sub

Private Sub Form_Load()

    ' with this simple demo, the form must be borderless

    Dim cImage As c32bppDIB
    Dim lBlend As Long
    Dim srcPt As POINTAPI
    Dim srcSize As Size
    Dim lBlendFunc As Long


    ' this will be the class we hold the image in
    Set cComposite = New c32bppDIB
    cComposite.ManageOwnDC = True
    Set cImage = New c32bppDIB
    cComposite.LoadPicture_File App.Path & "\background.png"
    
    ' example of composing multiple images to form a single image
    ' Create a new class and load the next layer in it
    cImage.LoadPicture_File App.Path & "\over.png"
    ' now render that layer over the composite
    cImage.Render cComposite.LoadDIBinDC(True), 146, 19
    
    Set cImage = Nothing    ' not needed any longer
    
' // FYI - This is how I got the jpg I used for frmControls
' so that I can overlay controls on it.
    
'    Dim cGDI As cGDIPlus, aDummy() As Byte
'    Set cGDI = New cGDIPlus
'    cGDI.SaveToJPG App.Path & "\_testBkg.jpg", aDummy(), cComposite
    
    ' the image will be the width/height of the background form and the controls form
    srcSize.cX = cComposite.Width
    srcSize.cY = cComposite.Height
    
    ' apply the layered attribute
    SetWindowLong Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    
    ' create a blend function. Change 180 below to whatever opacity you want
    lBlendFunc = AC_SRC_OVER Or (180 * &H10000) Or (AC_SRC_ALPHA * &H1000000)
    
    ' tell windows to draw our background form whenever it needs redrawing
    UpdateLayeredWindow Me.hwnd, 0&, ByVal 0&, srcSize, cComposite.LoadDIBinDC(True), srcPt, 0&, lBlendFunc, ULW_ALPHA
    
    ' load the controls form. It has a awful background color on purpose
    Load frmControls
    ' set its layered attribute
    SetWindowLong frmControls.hwnd, GWL_EXSTYLE, GetWindowLong(frmControls.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    ' now tell windows that the form is transparent wherever the background color exists
    SetLayeredWindowAttributes frmControls.hwnd, frmControls.BackColor, 0&, ULW_COLORKEY

    ' ensure the controls form does not have a border around it. Windows may attempt to do so
    SetWindowLong frmControls.hwnd, GWL_STYLE, GetWindowLong(frmControls.hwnd, GWL_STYLE) And Not WS_BORDER
    
    ' position the controls form over the background form
    frmControls.Move Me.Left, Me.Top, srcSize.cX * Screen.TwipsPerPixelX, srcSize.cY * Screen.TwipsPerPixelY

    Show
    frmControls.Show , Me

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' moving the controls form when the background form is dragged
    ' See the ULW_Readme.txt file for more
    
    
    ' baby hack. without subclassing, we don't want a simple click on the form
    ' to trigger the move routine below. We want to wait until the mouse is
    ' actually moved with the left button down. That is what the mMousePoints
    ' and mButton are used for. The mButton is used additionally so that
    ' should you hit escape while dragging, the mouse doesn't trigger another
    ' move. It will unfortunately, because this form never gets the mouse up
    ' event after moving is complete. Change mButton to Button and try it.
    
    If mButton = vbLeftButton Then
        If mMousePoints.X <> X And mMousePoints.Y <> Y Then
            ReleaseCapture
            ' draw the composite over the controls form, alphablending so the awful bkg color shows thru
            frmControls.AutoRedraw = True
            cComposite.Render frmControls.hDC
            frmControls.Refresh   ' ensure windowless controls are refreshed
            
            ' hide our composite form & move the controls form
            frmBkg.Visible = False
            
            SendMessage frmControls.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&
            ' done moving when we get here
            
            ' position our bkg form and show it
            frmBkg.Move frmControls.Left, frmControls.Top
            frmBkg.Visible = True
            
            ' erase the controls form image and refresh so it is transparent again
            frmControls.Cls
            frmControls.AutoRedraw = False
            
            mButton = 0
        End If
    End If
    
    
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        mButton = vbLeftButton
        mMousePoints.X = X
        mMousePoints.Y = Y
    End If
End Sub
