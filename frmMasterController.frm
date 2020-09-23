VERSION 5.00
Begin VB.Form frmMasterController 
   Caption         =   "ULW Tester"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMasterController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Const WM_NCACTIVATE As Long = &H86

Private Sub Form_Load()
    Move -280000, -280000
    Load frmBkg
    Show
    frmBkg.Show , Me
    frmControls.Show , frmBkg
    SendMessage Me.hwnd, WM_NCACTIVATE, 1, ByVal 0&
End Sub
