VERSION 5.00
Begin VB.Form frmControls 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   0  'None
   ClientHeight    =   4635
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   7695
   ControlBox      =   0   'False
   LinkTopic       =   "frmControls"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmControls.frx":0000
   ScaleHeight     =   4635
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3840
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1110
      Width           =   1590
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Unload Me"
      Height          =   435
      Left            =   3405
      TabIndex        =   1
      Top             =   2715
      Width           =   1290
   End
   Begin VB.Label Label3 
      Caption         =   "If popup menus are used:"
      Height          =   330
      Left            =   4050
      TabIndex        =   7
      Top             =   135
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.Label Label2 
      Caption         =   "Img is only used to help position controls.  It must be removed in the Form_Load event"
      Height          =   585
      Left            =   3915
      TabIndex        =   6
      Top             =   3720
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cooler GUIs are possible"
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Index           =   2
      Left            =   2955
      TabIndex        =   5
      Top             =   1935
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Using 2 Windows to Make 1"
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Index           =   1
      Left            =   2970
      TabIndex        =   4
      Top             =   1755
      Width           =   2175
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   1875
      Picture         =   "FrmControls.frx":8B48
      ToolTipText     =   "Change Opacity"
      Top             =   2445
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   1815
      Picture         =   "FrmControls.frx":B2EA
      ToolTipText     =   "Close"
      Top             =   3300
      Width           =   240
   End
   Begin VB.Label Label4 
      Caption         =   "ControlBox Property = False     BorderStyle = 0  Caption=Null"
      Height          =   630
      Left            =   4080
      TabIndex        =   3
      Top             =   525
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Win2K/XP/Vista Only"
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Index           =   0
      Left            =   2760
      TabIndex        =   0
      Top             =   855
      Width           =   1695
   End
   Begin VB.Menu mnuOpacity 
      Caption         =   "mnuOpacity"
      Visible         =   0   'False
      Begin VB.Menu mnuOpacitySub 
         Caption         =   "100% Opacity"
         Index           =   0
      End
      Begin VB.Menu mnuOpacitySub 
         Caption         =   "75% Opacity"
         Index           =   1
      End
      Begin VB.Menu mnuOpacitySub 
         Caption         =   "50% Opacity"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function UpdateLayeredWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal hdcDst As Long, ByRef pptDst As Any, ByRef psize As Any, ByVal hdcSrc As Long, ByRef pptSrc As Any, ByVal crKey As Long, ByRef pblend As Long, ByVal dwFlags As Long) As Long
Private Const AC_SRC_ALPHA As Long = &H1
Private Const AC_SRC_OVER As Long = &H0
Private Const ULW_ALPHA As Long = &H2

Private Sub Command1_Click()
    Unload frmBkg   ' unload
End Sub

Private Sub Form_Load()
    ' remove our picture we used for positioning controls
    Set Me.Picture = LoadPicture("")
End Sub


Private Sub Image1_Click()
    Unload frmBkg ' unload
End Sub

Private Sub Image2_Click()

    frmBkg.SetFocus
    PopupMenu mnuOpacity
    Me.SetFocus
    
End Sub

Private Sub mnuOpacitySub_Click(Index As Integer)
        
    Dim newOpacity As Long, lBlendFunc As Long
    Select Case Index
    Case 0: newOpacity = 255
    Case 1: newOpacity = 255 * 0.75
    Case 2: newOpacity = 255 \ 2
    End Select
    
    ' create a blend function.
    lBlendFunc = AC_SRC_OVER Or (newOpacity * &H10000) Or (AC_SRC_ALPHA * &H1000000)
    
    ' tell windows to draw our background from whenever it needs redrawing
    UpdateLayeredWindow frmBkg.hwnd, 0&, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, 0&, lBlendFunc, ULW_ALPHA


End Sub


