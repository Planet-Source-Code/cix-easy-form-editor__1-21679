VERSION 5.00
Begin VB.Form Easyform 
   BorderStyle     =   0  'None
   Caption         =   "Easyform"
   ClientHeight    =   6825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   Picture         =   "STORFORM.frx":0000
   ScaleHeight     =   6825
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "EasyForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_SHOWINWINDOW = &H40
Private Const SWP_HIDEWINDOWS = &H80
Private Const SWP_NOCOPYBITS = &H100
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Private Const HWND_TOP = 0
Private Const HWND_BOTTOM = 1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInstrtAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Sub Form_Load()
'Show the from
Show
'Here i create the Shape of the form.
SetWindowRgn hwnd, CreateRoundRectRgn(4, 1, 491, 450, 30, 30), True
'Here i set the form on top most
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Here i move the from when the mouse is down.
ReleaseCapture
SendMessage Me.hwnd, &HA1, 2, 0&
End Sub
