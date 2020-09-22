VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EasyForm"
   ClientHeight    =   1050
   ClientLeft      =   3240
   ClientTop       =   2715
   ClientWidth     =   8235
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   8235
   Begin VB.CommandButton Command1 
      Caption         =   "&Exit"
      Height          =   330
      Left            =   6960
      TabIndex        =   9
      Top             =   720
      Width           =   1215
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "AutoSize Form"
      Height          =   255
      Left            =   3960
      TabIndex        =   8
      Top             =   720
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   7800
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Easy Form"
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Set form onTop"
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   720
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Make form Moveable"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   255
      Left            =   5760
      TabIndex        =   5
      Top             =   480
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   50
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   60
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   80
      Width           =   1095
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Left            =   3360
      TabIndex        =   0
      Top             =   480
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   50
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   8400
      Top             =   1320
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Form Width :"
      Height          =   255
      Left            =   5880
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Form Hieght : "
      Height          =   255
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image Image4 
      Height          =   540
      Index           =   2
      Left            =   2520
      Picture         =   "Form1.frx":0442
      Top             =   1680
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image Image3 
      Height          =   540
      Index           =   2
      Left            =   1680
      Picture         =   "Form1.frx":19E6
      Top             =   1680
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image Image2 
      Height          =   540
      Index           =   2
      Left            =   840
      Picture         =   "Form1.frx":2F8A
      Top             =   1680
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image Image1 
      Height          =   540
      Index           =   2
      Left            =   0
      Picture         =   "Form1.frx":452E
      Top             =   1680
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image Image4 
      Height          =   540
      Index           =   1
      Left            =   2520
      Picture         =   "Form1.frx":5AD2
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image Image3 
      Height          =   540
      Index           =   1
      Left            =   1680
      Picture         =   "Form1.frx":7076
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image Image2 
      Height          =   540
      Index           =   1
      Left            =   840
      Picture         =   "Form1.frx":861A
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image Image1 
      Height          =   540
      Index           =   1
      Left            =   0
      Picture         =   "Form1.frx":9BBE
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image Image4 
      Height          =   540
      Index           =   0
      Left            =   2520
      Picture         =   "Form1.frx":B162
      ToolTipText     =   "Draw a Circle"
      Top             =   0
      Width           =   750
   End
   Begin VB.Image Image3 
      Height          =   540
      Index           =   0
      Left            =   1680
      Picture         =   "Form1.frx":C706
      ToolTipText     =   "Draw a RectAngle"
      Top             =   0
      Width           =   750
   End
   Begin VB.Image Image2 
      Height          =   540
      Index           =   0
      Left            =   840
      Picture         =   "Form1.frx":DCAA
      Top             =   0
      Width           =   750
   End
   Begin VB.Image Image1 
      Height          =   540
      Index           =   0
      Left            =   0
      Picture         =   "Form1.frx":F24E
      Top             =   0
      Width           =   750
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Oki First of all SORRY if my english is bad i am danish.
' I build this couse somebody asked me too
' I seen somebody try to Make money on a program like this.
' So i build this one just to same my self some time, and you some cash.
' DONT tell me! I know i should comment my work better, but then agian.
' I didt have the time, and if you like you can keep it and comment the part you like your self. ;-)
' Just credit me if you write or use my code..
' Well thats it... and Thank you for trying it out.
' Be so kind as to rate my code if you like it.

Option Explicit

Private Sub Command1_Click()
' Jubii gess what is exit the program. :-)
 Unload Form2
 Unload Me
End Sub

Private Sub Form_Load()
' Just scaleing
 Form2.Show
 Form2.Left = Me.Left
 Form2.Top = Me.Top + Me.Height
 Slider1.Max = VB.Screen.Height / 15
 Slider2.Max = VB.Screen.Width / 15
 Form2.ScaleMode = vbPixels
 Form2.Width = Me.Width
 Form2.Height = Me.Height * 2
 Slider1.Value = Form2.ScaleHeight
 Slider2.Value = Form2.ScaleWidth
 Text1.Text = Slider1.Value
 Text2.Text = Slider2.Value
' Just a boolean i use to find out what kinda shape to draw.
' If its True then its a Ellipse if not a Round Rectangle
 Form2.xCirkel = False
End Sub

Private Sub Form_Resize()
 Form2.Left = Me.Left
 Form2.Top = Me.Top + Me.Height
End Sub

Private Sub Image1_Click(Index As Integer)
 Form2.Width = Me.Width
 Form2.Height = Me.Height * 2
 Slider1.Value = Form2.ScaleHeight
 Slider2.Value = Form2.ScaleWidth
 Form2.Picture1.Cls
End Sub

Private Sub Image2_Click(Index As Integer)
 Dialog1.DialogTitle = "Save vbform"
Dialog1.Filter = "VB FORMS [*.frm]|*.frm"
' If you dont know this is the way to hide the Readonly box and
' Make it ask before overwriteing any files.
Dialog1.Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt
Dialog1.ShowSave
On Error GoTo Err
Open Dialog1.FileName For Output As 1
SavePicture Form2.Picture1.Image, Mid(Dialog1.FileName, 1, Len(Dialog1.FileName) - 3) & "bmp"


' Oki all i basicly do is write all code here to for the form
' You know IF you checked this box write this and that...

Print #1, "VERSION 5.00"
Print #1, "Begin VB.Form Easyform"
Print #1, "   BorderStyle     =   0    'None"
Print #1, "   Caption         =   " & """" & "Easyform" & """"
Print #1, "   ClientHeight    =   " & Form2.Height
Print #1, "   ClientLeft      =   0"
Print #1, "   ClientTop       =   0"
Print #1, "   ClientWidth     =   " & Form2.Width
Print #1, "   LinkTopic       =   " & """" & "Form1" & """"

'Form2.ScaleHeight * 15 is my cheap way on turning pixels into twips.

Print #1, "   ScaleHeight     =   " & Form2.ScaleHeight * 15
Print #1, "   ScaleWidth      =   " & Form2.ScaleWidth * 15
Print #1, "   ShowInTaskbar   =   0    'False"
Print #1, "   StartUpPosition =   2    'Center Of Screen"
Print #1, "End"
Print #1, "Attribute VB_Name = " & """" & "EasyForm" & """"
Print #1, "Attribute VB_GlobalNameSpace = False"
Print #1, "Attribute VB_Creatable = False"
Print #1, "Attribute VB_PredeclaredId = True"
Print #1, "Attribute VB_Exposed = False"
Print #1, ""

If Check2.Value = Checked Then
 Print #1, "Private Const SWP_NOSIZE = &H1"
 Print #1, "Private Const SWP_NOMOVE = &H2"
 Print #1, "Private Const SWP_NOZORDER = &H4"
 Print #1, "Private Const SWP_NOREDRAW = &H8"
 Print #1, "Private Const SWP_NOACTIVATE = &H10"
 Print #1, "Private Const SWP_FRAMECHANGED = &H20"
 Print #1, "Private Const SWP_SHOWINWINDOW = &H40"
 Print #1, "Private Const SWP_HIDEWINDOWS = &H80"
 Print #1, "Private Const SWP_NOCOPYBITS = &H100"
 Print #1, "Private Const SWP_NOOWNERZORDER = &H200"
 Print #1, "Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED"
 Print #1, "Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER"
 Print #1, "Private Const HWND_TOP = 0"
 Print #1, "Private Const HWND_BOTTOM = 1"
 Print #1, "Private Const HWND_TOPMOST = -1"
 Print #1, "Private Const HWND_NOTOPMOST = -2"
 Print #1, "Private Declare Function SetWindowPos Lib " & """" & "user32" & """" & " (ByVal hwnd As Long, ByVal hWndInstrtAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long"
End If

Print #1, "Private Declare Function SetWindowRgn Lib " & """" & "user32" & """" & " (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long"
If Check1 Then
 Print #1, "Private Declare Function SendMessage Lib " & """" & "user32" & """" & " Alias " & """" & "SendMessageA" & """" & " (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long"
 Print #1, "Private Declare Function ReleaseCapture Lib " & """" & "user32" & """" & " () As Long"
End If

If Form2.xCirkel Then
 Print #1, "Private Declare Function CreateEllipticRgn Lib " & """" & "gdi32" & """" & " (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long"
Else
 Print #1, "Private Declare Function CreateRoundRectRgn Lib " & """" & "gdi32" & """" & " (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long"
End If
 Print #1, "Private Sub Form_Load()"
 Print #1, "'Show the from"
 Print #1, "Show"
 Print #1, "'Here i create the Shape of the form."
If Form2.xCirkel Then
 Print #1, "SetWindowRgn hWnd, CreateEllipticRgn(" & FocusRec.Left & ", " & FocusRec.Top & ", " & FocusRec.Right & ", " & FocusRec.Bottom & "), True"
Else
 Print #1, "SetWindowRgn hWnd, CreateRoundRectRgn(" & FocusRec.Left & ", " & FocusRec.Top & ", " & FocusRec.Right & ", " & FocusRec.Bottom & ", 30, 30), True"
End If

If Check2 Then
 Print #1, "'Here i set the form on top most"
 Print #1, "SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE"
End If
Print #1, "End Sub"
Print #1, ""
If Check1 Then
 Print #1, "Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)"
 Print #1, "'Here i move the from when the mouse is down."
 Print #1, "ReleaseCapture"
 Print #1, "SendMessage Me.hwnd, &HA1, 2, 0&"
 Print #1, "End Sub"
End If
Close #1
Err:
End Sub


Private Sub Image3_Click(Index As Integer)
 If Image3(Index).Picture = Image3(1).Picture Then
  Image3(Index).Picture = Image3(2).Picture
  Image4(Index).Picture = Image4(1).Picture
  Form2.xCirkel = True
 Else
  Image3(Index).Picture = Image3(1).Picture
  Image4(Index).Picture = Image4(2).Picture
  Form2.xCirkel = False
 End If
  
End Sub

Private Sub Image4_Click(Index As Integer)
If Image4(Index).Picture = Image4(1).Picture Then
  Image4(Index).Picture = Image4(2).Picture
  Image3(Index).Picture = Image3(1).Picture
  Form2.xCirkel = False
 Else
  Image4(Index).Picture = Image4(1).Picture
  Image3(Index).Picture = Image3(2).Picture
  Form2.xCirkel = True
 End If
End Sub

Private Sub Image4_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 Image4(Index).Picture = Image4(1).Picture
End Sub

Private Sub Image4_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Image4(Index).Picture = Image4(2).Picture
End Sub

Private Sub Image3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 Image3(Index).Picture = Image3(1).Picture
End Sub

Private Sub Image3_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Image3(Index).Picture = Image3(2).Picture
End Sub


Private Sub Image2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 Image2(Index).Picture = Image2(1).Picture
End Sub

Private Sub Image2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Image2(Index).Picture = Image2(2).Picture
End Sub


Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 Image1(Index).Picture = Image1(1).Picture
End Sub

Private Sub Image1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Image1(Index).Picture = Image1(2).Picture
End Sub

Private Sub Slider1_Change()
Form2.Height = Slider1.Value * 15
Slider1.Value = Form2.ScaleHeight
Slider2.Value = Form2.ScaleWidth
Text1.Text = Slider1.Value
Text2.Text = Slider2.Value

End Sub


Private Sub Slider2_Change()
Form2.Width = Slider2.Value * 15
Slider1.Value = Form2.ScaleHeight
Slider2.Value = Form2.ScaleWidth
Text1.Text = Slider1.Value
Text2.Text = Slider2.Value

End Sub

Private Sub Timer1_Timer()
 Form2.Left = Me.Left
 Form2.Top = Me.Top + Me.Height
 DoEvents
End Sub
