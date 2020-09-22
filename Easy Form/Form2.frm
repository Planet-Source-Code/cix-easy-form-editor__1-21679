VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Easy From"
   ClientHeight    =   3360
   ClientLeft      =   3195
   ClientTop       =   3375
   ClientWidth     =   7725
   DrawMode        =   1  'Blackness
   LinkTopic       =   "Form2"
   ScaleHeight     =   224
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   515
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   7665
      TabIndex        =   1
      Top             =   0
      Width           =   7695
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   2985
      Left            =   0
      MousePointer    =   2  'Cross
      ScaleHeight     =   197
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   511
      TabIndex        =   0
      Top             =   360
      Width           =   7695
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   0
         ScaleHeight     =   127
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   207
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   3135
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

Private sngX1 As Single
Private sngY1 As Single
Private sngX2 As Single
Private sngY2 As Single

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Public xCirkel As Boolean

Private Sub Form_Resize()
' Oki welll where i make sure all the pictures are in place.
 Picture2.Top = 0
 Picture2.Left = 0
 Picture2.Width = Me.ScaleWidth
 Picture1.Top = Picture2.Height
 Picture1.Left = 0
 Picture1.Height = Me.ScaleHeight - Picture2.Height
 Picture1.Width = Me.ScaleWidth
 Picture2.CurrentY = (Picture2.Width / 2)
 Picture2.Cls
 Picture2.Print "  Easy Form"
 Picture3.Left = 0
 Picture3.Height = Picture1.Height
 Picture3.Width = Picture1.Width
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 ' Move the form
 ReleaseCapture
 SendMessage Me.hwnd, &HA1, 2, 0&
End Sub

Private Sub Picture1_MouseDown( _
    Button As Integer, _
    Shift As Integer, _
    X As Single, _
    Y As Single _
)
    
    If (Button And vbLeftButton) = 0 Then Exit Sub
    ' Clean the picture..
    Picture1.Cls
    'Show a picture with out Autoredraw, so you can draw the Focus on the form
    Picture3.Visible = True
    sngX1 = X
    sngY1 = Y
End Sub

Private Sub Picture1_MouseMove( _
    Button As Integer, _
    Shift As Integer, _
    X As Single, _
    Y As Single _
)
    
    If (Button And vbLeftButton) = 0 Then Exit Sub
    
    If (sngX2 <> 0) Or (sngY2 <> 0) Then
       ' Draw the focus on the form with out Autoredraw
        DrawFocusRect Picture3.hdc, FocusRec
    End If
    
    sngX2 = X
    sngY2 = Y
    
    FocusRec.Left = sngX1
    FocusRec.Top = sngY1
    FocusRec.Right = sngX2
    FocusRec.Bottom = sngY2
    
    ' If the focus is draw'en from the right to the left then swap the values.
    If sngY2 < sngY1 Then Swap FocusRec.Top, FocusRec.Bottom
    If sngX2 < sngX1 Then Swap FocusRec.Left, FocusRec.Right
    
    DrawFocusRect Picture3.hdc, FocusRec
    Refresh
End Sub

Private Sub Picture1_MouseUp( _
    Button As Integer, _
    Shift As Integer, _
    X As Single, _
    Y As Single _
)
    
    If (Button And vbLeftButton) = 0 Then Exit Sub
    
    If FocusRec.Right Or FocusRec.Bottom Then
        DrawFocusRect Picture3.hdc, FocusRec
    End If
    Picture3.Visible = False
    Picture1.ScaleMode = vbPixels
    Picture1.FillColor = vbBlue
    Picture1.ForeColor = vbBlue
    If xCirkel Then
     Picture1.FillColor = vbRed
     Picture1.ForeColor = vbRed
     Ellipse Picture1.hdc, FocusRec.Left, FocusRec.Top, FocusRec.Right, FocusRec.Bottom
     
     
   Else
     ' This is how you draw a normale square
     'Picture1.Line (sngX1, sngY1)-(sngX2, sngY2), vbBlue, B
     'Picture1.Refresh
      'Draw the Round Rectangle
      ' 30, 30 is the softness of the corners.
      RoundRect Picture1.hdc, FocusRec.Left, FocusRec.Top, FocusRec.Right, FocusRec.Bottom, 30, 30
    End If
    
    If Form1.Check3.Value Then
    'if autosize is set then size the form
    ' X * 15 is just my cheap way  of turn the scalemode pixles into Twips.
     Form2.Width = (X * 15) + 200
     Form2.Height = (Y * 15) + 500
     Form1.Slider1.Value = Me.ScaleHeight
     Form1.Slider2.Value = Me.ScaleWidth
     
    End If
    'Clean up.
    sngX1 = 0
    sngY1 = 0
    sngX2 = 0
    sngY2 = 0
End Sub

Private Sub Swap(vntA As Variant, vntB As Variant)
    Dim vntT As Variant
    vntT = vntA
    vntA = vntB
    vntB = vntT
End Sub



