Attribute VB_Name = "Module1"
Option Explicit


Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public FocusRec As RECT

Public Declare Function DrawFocusRect _
Lib "user32" ( _
    ByVal hdc As Long, _
    lpRect As RECT _
) As Long


