Attribute VB_Name = "modDrawText"
Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Private Const DT_BOTTOM = &H8
Private Const DT_CALCRECT = &H400
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_NOCLIP = &H100
Private Const DT_NOPREFIX = &H800
Private Const DT_RIGHT = &H2
Private Const DT_SINGLELINE = &H20
Private Const DT_TABSTOP = &H80
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_WORDBREAK = &H10
Private Const DT_PATH_ELLIPSIS = &H4000
Private Const DT_MODIFYSTRING = &H10000
Private Const DT_END_ELLIPSIS = &H8000&
Private Const DT_EXPANDTABS = &H40
Private Const DT_EXTERNALLEADING = &H200
Private Const DT_INTERNAL = &H1000
Private Const DT_EDITCONTROL = &H2000
Private Const DT_RTLREADING = &H20000
Private Const DT_WORD_ELLIPSIS = &H40000
Private Const DT_NOFULLWIDTHCHARBREAK = &H80000
Private Const DT_HIDEPREFIX = &H100000
Private Const DT_PREFIXONLY = &H200000

Public Sub pDrawText(hdc As Long, ByVal sText As String, x1 As Long, y1 As Long, x2 As Long, y2 As Long)
Dim R As RECT
    R.Left = x1: R.Top = y1: R.Right = x2: R.Bottom = y2
    DrawText hdc, sText, Len(sText), R, DT_LEFT Or DT_TOP Or DT_WORDBREAK Or DT_SINGLELINE Or DT_MODIFYSTRING Or DT_END_ELLIPSIS Or DT_RTLREADING
    'DrawText hdc, StrConv(sText, vbUnicode) & Chr(0), Len(sText), R, DT_LEFT Or DT_TOP Or DT_WORDBREAK Or DT_SINGLELINE Or DT_MODIFYSTRING Or DT_END_ELLIPSIS Or DT_RTLREADING
End Sub

