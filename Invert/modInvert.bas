Attribute VB_Name = "modInvert"
' deja_vu
' feedback : deja_vu555@yahoo.com
'
Option Explicit

'api call to invert a given rect
Private Declare Function InvertRect Lib "user32" _
        (ByVal hdc As Long, _
         lpRect As RECT) As Long
         
'rect structure
Private Type RECT
    xx As Long
    yy As Long
    wi As Long
    hi As Long
End Type

'InvertDc function - which inverts a given DC
Public Function InvertDc(destDc As Long, X As Long, Y As Long, h As Long, w As Long) As Boolean
On Error GoTo invERR
Dim recta As RECT

InvertDc = False

'puts the given values into a rect
recta.hi = h
recta.wi = w
recta.xx = X
recta.yy = Y

DoEvents
'inverts the rect region
InvertRect destDc, recta

InvertDc = True

Exit Function
invERR:
InvertDc = False
End Function
