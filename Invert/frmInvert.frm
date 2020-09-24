VERSION 5.00
Begin VB.Form frmInvert 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Invert Example"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   592
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   352
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pbInv 
      AutoSize        =   -1  'True
      Height          =   8865
      Left            =   0
      Picture         =   "frmInvert.frx":0000
      ScaleHeight     =   587
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   348
      TabIndex        =   0
      Top             =   0
      Width           =   5280
   End
End
Attribute VB_Name = "frmInvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' deja_vu
' feedback : deja_vu555@yahoo.com
'
Option Explicit

'used to store the old x and y values.. during the mouse down on the picture
Dim oldX As Long
Dim oldY As Long

Private Sub pbInv_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
'saves the old mouse x and y vals
oldX = X
oldY = Y
End If
End Sub

Private Sub pbInv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    'clears the picture box
    pbInv.Cls
    'draws a rectangular selection
    pbInv.Line (oldX, oldY)-(X, Y), , B
    'uncomment this if you want to inver as you go on selecting
    'InvertDc pbInv.hdc, oldX, oldY, CLng(Y), CLng(X)
End If
End Sub

Private Sub pbInv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    'uncomment this is you uncommented the above line
    'the next invert function call will invert an already inverted region
    'if we dont clear the picture box.
    'pbInv.Cls
    'calling the inver function
    InvertDc pbInv.hdc, oldX, oldY, CLng(Y), CLng(X)
End If
End Sub

