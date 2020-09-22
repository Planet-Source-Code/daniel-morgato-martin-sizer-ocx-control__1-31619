VERSION 5.00
Begin VB.Form frmNode 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   105
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   7
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Paint()
Dim Bordercolor As Long, BackFillColor As Long
    BackColor = RGB(255, 255, 255)
    Line (0, 0)-(100, 100), RGB(0, 0, 128), BF
End Sub


