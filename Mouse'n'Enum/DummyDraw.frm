VERSION 5.00
Begin VB.Form DummyDraw 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'Kein
   ClientHeight    =   585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   39
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   69
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
End
Attribute VB_Name = "DummyDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call StayOnTop(Me)
    Me.Tag = Me.Name
End Sub
