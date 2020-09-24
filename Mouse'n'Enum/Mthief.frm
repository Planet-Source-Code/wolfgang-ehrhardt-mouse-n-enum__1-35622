VERSION 5.00
Begin VB.Form Mthief 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   1410
   ClientLeft      =   150
   ClientTop       =   420
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   94
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   363
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox Owner 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox mID 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox SysID 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Message 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "No Menus were found"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4095
   End
   Begin VB.Menu MyMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "Mthief"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public mRefresh As Boolean

Dim PrevWndProc As Long
Private Sub Form_Load()
    Call LoadStandardForm(Me, False)
    Me.Tag = "mT_" & Me.Hwnd
    
    Me.Height = 660
    Me.Width = 4365
    
    PrevWndProc = WatchWM(Me.Hwnd)
    
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Call FormMove(Me)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim fIndex As Long
    
    Call TerminateWatchWM(Me.Hwnd, PrevWndProc)
    
    If mRefresh Then
        fIndex = FindForm("Ac", "mRefresh")
        If fIndex > -1 Then Unload Forms(fIndex)
    End If
    
End Sub
Private Sub Message_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Call FormMove(Me)
End Sub
