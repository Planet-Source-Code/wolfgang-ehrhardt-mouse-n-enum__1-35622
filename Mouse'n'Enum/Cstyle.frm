VERSION 5.00
Begin VB.Form Cstyle 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Form1"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   1815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   1815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command 
      Caption         =   "Restore"
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   3840
      Width           =   1095
   End
   Begin VB.HScrollBar HScroll 
      Height          =   135
      Index           =   1
      LargeChange     =   10
      Left            =   120
      Max             =   810
      Min             =   -810
      TabIndex        =   14
      Top             =   840
      Width           =   1575
   End
   Begin VB.HScrollBar HScroll 
      Height          =   135
      Index           =   0
      LargeChange     =   10
      Left            =   120
      Max             =   812
      TabIndex        =   13
      Top             =   2160
      Width           =   1575
   End
   Begin VB.VScrollBar VScroll 
      Height          =   255
      Index           =   1
      LargeChange     =   10
      Left            =   1440
      Min             =   -32767
      TabIndex        =   12
      Top             =   1320
      Width           =   255
   End
   Begin VB.VScrollBar VScroll 
      Height          =   255
      Index           =   0
      LargeChange     =   10
      Left            =   1440
      Max             =   612
      TabIndex        =   11
      Top             =   2640
      Width           =   255
   End
   Begin VB.Timer GR 
      Interval        =   1000
      Left            =   1440
      Top             =   120
   End
   Begin VB.ComboBox WState 
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown-Liste
      TabIndex        =   1
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label LMore 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Top"
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label LMore 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Left"
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label LMore 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Width"
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label LMore 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Height"
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Ko 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Top"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Ko 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Left"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Ko 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Width"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Ko 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Height"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label LMore 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Hwnd"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label LMore 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "WindowState"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   1575
   End
End
Attribute VB_Name = "Cstyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RC(3) As Long

Dim OrgR As RECT

Dim MyHwnd As Long, OrgState As Long
Private Sub Command_Click()
    Call RestoreWindow(MyHwnd, OrgState, OrgR)
End Sub
Private Sub Form_Load()
    Dim P As Integer
    
    Call LoadStandardForm(Me, True)
    
    For P = HScroll.LBound To HScroll.UBound
        Call SetMP(HScroll(P), True)
        Call SetMP(VScroll(P), True)
    Next P
    
    Call SetMP(WState, True)
    Call SetMP(Command, True)
        
    WState.Clear
    WState.AddItem "Max."
    WState.AddItem "Min."
    WState.AddItem "Norm."
    
    HScroll(0).Max = Screen.Width / Screen.TwipsPerPixelX + 100
    HScroll(1).Max = Screen.Width / Screen.TwipsPerPixelX + 10
    VScroll(0).Max = Screen.Height / Screen.TwipsPerPixelY + 100
    VScroll(1).Max = Screen.Height / Screen.TwipsPerPixelY + 10

    MyHwnd = -1

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Call FormMove(Me)
End Sub
Private Sub GR_Timer()
    Dim TaskID As Long
    
    Call GetWindowThreadProcessId(MyHwnd, TaskID)
    
    If TaskID < 1 Then
        Unload Me
    Else
        Call GetRect(MyHwnd)
    End If
    
End Sub
Private Sub HScroll_Change(Index As Integer)
    If Me.Visible And isUserClick(HScroll(Index).Hwnd) Then
        Call MoveWindow(MyHwnd, HScroll(1).Value, VScroll(1).Value, _
                                HScroll(0).Value, VScroll(0).Value, 1)
    
        Ko(1).Caption = HScroll(0).Value
        Ko(2).Caption = HScroll(1).Value
    End If
End Sub
Private Sub LMore_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Call FormMove(Me)
End Sub
Public Sub GetHwnd(ByVal Hwnd As Long)
    Dim State As Long
    
    If MyHwnd = -1 Then MyHwnd = Hwnd: _
                        OrgState = GetWindowState(MyHwnd): _
                        Call GetWindowRect(MyHwnd, OrgR): _
                        LMore(5).Caption = "Pos from " & MyHwnd
                                  
    Call GetRect(MyHwnd)
        
End Sub
Private Sub GetRect(ByVal Hwnd As Long)
    Dim State As Long
    Dim R As RECT
    
    Call GetWindowRect(Hwnd, R)
    
    If R.Right - R.Left > HScroll(0).Max Then _
                                HScroll(0).Max = R.Right - R.Left
    If R.Right - R.Left < HScroll(0).Min Then _
                                HScroll(0).Min = R.Right - R.Left
    If R.Left > HScroll(1).Max Then HScroll(1).Max = R.Left
    If R.Left < HScroll(1).Min Then HScroll(1).Min = R.Left
    If R.Top > VScroll(1).Max Then VScroll(1).Max = R.Top
    If R.Top < VScroll(1).Min Then VScroll(1).Min = R.Top
    If R.Bottom - R.Top > VScroll(0).Max Then _
                                VScroll(0).Max = R.Bottom - R.Top
    If R.Bottom - R.Top < VScroll(0).Min Then _
                                VScroll(0).Min = R.Bottom - R.Top

    HScroll(0).Value = R.Right - R.Left
    HScroll(1).Value = R.Left
    VScroll(1).Value = R.Top
    VScroll(0).Value = R.Bottom - R.Top
    
    Ko(0).Caption = R.Bottom - R.Top
    Ko(1).Caption = R.Right - R.Left
    Ko(2).Caption = R.Left
    Ko(3).Caption = R.Top

    If CBisDropped(WState.Hwnd) = 0 Then
        State = GetWindowState(MyHwnd)
        
        If State = SW_MINIMIZE Then WState.ListIndex = 1
        If State = SW_MAXIMIZE Then WState.ListIndex = 0
        If State = SW_NORMAL Then WState.ListIndex = 2
    End If

End Sub
Private Sub SetRect()
    RC(0) = CLng(Ko(0).Caption)
    RC(1) = CLng(Ko(1).Caption)
    RC(2) = CLng(Ko(2).Caption)
    RC(3) = CLng(Ko(3).Caption)
    
    Call MoveWindow(MyHwnd, RC(2), RC(3), RC(1), RC(0), 1)
    
    Call GetRect(MyHwnd)
    
End Sub
Private Sub VScroll_Change(Index As Integer)
    If Me.Visible And isUserClick(VScroll(Index).Hwnd) Then
        Call MoveWindow(MyHwnd, HScroll(1).Value, VScroll(1).Value, _
                                HScroll(0).Value, VScroll(0).Value, 1)
    
        Ko(0).Caption = VScroll(0).Value
        Ko(3).Caption = VScroll(1).Value
    End If
End Sub
Private Sub WState_Click()
    Dim nCmdShow As Long
    
    nCmdShow = -1
    
    Select Case WState.ListIndex
        Case 0
            nCmdShow = SW_MAXIMIZE
        Case 1
            nCmdShow = SW_MINIMIZE
        Case 2
            nCmdShow = SW_NORMAL
    End Select
    
    If nCmdShow <> -1 And Me.Visible Then _
                                Call ShowWindow(MyHwnd, nCmdShow): _
                                Call GetHwnd(MyHwnd)
                                
End Sub
