VERSION 5.00
Begin VB.Form Dummy 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'Kein
   ClientHeight    =   945
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   1755
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   63
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   117
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Visible         =   0   'False
   Begin VB.Timer TakeSnapShot 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   480
   End
   Begin VB.Label DummyCap 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "DummyCap"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu MenuItem 
         Caption         =   "Hwnd"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu MenuItem 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu MenuItem 
         Caption         =   "Enumieren"
         Index           =   2
         Begin VB.Menu EnumItem 
            Caption         =   "Control enumieren"
            Index           =   0
         End
         Begin VB.Menu EnumItem 
            Caption         =   "Alle Childs enumieren"
            Index           =   1
         End
         Begin VB.Menu EnumItem 
            Caption         =   "Parent enumieren"
            Index           =   2
         End
         Begin VB.Menu EnumItem 
            Caption         =   "Menu enumieren"
            Index           =   3
            Begin VB.Menu MenuEnumItem 
               Caption         =   "MenuKurzInfo"
               Index           =   0
            End
            Begin VB.Menu MenuEnumItem 
               Caption         =   "Menu enumieren"
               Index           =   1
            End
            Begin VB.Menu MenuEnumItem 
               Caption         =   "Sysmenu enumieren"
               Index           =   2
            End
            Begin VB.Menu MenuEnumItem 
               Caption         =   "Alle Menus von allen Childs"
               Index           =   3
            End
         End
      End
      Begin VB.Menu MenuItem 
         Caption         =   "Text"
         Index           =   3
         Begin VB.Menu wTextItem 
            Caption         =   "SendText"
            Index           =   0
            Begin VB.Menu wTextSendItem 
               Caption         =   "SendText by String"
               Index           =   0
            End
            Begin VB.Menu wTextSendItem 
               Caption         =   "SendText by Char"
               Index           =   1
            End
         End
         Begin VB.Menu wTextItem 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu wTextItem 
            Caption         =   "SetText"
            Index           =   2
         End
         Begin VB.Menu wTextItem 
            Caption         =   "Set Text from Clipboard"
            Index           =   3
         End
         Begin VB.Menu wTextItem 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu wTextItem 
            Caption         =   "Copy Text to Clipboard"
            Index           =   5
         End
      End
      Begin VB.Menu MenuItem 
         Caption         =   "Window"
         Index           =   4
         Begin VB.Menu WinMenu 
            Caption         =   "Show/Hide"
            Index           =   0
         End
         Begin VB.Menu WinMenu 
            Caption         =   "Bring to Top"
            Index           =   1
         End
         Begin VB.Menu WinMenu 
            Caption         =   "Stay on Top"
            Index           =   2
         End
         Begin VB.Menu WinMenu 
            Caption         =   "Show Windows"
            Index           =   3
            Begin VB.Menu WinShowMenuItem 
               Caption         =   "Maximized"
               Index           =   0
            End
            Begin VB.Menu WinShowMenuItem 
               Caption         =   "Normal"
               Index           =   1
            End
            Begin VB.Menu WinShowMenuItem 
               Caption         =   "Minimized"
               Index           =   2
            End
         End
      End
      Begin VB.Menu MenuItem 
         Caption         =   "Change Position"
         Index           =   5
      End
      Begin VB.Menu MenuItem 
         Caption         =   "Change WindowStyle"
         Index           =   6
      End
      Begin VB.Menu MenuItem 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu MenuItem 
         Caption         =   "Click"
         Index           =   8
         Begin VB.Menu ClickMenuItem 
            Caption         =   "Click Control"
            Index           =   0
         End
         Begin VB.Menu ClickMenuItem 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu ClickMenuItem 
            Caption         =   "MousePos x,y"
            Index           =   2
            Begin VB.Menu ClickMPmenuItem 
               Caption         =   "Click"
               Index           =   0
            End
            Begin VB.Menu ClickMPmenuItem 
               Caption         =   "DoubleClick"
               Index           =   1
            End
            Begin VB.Menu ClickMPmenuItem 
               Caption         =   "-"
               Index           =   2
            End
            Begin VB.Menu ClickMPmenuItem 
               Caption         =   "RightClick"
               Index           =   3
            End
         End
      End
      Begin VB.Menu MenuItem 
         Caption         =   "Control SetParent"
         Index           =   9
      End
      Begin VB.Menu MenuItem 
         Caption         =   "Control WM_CLOSE"
         Index           =   10
      End
      Begin VB.Menu MenuItem 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu MenuItem 
         Caption         =   "Jump to"
         Index           =   12
         Begin VB.Menu SwitchMenuItem 
            Caption         =   "Parent"
            Index           =   0
         End
         Begin VB.Menu SwitchMenuItem 
            Caption         =   "Parent && show it"
            Index           =   1
         End
         Begin VB.Menu SwitchMenuItem 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu SwitchMenuItem 
            Caption         =   "TopParent"
            Index           =   3
         End
         Begin VB.Menu SwitchMenuItem 
            Caption         =   "TopParent && show it"
            Index           =   4
         End
      End
      Begin VB.Menu MenuItem 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu MenuItem 
         Caption         =   "SnapShot"
         Index           =   14
      End
      Begin VB.Menu MenuItem 
         Caption         =   "-"
         Index           =   15
      End
      Begin VB.Menu MenuItem 
         Caption         =   "Read List- or ComboBox"
         Index           =   16
      End
      Begin VB.Menu MenuItem 
         Caption         =   "-"
         Index           =   17
      End
      Begin VB.Menu MenuItem 
         Caption         =   "Terminate Thread ;)"
         Index           =   18
      End
      Begin VB.Menu MenuItem 
         Caption         =   "-"
         Index           =   19
      End
      Begin VB.Menu MenuItem 
         Caption         =   "Capture Control"
         Index           =   20
      End
      Begin VB.Menu MenuItem 
         Caption         =   "-"
         Index           =   21
      End
      Begin VB.Menu MenuItem 
         Caption         =   "Close Menu"
         Index           =   22
      End
   End
End
Attribute VB_Name = "Dummy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetCursorPos Lib "USER32" (ByVal X As Long, ByVal Y As Long) As Long

Private Type MYWIN1
    Parent    As Long
    TopParent As Long
End Type

Public mX As Long, mY As Long

Private MyForms()

Private MyWin As MYWIN1

Dim MyHwnd As Long, SnapHwnd As Long
Private Sub ClickMenuItem_Click(Index As Integer)
    
    Select Case Index
        Case 0
            Call ClickIcon(MyHwnd)
    End Select
    
End Sub
Private Sub ClickMPmenuItem_Click(Index As Integer)
    Dim nPoint As POINTAPI
    Dim P As Integer
    
    Select Case Index
        Case 0, 1, 3
            Call GetCursorPos(nPoint)
            SetCursorPos mX, mY
                        
            If Index < 3 Then
                For P = 1 To Index + 1
                    Call mouse_event(MOUSEEVENTF_LEFTDOWN Or _
                                     MOUSEEVENTF_LEFTUP, _
                                     0&, 0&, CLng(0), CLng(0))
                Next P
            Else
                Call mouse_event(MOUSEEVENTF_RIGHTDOWN Or _
                                     MOUSEEVENTF_RIGHTUP, _
                                     0&, 0&, CLng(0), CLng(0))
            End If
            
            SetCursorPos nPoint.X, nPoint.Y
    End Select

End Sub
Private Sub EnumItem_Click(Index As Integer)
    Dim Item As String
    
    Item = LCase(EnumItem(Index).Caption)
    
    Select Case Item
        Case "control enumieren"
            Call LEnum(MyHwnd, True, False)
        Case "alle childs enumieren"
            Call LEnum(MyHwnd, False, True)
        Case "parent enumieren"
            Call LEnum(MyHwnd, False, False)
    End Select
    
End Sub
Private Sub Form_Load()
        
    Me.Height = 0
    Me.Width = 0
    
    Call StayOnTop(Me)
    
End Sub
Public Sub PopUp(ByVal Hwnd As Long)
    Static MenuAction As Boolean

    If Not MenuAction And Not Me.TakeSnapShot.Enabled Then
        MyHwnd = Hwnd
         
        MenuItem(0).Caption = "Control " & str(MyHwnd)
    
        MenuAction = True
            PopupMenu Menu
        MenuAction = False
    End If

End Sub
Private Sub Menu_Click()
            
    If GetWindowLongA(MyHwnd, GWL_EXSTYLE) And WS_EX_TOPMOST Then
        WinMenu(2).Caption = "Remove from Top"
    Else
        WinMenu(2).Caption = "Stay on Top"
    End If
    
    ClickMenuItem(2).Caption = "MousePos (" & mX & ", " & mY & ")"
    
    WinMenu(0).Caption = "Show"
    If IsWindowVisible(MyHwnd) Then WinMenu(0).Caption = "Hide"
        
    If GetMenuItemCount(GetMenu(MyHwnd)) > 0 Then
        MenuEnumItem(1).Enabled = True
        MenuEnumItem(1).Caption = "Menu enumieren"
    Else
        MenuEnumItem(1).Enabled = False
        MenuEnumItem(1).Caption = "GetMenu(" & MyHwnd & ") = 0"
    End If
    
    If GetSystemMenu(MyHwnd, False) > 0 Then
        MenuEnumItem(2).Enabled = True
        MenuEnumItem(2).Caption = "SystemMenu enumieren"
    Else
        MenuEnumItem(2).Enabled = False
        MenuEnumItem(2).Caption = "GetSystemMenu(" & MyHwnd & ") = 0"
    End If
    
    MyWin.Parent = GetParent(MyHwnd)
    MyWin.TopParent = GetTopParent(MyWin.Parent)
    
    SwitchMenuItem(0).Enabled = False
    SwitchMenuItem(1).Enabled = False
    SwitchMenuItem(3).Enabled = False
    SwitchMenuItem(4).Enabled = False

    If MyWin.Parent Then SwitchMenuItem(0).Enabled = True: _
                         SwitchMenuItem(1).Enabled = True
    If MyWin.TopParent And MyWin.TopParent <> MyWin.Parent Then _
                               SwitchMenuItem(3).Enabled = True: _
                               SwitchMenuItem(4).Enabled = True
    
End Sub
Private Sub MenuEnumItem_Click(Index As Integer)
    Dim A As New Ac
    Dim Mi As New MnuInfo
    
    If Index Then Load A
    
    Select Case Index
        Case 0
            Load Mi
            Call Mi.GetHwnd(MyHwnd)
            Mi.Visible = True
        Case 1
            Call A.MenuEnum(MyHwnd, False, False)
            A.Visible = True
        Case 2
            Call A.MenuEnum(MyHwnd, True, False)
            A.Visible = True
        Case 3
            Call A.MenuEnum(MyHwnd, False, True)
    End Select
    
End Sub
Private Sub MenuItem_Click(Index As Integer)
    Dim X As Long, Li As Long, fIndex As Long
    Dim t As String, Text As String, Item As String, Answer As String
    Dim V As Boolean
    Dim C As New Cap, CB As New CbLB
    Dim AT As New AllStyle, CS As New Cstyle
    Dim P As Integer

    On Local Error Resume Next
    
    Item = LCase(MenuItem(Index).Caption)
    
    Select Case Item
            Case "change windowstyle"
                Load AT
                Call AT.GetHwnd(MyHwnd)
            Case "change position"
                Load CS
                Call CS.GetHwnd(MyHwnd)
                CS.Visible = True
            Case "snapshot"
                SnapHwnd = MyHwnd
                TakeSnapShot.Enabled = True
            Case "terminate thread ;)"
                Call Terminate(MyHwnd)
            Case "control setparent"
                t = "Bitte das neue Parentwindow angeben " & _
                    "(Gebe Me für Hwnd dieses Programms ein"
                Answer = uInput(t, "SetParent")
                If LCase(Answer) = "me" Then t = Main.Hwnd
                X = CLng(str(t))
                If X Then Call SetParent(MyHwnd, X)
            Case "control wm_close"
                Call SendMessage(MyHwnd, WM_CLOSE, 0, 0)
            Case "capture control"
                fIndex = FindForm("Cap", "Cap" & MyHwnd)
                If fIndex = -1 Then
                    Load C
                    Call C.GetStaticInfo(MyHwnd)
                    Main.CapC.AddItem MyHwnd
                    Main.CapC.Enabled = True
                Else
                    Forms(fIndex).Show
                    Call SetForegroundWindow(Forms(fIndex).Hwnd)
                    Forms(fIndex).Blink.Enabled = True
                End If
            Case "read list- or combobox"
                Load CB
                Call CB.GetHwnd(MyHwnd)
                CB.Show
    End Select
    
End Sub
Private Sub SwitchMenuItem_Click(Index As Integer)
    Dim h As Long
    Dim R As RECT
    
    h = 0
    
    Select Case Index
        Case 0, 1
            h = MyWin.Parent
        Case 3, 4
            h = GetTopParent(h)
    End Select
    
    If h Then
        If Index = 1 Or Index = 3 Then Call ShowWindow(h, SW_SHOW)
        Call GetWindowRect(h, R)
        SetCursorPos R.Left + (R.Right - R.Left) / 2, R.Top + 10
        Call FlashWindow(h)
    End If
    
End Sub
Private Sub TakeSnapShot_Timer()
    Dim P As Integer
    Dim SS As New SnapShot
    
    TakeSnapShot.Enabled = False
    Main.GetHwnd.Enabled = False
    
    ReDim MyForms(0)
    
    For P = 0 To Forms.Count - 1
        If Forms(P).Visible _
        And Forms(P).Hwnd <> SnapHwnd Then
            Forms(P).Visible = False: _
            ReDim Preserve MyForms(UBound(MyForms) + 1)
            MyForms(UBound(MyForms)) = Forms(P).Hwnd
        End If
    Next P
                                          
    Load SS
    Call SS.GetWindowHwnd(SnapHwnd)
    
    Call Main.GetHwnd_Timer
    
    For P = 1 To UBound(MyForms)
        Call ShowWindow(MyForms(P), SW_SHOWNOACTIVATE)
    Next P

    ReDim MyForms(0)

    Main.GetHwnd.Enabled = True
    
    SS.SetFocus
    
End Sub
Private Sub WinMenu_Click(Index As Integer)
            
    Select Case LCase(WinMenu(Index).Caption)
        Case "hide"
            Call ShowWindow(MyHwnd, SW_HIDE)
        Case "show"
            Call ShowWindow(MyHwnd, SW_SHOW)
        Case "bring to top"
            Call BringWindowToTop(MyHwnd)
        Case "stay on top"
            Call StayWinOnTop(MyHwnd, False)
        Case "remove from top"
            Call StayWinOnTop(MyHwnd, True)
    End Select
    
End Sub
Private Sub WinShowMenuItem_Click(Index As Integer)
    Dim nCmdShow As Long
    
    nCmdShow = -1
    
    Select Case Index
        Case 0
            nCmdShow = SW_MAXIMIZE
        Case 1
            nCmdShow = SW_NORMAL
        Case 2
            nCmdShow = SW_MINIMIZE
    End Select
    
    If nCmdShow <> -1 Then Call ShowWindow(MyHwnd, nCmdShow)
    
End Sub
Private Sub wTextItem_Click(Index As Integer)
    Dim Text As String
    
    Select Case LCase(wTextItem(Index).Caption)
        Case "settext"
            Text = uInput("Bitte Text eingeben", "SetText")
            Call SetText(MyHwnd, Text)
        Case "set text from clipboard"
            Call SetText(MyHwnd, Clipboard.GetText)
        Case "copy text to clipboard"
            Clipboard.SetText GetText(MyHwnd)
    End Select

End Sub
Private Sub wTextSendItem_Click(Index As Integer)
    Dim Text As String
    Dim P As Integer, ASCII As Integer
    
    Text = uInput("Bitte den zu sendenen Text angeben", _
                   wTextSendItem(Index).Caption)
                   
    Select Case Index
        Case 0
            Call SetText(MyHwnd, Text)
        Case 1
            For P = 1 To Len(Text)
                ASCII = Asc(mID(Text, P, 1))
                Call SendMessageByNum(MyHwnd, WM_CHAR, ASCII, 0)
            Next P
    End Select
        
    Call SendMessageByNum(MyHwnd, WM_CHAR, vbKeyReturn, 0)
    
End Sub
