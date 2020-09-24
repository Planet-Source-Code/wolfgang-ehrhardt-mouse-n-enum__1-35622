VERSION 5.00
Begin VB.Form AllStyle 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Form1"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ListBox OrgName 
      Height          =   255
      Left            =   3480
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Caption         =   "Restore"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   2640
      Width           =   855
   End
   Begin VB.Timer CheckHwnd 
      Interval        =   1000
      Left            =   3840
      Top             =   2400
   End
   Begin VB.CommandButton Command 
      Caption         =   "Refresh"
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   7
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "Close"
      Height          =   255
      Index           =   0
      Left            =   2400
      TabIndex        =   6
      Top             =   2640
      Width           =   855
   End
   Begin VB.ListBox LB 
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1035
      Index           =   1
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   1560
      Width           =   3135
   End
   Begin VB.ListBox Value 
      Height          =   255
      Left            =   3480
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox Konst 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   255
      ItemData        =   "AllStyle.frx":0000
      Left            =   3480
      List            =   "AllStyle.frx":0002
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox LB 
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1035
      Index           =   0
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label Label 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
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
      Index           =   1
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   3135
   End
   Begin VB.Label Label 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
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
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu MenuItem 
         Caption         =   "Change to"
         Index           =   0
      End
      Begin VB.Menu MenuItem 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu MenuItem 
         Caption         =   "Close Menu"
         Index           =   2
      End
   End
End
Attribute VB_Name = "AllStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type WS1
    Style As Long
    nIndex As Long
End Type

Dim WSi As WS1, OrgR As RECT

Dim MyHwnd As Long, OrgState As Long
Private Sub CheckHwnd_Timer()
    Dim TaskID As Long
    
    Call GetWindowThreadProcessId(MyHwnd, TaskID)
    
    If TaskID < 1 Then Unload Me
    
End Sub
Private Sub Command_Click(Index As Integer)
    Dim P As Integer
    Dim Hwnd As Long, nIndex As Long, Style As Long, lngStyle As Long
    Dim Item As String
    Dim Has As Boolean
    
    Select Case Index
        Case 0
            Unload Me
        Case 1
            Call GetHwnd(MyHwnd)
        Case 2
            Call RestoreWindow(MyHwnd, OrgState, OrgR)
            
            For P = 0 To Konst.ListCount - 1
                Item = Konst.List(P)
                Style = Value.List(P)
                
                nIndex = GWL_STYLE
                If InStr(Item, "_EX_") Then nIndex = GWL_EXSTYLE
                
                Has = False
                lngStyle = GetWindowLong(MyHwnd, nIndex)
                If (lngStyle And Style) Then Has = True
                
                Li = SearchLB(OrgName.Hwnd, LB_FINDSTRINGEXACT, _
                              -1, Item)
                              
                If (Not Has And Li > -1) Or (Has And Li = -1) Then _
                    Call ChangeWindowStyle(MyHwnd, nIndex, Style)
            Next P
            
            Hwnd = MyHwnd
            MyHwnd = -1
            
            Call GetHwnd(Hwnd)
    End Select
End Sub
Private Sub Form_Load()
    Dim P As Integer
    
    Call LoadStandardForm(Me, True)
    
    Me.Height = 3300
    Me.Width = 3480
    
    Call FillBox
    
    Call SetMP(LB(0), True)
    Call SetMP(LB(1), True)
    
    For P = Command.LBound To Command.UBound
        Call SetMP(Command(P), True)
    Next P
    
    MyHwnd = -1
    
    GetOrg = True
    
End Sub
Public Sub GetHwnd(ByVal Hwnd As Long)
    Dim P As Integer, Li As Integer
    Dim nIndex As Long, Ivalue As Long, Style As Long
    Dim R As RECT
    Dim Item As String
    
    LB(0).Clear
    LB(1).Clear
    
    If MyHwnd = -1 Then _
        Label(1).Caption = "Control " & Hwnd & " has style": _
        Label(0).Caption = "Control " & Hwnd & " has not style": _
        Call GetWindowRect(Hwnd, R): _
        Call GetWindowRect(Hwnd, OrgR): _
        OrgState = GetWindowState(Hwnd)
        
    For P = 0 To Value.ListCount - 1
        nIndex = GWL_STYLE
        If InStr(Konst.List(P), "_EX_") Then nIndex = GWL_EXSTYLE
        Style = GetWindowLongA(Hwnd, nIndex)
        
        Ivalue = CLng(Value.List(P))
        
        If (Style And Ivalue) Then
            LB(0).AddItem Konst.List(P)
        Else
            LB(1).AddItem Konst.List(P)
        End If
    Next P
    
    If MyHwnd = -1 Then
        OrgName.Clear
        
        For P = 0 To LB(0).ListCount - 1
            OrgName.AddItem LB(0).List(P)
        Next P
        
        MyHwnd = Hwnd
        Me.Visible = True
    End If

End Sub
Private Sub FillBox()
    Dim P As Integer, X As Integer

    Konst.Clear
    Value.Clear

    Konst.AddItem "WS_ACTIVECAPTION,&H1"
    Konst.AddItem "WS_BORDER,&H800000"
    Konst.AddItem "WS_CAPTION,&HC00000"
    Konst.AddItem "WS_CHILD,&H40000000"
    Konst.AddItem "WS_CLIPCHILDREN,&H2000000"
    Konst.AddItem "WS_CLIPSIBLINGS,&H4000000"
    Konst.AddItem "WS_DISABLED,&H8000000"
    Konst.AddItem "WS_DLGFRAME,&H400000"
    Konst.AddItem "WS_EX_ACCEPTFILES,&H10"
    Konst.AddItem "WS_EX_APPWINDOW,&H40000"
    Konst.AddItem "WS_EX_CLIENTEDGE,&H200"
    Konst.AddItem "WS_EX_CONTEXTHELP,&H400"
    Konst.AddItem "WS_EX_CONTROLPARENT,&H10000"
    Konst.AddItem "WS_EX_DLGMODALFRAME,&H1"
    Konst.AddItem "WS_EX_LAYERED,&H80000"
    Konst.AddItem "WS_EX_LAYOUTRTL,&H400000"
    Konst.AddItem "WS_EX_LEFT,&H0"
    Konst.AddItem "WS_EX_LEFTSCROLLBAR,&H4000"
    Konst.AddItem "WS_EX_LTRREADING,&H0"
    Konst.AddItem "WS_EX_MDICHILD,&H40"
    Konst.AddItem "WS_EX_NOACTIVATE,&H8000000"
    Konst.AddItem "WS_EX_NOINHERITLAYOUT,&H100000"
    Konst.AddItem "WS_EX_NOPARENTNOTIFY,&H4"
    Konst.AddItem "WS_EX_OVERLAPPEDWINDOW,&H300"
    Konst.AddItem "WS_EX_PALETTEWINDOW,&H188"
    Konst.AddItem "WS_EX_RIGHT,&H1000"
    Konst.AddItem "WS_EX_RIGHTSCROLLBAR,&H0"
    Konst.AddItem "WS_EX_RTLREADING,&H2000"
    Konst.AddItem "WS_EX_STATICEDGE,&H20000"
    Konst.AddItem "WS_EX_TOOLWINDOW,&H80"
    Konst.AddItem "WS_EX_TOPMOST,&H8"
    Konst.AddItem "WS_EX_TRANSPARENT,&H20"
    Konst.AddItem "WS_EX_WINDOWEDGE,&H100"
    Konst.AddItem "WS_GROUP,&H20000"
    Konst.AddItem "WS_GT,&H30000"
    Konst.AddItem "WS_HSCROLL,&H100000"
    Konst.AddItem "WS_ICONIC,&H20000000"
    Konst.AddItem "WS_MAXIMIZE,&H1000000"
    Konst.AddItem "WS_MAXIMIZEBOX,&H10000"
    Konst.AddItem "WS_MINIMIZE,&H20000000"
    Konst.AddItem "WS_MINIMIZEBOX,&H20000"
    Konst.AddItem "WS_OVERLAPPED,&H0"
    Konst.AddItem "WS_OVERLAPPEDWINDOW,&HCF0000"
    Konst.AddItem "WS_POPUP,&H80000000"
    Konst.AddItem "WS_POPUPWINDOW,&H80880000"
    Konst.AddItem "WS_SIZEBOX,&H40000"
    Konst.AddItem "WS_SYSMENU,&H80000"
    Konst.AddItem "WS_TABSTOP,&H10000"
    Konst.AddItem "WS_THICKFRAME,&H40000"
    Konst.AddItem "WS_TILED,&H0"
    Konst.AddItem "WS_TILEDWINDOW,&HCF0000"
    Konst.AddItem "WS_VISIBLE,&H10000000"
    Konst.AddItem "WS_VSCROLL,&H200000"

    For P = 0 To Konst.ListCount - 1
        X = InStr(Konst.List(P), ",")
        Value.AddItem mID(Konst.List(P), X + 1), P
        Konst.List(P) = mID(Konst.List(P), 1, X - 1)
    Next P
        
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Call FormMove(Me)
End Sub
Private Sub Label_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Call FormMove(Me)
End Sub
Private Sub LB_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Li As Integer
    Dim Item As String
    
    If Button = vbRightButton Then
        Item = LB(Index).List(LB(Index).ListIndex)
        Li = SearchLB(Konst.Hwnd, LB_FINDSTRINGEXACT, -1, Item)
        
        WSi.nIndex = GWL_STYLE
        If InStr(Konst.List(Li), "_EX_") Then nIndex = GWL_EXSTYLE
        
        WSi.Style = CLng(Value.List(Li))
        
        MenuItem(0).Caption = "Change to 'has'"
        If Index = 0 Then MenuItem(0).Caption = _
                              "Change to 'has not'"
                              
        PopupMenu Menu
    End If
End Sub
Private Sub LB_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call LBmouseMove(LB(Index), Button, Shift, X, Y)
End Sub
Private Sub MenuItem_Click(Index As Integer)
    If Index = 0 Then Call ChangeWindowStyle(MyHwnd, _
                                             WSi.nIndex, _
                                             WSi.Style): _
                           Call GetHwnd(MyHwnd)
End Sub
