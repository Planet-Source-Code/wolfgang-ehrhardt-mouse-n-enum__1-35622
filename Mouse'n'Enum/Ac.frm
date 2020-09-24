VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Ac 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Form1"
   ClientHeight    =   4230
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ListBox SysID 
      Height          =   255
      Left            =   3720
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer CheckHwnd 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   4080
      Top             =   1440
   End
   Begin VB.ListBox Childs 
      Height          =   255
      Left            =   3720
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox Par 
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComctlLib.TreeView Tree 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   6165
      _Version        =   393217
      HideSelection   =   0   'False
      LineStyle       =   1
      Style           =   6
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Index           =   2
      Left            =   1080
      TabIndex        =   5
      Top             =   3960
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   4
      Top             =   3960
      Width           =   2655
   End
   Begin VB.Label Work 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Working..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   4215
   End
   Begin VB.Label Info 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   3960
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu MItem 
         Caption         =   "Refresh"
         Index           =   0
      End
      Begin VB.Menu MItem 
         Caption         =   "Expand all"
         Index           =   1
      End
      Begin VB.Menu MItem 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu MItem 
         Caption         =   "View"
         Index           =   3
         Begin VB.Menu ViewItem 
            Caption         =   "Only vsisble Controls"
            Index           =   0
         End
         Begin VB.Menu ViewItem 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu ViewItem 
            Caption         =   "Only unvisible Controls"
            Index           =   2
         End
         Begin VB.Menu ViewItem 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu ViewItem 
            Caption         =   "Visible and unvisible Controls"
            Index           =   4
         End
         Begin VB.Menu ViewItem 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu ViewItem 
            Caption         =   "Only with Text"
            Index           =   6
         End
      End
      Begin VB.Menu MItem 
         Caption         =   "Infos"
         Index           =   4
         Begin VB.Menu MenuItem 
            Caption         =   "Hwnd"
            Checked         =   -1  'True
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu MenuItem 
            Caption         =   "Classname"
            Index           =   1
         End
         Begin VB.Menu MenuItem 
            Caption         =   "TaskID"
            Index           =   2
         End
         Begin VB.Menu MenuItem 
            Caption         =   "Visible"
            Index           =   3
         End
         Begin VB.Menu MenuItem 
            Caption         =   "Text"
            Index           =   4
         End
      End
      Begin VB.Menu MItem 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu MItem 
         Caption         =   "Close Window"
         Index           =   6
      End
      Begin VB.Menu MItem 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu MItem 
         Caption         =   "Close Menu"
         Index           =   8
      End
   End
   Begin VB.Menu mMenu 
      Caption         =   "mMenu"
      Visible         =   0   'False
      Begin VB.Menu mMenuItem 
         Caption         =   "Refresh"
         Index           =   0
      End
      Begin VB.Menu mMenuItem 
         Caption         =   "Expand All"
         Index           =   1
      End
      Begin VB.Menu mMenuItem 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mMenuItem 
         Caption         =   "Infos"
         Index           =   3
         Begin VB.Menu mInfos 
            Caption         =   "Menutext"
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu mInfos 
            Caption         =   "MenuID"
            Index           =   1
         End
         Begin VB.Menu mInfos 
            Caption         =   "MenuHwnd"
            Index           =   2
         End
      End
      Begin VB.Menu mMenuItem 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mMenuItem 
         Caption         =   "Close Menu"
         Index           =   5
      End
   End
   Begin VB.Menu mClick 
      Caption         =   "mClick"
      Visible         =   0   'False
      Begin VB.Menu mClickItem 
         Caption         =   "MenuInfo"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mClickItem 
         Caption         =   "Type:"
         Index           =   1
         Begin VB.Menu MenuTypeItem 
            Caption         =   "Edit Text"
            Index           =   0
         End
      End
      Begin VB.Menu mClickItem 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mClickItem 
         Caption         =   "Enabled"
         Index           =   3
      End
      Begin VB.Menu mClickItem 
         Caption         =   "Checked"
         Index           =   4
      End
      Begin VB.Menu mClickItem 
         Caption         =   "Grayed"
         Index           =   5
      End
      Begin VB.Menu mClickItem 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mClickItem 
         Caption         =   "Click MenuItem"
         Index           =   7
      End
      Begin VB.Menu mClickItem 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mClickItem 
         Caption         =   "Delete MenuItem"
         Index           =   9
      End
      Begin VB.Menu mClickItem 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mClickItem 
         Caption         =   "Close Menu"
         Index           =   11
      End
   End
   Begin VB.Menu OwnerMenu 
      Caption         =   "OwnerMenu"
      Visible         =   0   'False
      Begin VB.Menu OwnerMenuItem 
         Caption         =   "Bring Owner to Top"
         Index           =   0
      End
      Begin VB.Menu OwnerMenuItem 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu OwnerMenuItem 
         Caption         =   "Close Menu"
         Index           =   2
      End
   End
End
Attribute VB_Name = "Ac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetMenuItemInfo Lib "USER32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
Private Declare Function ModifyMenu Lib "USER32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Private Declare Function GetMenuState Lib "USER32" (ByVal hMenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long
Private Declare Function AppendMenu Lib "USER32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function CreatePopupMenu Lib "USER32" () As Long
Private Declare Function InsertMenu Lib "USER32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function GetMenuItemRect Lib "USER32" (ByVal Hwnd As Long, ByVal hMenu As Long, ByVal uItem As Long, lprcItem As RECT) As Long
Private Declare Function GetMenuString Lib "USER32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Private Declare Function IsMenu Lib "USER32" (ByVal hMenu As Long) As Long

Public cRefresh  As Boolean, mRefresh As Boolean, MenuE As Boolean

Dim MTf As Mthief

Dim MyHwnd As Long, mCount As Long, Owner As Long, xHwnd As Long

Dim AllTask As Boolean, SysMenu As Boolean, ChildMenu As Boolean
Dim MTfLoaded As Boolean, Search As Boolean, fC As Boolean
Dim EnumAllMenu As Boolean

Dim FindStr As String, tF As String

Dim SysMenuParent As Integer, SysMenuParentCount As Integer


Dim iNode As Nodes

Private smI() As MyMenuItem
Private Sub CheckHwnd_Timer()
    Dim TaskID As Long
    
    If Not cRefresh And Not mRefresh Then
        Call GetWindowThreadProcessId(MyHwnd, TaskID)
    
        If TaskID < 1 And MTfLoaded Then Unload MTf: _
                                            Unload Me
    End If
    
End Sub
Private Sub Form_Activate()
    Call FadeForm(Me, &HC00000, 255)
End Sub
Private Sub Form_Load()
    Dim P As Integer
    
    Call LoadStandardForm(Me, True)
    
    Call SetMP(Tree, True)
    Call SetMP(Label(0), True)
    Call SetMP(Label(2), True)
    
    Set iNode = Tree.Nodes
    
    For P = 0 To 6 Step 2
        ViewItem(P).Checked = View(P)
    Next P
    
    For P = 0 To 4
        MenuItem(P).Checked = Infos(P)
    Next P
    
    mInfos(0).Checked = True
    mInfos(1).Checked = Cmenu(1)
    mInfos(2).Checked = Cmenu(2)
    
End Sub
Public Sub GetHwnd(ByVal Hwnd As Long, ByVal All As Boolean)
    Dim mHwnd As Long, mCount As Long
    Dim P As Integer, X As Integer
    
    If Not cRefresh Then Tree.Visible = False: _
                         Label(0).Visible = False: _
                         Label(2).Visible = False: _
                         Me.Visible = True
        
    Info.Caption = ""
    Label(1).Caption = ""
    Me.Refresh
    
    MyHwnd = Hwnd
    xHwnd = Hwnd
    
    Par.Clear
    iNode.Clear
    Main.Temp.Clear

    If Hwnd < 0 Then
        AllTask = False
        
        Main.AllC.Clear
        
        Call EnumW(0, 3, True)
        
        For P = 0 To Main.AllC.ListCount - 1
            Main.Temp.Clear
            
            Call EnumW(Main.AllC.List(P), 3, False)
            
            If Main.Temp.ListCount = 0 Or xHwnd = -2 Then _
                             Main.Temp.AddItem Main.AllC.List(P)
        
            Call ShowTree(Main.AllC.List(P), False)
            
        Next P
        
        Call WriteInfos(False, True)
        
    Else
        AllTask = All
                
        If Not AllTask Then
            Call EnumW(MyHwnd, 3, False)
        Else
            Call EnumW(0, 1, True)
        End If
        
        Call ShowTree(MyHwnd, AllTask)
        Call WriteInfos(AllTask, False)
    End If
    
    Tree.Visible = True
    
    Label(0).Visible = True
    Label(2).Visible = True
        
End Sub
Private Sub ShowTree(Hwnd As Long, All As Boolean)
    Dim P As Integer, X As Integer, Y As Integer
    Dim ParentItem As Integer
    Dim h As Long
    
    For P = 0 To Main.Temp.ListCount - 1
        Childs.Clear
        
        h = Main.Temp.List(P)
                
        If ViewItem(0).Checked And IsWindowVisible(h) = 0 Or _
           ViewItem(2).Checked And IsWindowVisible(h) = 1 Then _
                                                         GoTo NextP
        
        If ViewItem(6).Checked Then If GetText(h) = "" Then _
                                                         GoTo NextP
        
        Do Until GetParent(h) = 0
            h = GetParent(h)
            Childs.AddItem h, 0
        Loop
        
        Childs.AddItem Main.Temp.List(P)
        
        For X = Childs.ListCount - 1 To 0 Step -1
            If SearchLB(Par.Hwnd, LB_FINDSTRINGEXACT, -1, _
                        Childs.List(X)) > -1 Then Exit For
        Next X
        
        If X < 0 Then
            Call iNode.Add(, tvwLast, , Childs.List(0))
            Par.AddItem Childs.List(0)
            ParentItem = iNode.Count
            X = 1
        Else
            ParentItem = SearchLB(Par.Hwnd, LB_FINDSTRINGEXACT, _
                                            -1, Childs.List(X)) + 1
            X = X + 1
        End If
            
        For Y = X To Childs.ListCount - 1
            Call iNode.Add(ParentItem, tvwChild, , Childs.List(Y))
            ParentItem = iNode.Count
            Par.AddItem Childs.List(Y)
        Next Y
NextP:
    Next P
    
End Sub
Private Sub WriteInfos(AllTask As Boolean, GetChilds As Boolean)
    Dim P As Integer
    Dim h As Long, TaskID As Long
    Dim t As String, V As String
    
    For P = 1 To iNode.Count
        h = iNode(P).Text
        t = "Hwnd: " & h & " - "
        If MenuItem(1).Checked Then t = t & "Class: " & GetClass(h) & " - "
        If MenuItem(2).Checked Then Call GetWindowThreadProcessId(h, TaskID): _
                                    t = t & "TaskID: " & TaskID & " - "
        If MenuItem(3).Checked Then
            V = "No"
            If IsWindowVisible(h) Then V = "Yes"
            t = t & "Visible: " & V & " - "
        End If
        
        If MenuItem(4).Checked Then
            t = t & "Text: " & Chr(34) & GetText(h) & Chr(34)
        Else
            If Right(t, 3) = " - " Then t = Left(t, Len(t) - 3)
        End If
        
        iNode(P).Text = t
    Next P
    
    Main.Temp.Clear
    
    If Not GetChilds And Not Search Then
        If AllTask Then
            Info.Caption = "All Tasks"
            Label(1).Caption = "Found " & iNode.Count & " Controls"
        Else
            If iNode.Count Then
                Info.Caption = "All Childs from " & MyHwnd
                Label(1).Caption = "Found " & iNode.Count - 1 & " Childs"
                Call MItem_Click(1)
            Else
                Info.Caption = "Control " & MyHwnd & " has no Child"
                If Not cRefresh Then Tree.Visible = False: _
                                     Me.Height = 825
            End If
        End If
    Else
        If Not Search Then _
            Info.Caption = "All Tasks with Childs": _
            Label(1).Caption = "Found " & iNode.Count & " Controls/Childs"
    End If
    
    Me.Visible = True
    Tree.Visible = True
    Me.Refresh
    
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Call FormMove(Me)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim P As Integer
    
    On Local Error Resume Next
    
    If MTfLoaded Then Unload MTf
    
    If Me.Tag = "AllMenu" Then Main.ToolMenuItem(MainMenu.AllMenu).Checked = False
    
    If Not MenuE Then
        For P = 0 To 6 Step 2
            View(P) = ViewItem(P).Checked
        Next P
    
        For P = 0 To 4
            Infos(P) = MenuItem(P).Checked
        Next P
    
        If cRefresh Then cRefresh = False: _
                         Cchilds.Show = False: _
                         Cchilds.Stopped = False: _
                         Main.RefreshMenuItem(0).Checked = False
    Else
        If mRefresh Then mRefresh = False: _
                         MyMenu.Show = False: _
                         MyMenu.Stopped = False: _
                         Main.RefreshMenuItem(1).Checked = False

        Cmenu(1) = mInfos(1).Checked
        Cmenu(2) = mInfos(2).Checked
    End If

End Sub
Private Sub Info_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Call FormMove(Me)
End Sub
Private Sub Label_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Si As Integer, P As Integer, Count As Integer
    
    Select Case Index
        Case 0
            If Not MenuE Then
                PopupMenu Menu
            Else
                PopupMenu mMenu
            End If
        Case 1, 3
            If Button = vbLeftButton Then Call FormMove(Me)
        Case 2
            If iNode.Count = 0 Then Exit Sub
            Si = Tree.SelectedItem.Index + 1
            Unload DummyDraw
            Me.Hide
            FindStr = uInput("Type in what to find", "Search", FindStr)
            Count = 0
            If FindStr <> "" Then
                If Si = iNode.Count Then Si = 1
Look:
                For P = Si To iNode.Count
                    If InStr(LCase(iNode(P).Text), LCase(FindStr)) Then _
                            iNode(P).Selected = True: _
                            Exit For
                Next P
                If P > iNode.Count Then
                    Count = Count + 1
                    If Count = 1 Then Si = 1: _
                                    GoTo Look
                    Call uMsg("Cannot find '" & FindStr & "'", _
                              "Not found", vbCritical + vbOKOnly)
                    FindStr = ""
                End If
            End If
Fehler:
            Me.Show
    End Select

End Sub
Private Sub GetMenuStatus(mCount As Long)
    Dim F As Long, E As Long, C As Long, G As Long ', B As Long
    
    E = 0: C = 0: G = 0 ': B = 0

    smI(mCount).AllType = GetMenuState(smI(mCount).Hwnd, _
                                    smI(mCount).ID, MF_BYCOMMAND)
    
    Select Case smI(mCount).AllType
        Case MF_DISABLED
            E = 0
        'Case MF_DISABLED + MF_BITMAP
            'E = 0: B = 1
        Case MF_DISABLED + MF_CHECKED
            E = 0: C = 1
        Case MF_DISABLED + MF_GRAYED
            E = 0: G = 1
        'Case MF_DISABLED + MF_BITMAP + MF_CHECKED
            'E = 0: B = 1: C = 1
        'Case MF_DISABLED + MF_BITMAP + MF_GRAYED
            'E = 0: B = 1: G = 1
        Case MF_DISABLED + MF_CHECKED + MF_GRAYED
            E = 0: C = 1: G = 1
        Case MF_ENABLED
            E = 1
        'Case MF_ENABLED + MF_BITMAP
            'E = 1: B = 1
        Case MF_ENABLED + MF_CHECKED
            E = 1: C = 1
        Case MF_ENABLED + MF_GRAYED
            E = 1: G = 1
        'Case MF_ENABLED + MF_BITMAP + MF_CHECKED
            'E = 1: B = 1: C = 1
        'Case MF_ENABLED + MF_BITMAP + MF_GRAYED
            'E = 1: B = 1: G = 1
        Case MF_ENABLED + MF_CHECKED + MF_GRAYED
            E = 1: C = 1: G = 1
       Case MF_CHECKED
            C = 1
       Case MF_CHECKED + MF_GRAYED
            C = 1: G = 1
       'Case MF_CHECKED + MF_BITMAP + MF_GRAYED
            'C = 1: B = 1: G = 1
        Case MF_GRAYED
            G = 1
        'Case MF_BITMAP
            'B = 1
        'Case MF_BITMAP + MF_GRAYED
            'B = 1: G = 1
        Case Else
            G = -1: E = -1: C = -1 ': B = -1
    End Select
    
    smI(mCount).Enabled = E
    smI(mCount).Checked = C
    smI(mCount).Grayed = G
    
End Sub
Private Sub mClickItem_Click(Index As Integer)
    Dim wMsg As Long
    Dim MItem As Integer

    MItem = Tree.SelectedItem.Index
    
    wMsg = -1

    Select Case LCase(mClickItem(Index).Caption)
        Case "click menuitem"
            wMsg = WM_COMMAND
            If smI(MItem).SysMenu Then wMsg = WM_SYSCOMMAND
            Call PostMessage(smI(MItem).Owner, wMsg, smI(MItem).ID, 0&)
            wMsg = -1
        Case "delete menuitem"
            Call RemoveMenu(smI(MItem).Hwnd, smI(MItem).MItem, _
                            MF_BYPOSITION Or MF_REMOVE)
        Case "enabled"
            wMsg = MF_ENABLED
            If mClickItem(Index).Checked Then wMsg = MF_DISABLED
        Case "checked"
            wMsg = MF_CHECKED
            If mClickItem(Index).Checked Then wMsg = MF_UNCHECKED
        Case "grayed"
            wMsg = MF_GRAYED
            If mClickItem(Index).Checked Then wMsg = MF_ENABLED
    End Select
    
    If wMsg > -1 Then
        Call ModifyMenu(smI(MItem).Hwnd, smI(MItem).ID, wMsg, _
                                     smI(MItem).ID, smI(MItem).Text)
        Call MenuEnum(MyHwnd, SysMenu, ChildMenu)
    End If

End Sub
Private Sub MenuItem_Click(Index As Integer)
    MenuItem(Index).Checked = Not (MenuItem(Index).Checked)
    
    If tF <> "" Then
        Call Me.SearchHwnd(tF, fC)
    Else
        Call GetHwnd(MyHwnd, AllTask)
    End If
    
End Sub
Private Sub MenuTypeItem_Click(Index As Integer)
    Dim Answer As String, t As String
    Dim MItem As Long, wFlags As Long
    
    MItem = Tree.SelectedItem.Index
    
    wFlags = -1
    
    Select Case Index
        Case 0
            t = "Gebe den neuen Menutext für '" & _
                 smI(MItem).Text & "' ein"
            Answer = uInput(t, "Edit Menutext")
            If Answer <> "" Then smI(MItem).Text = Answer: _
                                 wFlags = MF_STRING
    End Select
    
    If wFlags > -1 Then
        Call ModifyMenu(smI(MItem).Hwnd, smI(MItem).ID, wFlags, _
                                    smI(MItem).ID, smI(MItem).Text)
        Call MenuEnum(MyHwnd, SysMenu, ChildMenu)
    End If

End Sub
Private Sub mInfos_Click(Index As Integer)
    mInfos(Index).Checked = Not (mInfos(Index).Checked)
    Call WriteMenuInfos
End Sub
Private Sub MItem_Click(Index As Integer)
    Dim i As Integer, selItem As Integer
    
    On Local Error Resume Next
    
    selItem = Tree.SelectedItem.Index
    If selItem < 1 Then selItem = 1
    
    Select Case Index
        Case 0
            Call GetHwnd(MyHwnd, AllTask)
        Case 1
            For i = 1 To iNode.Count
                iNode(i).Expanded = True
            Next i
        Case 6
            Unload Me
    End Select
    
    iNode(selItem).Selected = True

End Sub
Private Sub mMenuItem_Click(Index As Integer)

    Select Case LCase(mMenuItem(Index).Caption)
        Case "refresh"
            If EnumAllMenu Then
                Call EnumAllWindowsMenu
            Else
                Call MenuEnum(MyHwnd, SysMenu, ChildMenu)
            End If
        Case "expand all"
            Call MItem_Click(1)
    End Select

End Sub
Private Sub OwnerMenuItem_Click(Index As Integer)
    Dim TreeItem As Integer
    
    TreeItem = Tree.SelectedItem.Index
    
    Select Case Index
        Case 0
            Call WinRepaint(smI(TreeItem).Owner, True)
    End Select
    
End Sub
Private Sub Tree_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim Z As Integer, TreeItem As Integer, TreeParent As Integer
    Dim t As String, TreeEndItem As Long, TmpParent As Long
    Dim Item As Long, lHwnd As Long
    Dim Mt As New Mthief
    
    On Local Error GoTo Fehler
    
    TreeItem = Tree.SelectedItem.Index
    
    If xHwnd = -2 Then
        If Not MTfLoaded Then
            Load Mt
            Set MTf = Mt
            MTfLoaded = True
        Else
            Call RemoveThiefMenu
        End If
        
        On Local Error Resume Next
        
        TreeParent = TreeItem
        
        Err.Clear
        Do Until Err.Number = 91
            TreeParent = iNode(TreeParent).Parent.Index
        Loop
        
        Err.Clear
        TreeEndItem = TreeParent
        Do Until Err.Number = 91 Or TreeEndItem = iNode.Count
            TreeEndItem = TreeEndItem + 1
            X = iNode(TreeEndItem).Parent.Index
        Loop
            
        MTf.mID.Clear
        MTf.Owner.Clear
        MTf.SysID.Clear
        
        TmpParent = TreeParent
        If Not smI(TreeParent).SysMenu Then _
                                        TreeParent = TreeParent + 1
        
        For Z = TreeParent To TreeEndItem - 1
            MTf.mID.AddItem smI(Z).ID
            MTf.Owner.AddItem smI(Z).Owner
            SysMenuParent = Z - 1

            SysMenu = smI(Z).SysMenu
            If SysMenu Then MTf.SysID.AddItem smI(Z).ID
            
            If SysMenu Then
                If Z = TmpParent Then
                    Call CreateNewMenu(Z)
                Else
                    Call CreateMenuItem(Z)
                End If
            Else
                If iNode(Z).Parent.Index = TmpParent Then
                    Call CreateNewMenu(Z)
                Else
                    Call CreateMenuItem(Z)
                End If
            End If
        Next Z
        
        If SysMenu Then
            MTf.Caption = My.Name
        Else
            MTf.Caption = "Menu from " & smI(TreeItem).Owner
        End If
        
        MTf.Show
        Call ShowWindow(MTf.Hwnd, SW_SHOWNOACTIVATE)
    
        If Tree.Visible And smI(TreeItem).ID = -1 Then _
                                            PopupMenu OwnerMenu
    End If


    If Not MenuE Then
        t = Node.Text

        If t <> "" Then
            t = mID(t, InStr(t, " ") + 1)
            Z = InStr(t, " ")
            If Z Then t = mID(t, 1, Z - 1)
            Call Dummy.PopUp(Val(t))
        End If
    Else
        If smI(TreeItem).ID > -1 Then
            If smI(TreeItem).Type <> MF_SEPARATOR Then
                mClickItem(1).Caption = "Type: String"
            
                MenuTypeItem(0).Enabled = True
            Else
                mClickItem(1).Caption = "Type: Separator"
                
                MenuTypeItem(0).Enabled = False
            End If
            
            If smI(TreeItem).Type <> MF_SEPARATOR Then
                mClickItem(3).Enabled = True
                mClickItem(4).Enabled = True
                mClickItem(5).Enabled = True
                
                mClickItem(3).Checked = smI(TreeItem).Enabled
                mClickItem(4).Checked = smI(TreeItem).Checked
                mClickItem(5).Checked = smI(TreeItem).Grayed
            Else
                For Z = 3 To 5
                    mClickItem(Z).Enabled = False
                    mClickItem(Z).Checked = False
                Next Z
            End If
            
            PopupMenu mClick
        End If
    End If
    
Fehler:
End Sub
Private Sub ViewItem_Click(Index As Integer)
    Dim P As Integer
    
    If Index < 6 Then
        If Not ViewItem(Index).Checked Then
            ViewItem(0).Checked = False
            ViewItem(2).Checked = False
            ViewItem(4).Checked = False
            ViewItem(Index).Checked = True
        Else
            ViewItem(Index).Checked = False
            
            If Not ViewItem(0).Checked And _
               Not ViewItem(2).Checked And _
               Not ViewItem(4).Checked Then
                    ViewItem(Index).Checked = True
                    Exit Sub
            End If
        End If
    Else
        ViewItem(6).Checked = Not (ViewItem(6).Checked)
    End If
    
    Call GetHwnd(MyHwnd, AllTask)

End Sub
Public Sub AutoRefresh(Auto As Boolean)
    
    If Not MenuE Then
        cRefresh = True
        Cchilds.Show = True
        Me.Tag = "cRefresh"
        Me.Visible = True
    Else
        mRefresh = True
        MyMenu.Show = True
        Me.Tag = "mRefresh"
    End If
    
    If Auto Then
        If Not MenuE Then
            Cchilds.Stopped = False
            Me.Caption = "All Childs - AutoRefresh (F11 to stop)"
        Else
            MyMenu.Stopped = False
            Me.Caption = "Menu - AutoRefresh (F10 to stop)"
        End If
    Else
        If Not MenuE Then
            Cchilds.Stopped = True
            Me.Caption = "All Childs - AutoRefresh (F11 to start)"
        Else
            MyMenu.Stopped = True
            Me.Caption = "Menu - AutoRefresh (F10 to start)"
        End If
    End If
        
End Sub
Public Sub MenuEnum(Hwnd As Long, GetSysMenu As Boolean, _
                                                  Childs As Boolean)
    Dim mHwnd As Long, h As Long, mCount As Long
    Dim P As Integer, selItem As Integer, X As Integer
    Dim mMaxHeight As Integer
    Dim Cap As String
    Dim Mt As New Mthief
    Dim Mr As RECT, Wr As RECT
    
    If Not MTfLoaded Then
            Load Mt
            Set MTf = Mt
            MTfLoaded = True
            If mRefresh Then Mt.mRefresh = True
        Else
            If xHwnd <> -2 Then MTf.SysID.Clear: _
                                Call RemoveThiefMenu
    End If

    
    If xHwnd <> -2 Then iNode.Clear: _
                        SysID.Clear: _
                        If Not MTf.Visible Then MTf.Visible = True
    
    ChildMenu = Childs
    SysMenu = GetSysMenu
    
    MenuE = True
    MyHwnd = Hwnd
                        
    If iNode.Count > 0 Then
        selItem = Tree.SelectedItem.Index
    Else
        selItem = 1
    End If
    
    If Not Label(0).Visible Then
        For P = 0 To Label.UBound
            Label(P).Visible = True
        Next P
    End If
    
    mCount = 0
    If xHwnd <> -2 Then SysMenuParentCount = -1: _
                        ReDim smI(0)
        
    Owner = MyHwnd
        
    If ChildMenu Then
        SMenu GetMenu(MyHwnd), Tree
        Cap = "All Menu from all Childs from " & MyHwnd
        MTf.Caption = Cap
    Else
        If SysMenu Then
            Call AddSysMenu(MyHwnd, True)
            SMenu GetSystemMenu(MyHwnd, False), Tree, "", True
            Cap = "SystemMenu from " & MyHwnd
            MTf.Caption = Cap
        Else
            If xHwnd <> -2 Then
                SMenu GetMenu(MyHwnd), Tree
                Cap = "Menu from " & MyHwnd
                MTf.Caption = Cap
            Else
                Call AddSysMenu(MyHwnd, False)
                SMenu GetMenu(MyHwnd), Tree, "", True
            End If
        End If
    End If

    If ChildMenu Then
        Main.Temp.Clear
        Call EnumW(MyHwnd, 3, False)

        For P = 0 To Main.Temp.ListCount - 1
            Owner = CLng(Main.Temp.List(P))
            
            If GetMenu(Owner) > 0 Then SMenu GetMenu(Owner), Tree
            
            If GetSystemMenu((Owner), False) > 0 Then
                Call AddSysMenu(Owner, True)
                SysMenu = True
                SMenu GetSystemMenu(Owner, False), Tree, "", True
                SysMenu = False
            End If
        Next P
        Me.Visible = True
    End If
    
    Info.Caption = Cap
        
    Label(1).Caption = iNode.Count & " Menuentries found"
    
    If iNode.Count < selItem Then selItem = 1
    If iNode.Count > 0 Then iNode(selItem).Selected = True
    
    If xHwnd <> -2 Then
        Call WriteMenuInfos
        
        Call GetWindowRect(MTf.Hwnd, Wr)
    
        mMaxHeight = 0

        h = GetMenu(MTf.Hwnd)
        mCount = GetMenuItemCount(h)

        For P = 1 To mCount
            Call GetMenuItemRect(MTf.Hwnd, h, CLng(P), Mr)
              
            If Mr.Bottom > mMaxHeight Then mMaxHeight = Mr.Bottom
        Next P
    
        If mMaxHeight > 0 Then
            MTf.Message.Visible = False
            MTf.Height = MTf.Height + _
                         ((15 * (mMaxHeight - Wr.Bottom)) + 60)
        Else
            MTf.Height = 675
            MTf.Message.Visible = True
        End If
    End If
    
    If xHwnd <> -2 Then CheckHwnd.Enabled = True
    
End Sub
Private Sub RemoveThiefMenu()
    Dim mHwnd As Long, P As Long
            
    mHwnd = GetMenu(MTf.Hwnd)
    
    Do Until GetMenuItemCount(mHwnd) = 0
        For P = 0 To GetMenuItemCount(mHwnd)
            Call RemoveMenu(mHwnd, P, MF_BYPOSITION Or MF_REMOVE)
        Next P
    Loop

End Sub
Private Sub WriteMenuInfos()
    Dim P As Integer, X As Integer
    Dim R$
    
    On Local Error Resume Next
    
    For P = 1 To iNode.Count

        R$ = smI(P).Text
        
        If smI(P).Type = MF_SEPARATOR Then R$ = "--- Separator ---"
        
        If mInfos(1).Checked Then R$ = R$ & " - ID: " & smI(P).ID
        If mInfos(2).Checked Then R$ = R$ & " - mHwnd: " & _
                                                        smI(P).Hwnd

        If xHwnd = -2 Then
            Err.Clear
            X = iNode(P).Parent.Index
            If Err.Number = 91 Then R$ = R$ & " - '" & _
                                         GetText(smI(P).Owner) & "'"
        End If
        
        iNode(P).Text = R$
    Next P
    
End Sub
Private Function GetPopupMenuString(mHwnd As Long, SubItem As Long, _
                                    MItem As Long) As String
    Dim menusX As MENUITEMINFO
    Dim str As String
        
    menusX.cbSize = Len(menusX)
    menusX.fMask = MIIM_TYPE
    menusX.dwTypeData = Space(255)
    menusX.cch = 255
    
    GetMenuItemInfo mHwnd, SubItem, True, menusX
    smI(MItem).Type = menusX.fType
     
    Select Case menusX.fType
        Case MF_STRING, MF_RIGHTJUSTIFY
            menusX.dwTypeData = Trim(menusX.dwTypeData)
            GetPopupMenuString = menusX.dwTypeData
        Case MF_SEPARATOR
            GetPopupMenuString = ""
    End Select
    
End Function
Private Sub SMenu(mHwnd As Long, Tree As TreeView, _
                                   Optional tmpKey As String, _
                                   Optional iSubFlag As Boolean)
    
    Dim n As Long, C As Long, i As Long, h As Long, Result As Long
    Dim iNode As Node
    Dim menusX As MENUITEMINFO
    Dim Temp As String, Buffer As String
    
    Static iKey As Integer
        
    'On Local Error Resume Next
    
    n = GetMenuItemCount(mHwnd)
    
    For i = 0 To n - 1
        ReDim Preserve smI(UBound(smI) + 1)

        Buffer = Space$(128)
    
        C = GetMenuItemID(mHwnd, i)

        smI(UBound(smI)).Owner = Owner
        smI(UBound(smI)).Hwnd = mHwnd
        smI(UBound(smI)).ID = C
        smI(UBound(smI)).MItem = i
        smI(UBound(smI)).SysMenu = False
        
        If xHwnd <> -2 Then
            MTf.mID.AddItem smI(UBound(smI)).ID
            MTf.Owner.AddItem smI(UBound(smI)).Owner
        
            If SysMenu Then smI(UBound(smI)).SysMenu = True: _
                            MTf.SysID.AddItem smI(UBound(smI)).ID
        End If
        
        Call GetMenuStatus(UBound(smI))
        
        Result = GetMenuString(mHwnd, C, Buffer, Len(Buffer), _
                                                       MF_BYCOMMAND)
        Buffer = Left$(Buffer, Result)
        
        If Buffer = "" Or smI(UBound(smI)).ID = -1 Then
            Buffer = GetPopupMenuString(mHwnd, i, UBound(smI))
            C = iKey
            iKey = iKey + 1
        Else
            C = C + 15000
        End If

        Buffer = Replace(Buffer, "&", "")
        If InStr(Buffer, Chr(0)) Then _
                Buffer = mID(Buffer, 1, InStr(Buffer, Chr(0)) - 1)
        If InStr(Buffer, Chr(9)) Then _
                Buffer = mID(Buffer, 1, InStr(Buffer, Chr(9)) - 1)

        smI(UBound(smI)).Text = Buffer
        
        If Not iSubFlag Then
            If tmpKey <> "" Then
                Set iNode = Tree.Nodes.Add(tmpKey, tvwNext, , "")
            Else
                Set iNode = Tree.Nodes.Add(, tvwLast, , "")
            End If
            
            Call CreateNewMenu(CInt(UBound(smI)))
            
            If GetSubMenu(mHwnd, i) > 1 Then
                iSubFlag = True
                iNode.Key = "k" & CStr(C)
                tmpKey = iNode.Key
                SMenu GetSubMenu(mHwnd, i), Tree, tmpKey, True
                iSubFlag = False
            End If
        Else
            If SysMenu Then
                tmpKey = "SysMenu" & SysMenuParentCount
                smI(UBound(smI)).SysMenu = True
                Set iNode = Tree.Nodes.Add(tmpKey, tvwChild, "", "")
            Else
                If xHwnd = -2 Then
                    If tmpKey = "" Then tmpKey = "Menu" & SysMenuParentCount
                    Set iNode = Tree.Nodes.Add(tmpKey, tvwChild, "", "")
                Else
                    Set iNode = Tree.Nodes.Add(tmpKey, tvwChild, _
                                               "k" & CStr(C), "")
                End If
            End If
                        
            If xHwnd <> -2 Then Call CreateMenuItem(CInt(UBound(smI)))
            
            If GetSubMenu(mHwnd, i) > 1 Then
                iSubFlag = True
                iNode.Key = "k" & CStr(C)
                Temp = tmpKey
                tmpKey = iNode.Key
                SMenu GetSubMenu(mHwnd, i), Tree, tmpKey, True
                tmpKey = Temp
            End If
        End If
   Next i
    
End Sub
Private Sub CreateMenuItem(Index As Integer)
    Dim mParent As Long, NewMenu As Long
    Dim pID As Long
      
    If SysMenu Then
        mParent = smI(SysMenuParent).tHwnd
        pID = smI(SysMenuParent).ID
    Else
        If xHwnd = -2 Then
            mParent = smI(SysMenuParent).tHwnd
        Else
            mParent = smI(iNode(UBound(smI)).Parent.Index).tHwnd
        End If
        pID = smI(iNode(UBound(smI)).Parent.Index).ID
    End If
    
    If smI(Index).ID = -1 Or pID = -1 Then
        If smI(Index).ID = -1 Then
            NewMenu = CreatePopupMenu
            Call AppendMenu(mParent, MF_POPUP, NewMenu, _
                                                    smI(Index).Text)
        Else
            Call AppendMenu(mParent, smI(Index).AllType, _
                                     smI(Index).ID, smI(Index).Text)
        End If
    Else
        Call AppendMenu(mParent, smI(Index).AllType, _
                                 smI(Index).ID, smI(Index).Text)
    End If
    
    smI(Index).tHwnd = GetLastMenuHwnd
    
End Sub
Private Function GetLastMenuHwnd() As Long
    Dim mCount As Long, h As Long
    
    h = GetMenu(MTf.Hwnd)
    mCount = GetMenuItemCount(h)
    
    Do Until GetSubMenu(h, mCount - 1) = 0
        h = GetSubMenu(h, mCount - 1)
        mCount = GetMenuItemCount(h)
    Loop
                
    GetLastMenuHwnd = h
    
End Function
Private Sub CreateNewMenu(Index As Integer)
    Dim NewMenu As Long, FndWindow As Long, FndWindowMenu As Long
        
    NewMenu = CreatePopupMenu
    FndWindow = MTf.Hwnd
    FndWindowMenu = GetMenu(FndWindow)
    
    Call AppendMenu(FndWindowMenu, MF_POPUP, NewMenu, smI(Index).Text)
    
    smI(Index).tHwnd = GetLastMenuHwnd
    
End Sub
Private Sub AddSysMenu(ByVal Hwnd As Long, isSysMenu As Boolean)
    Dim Text As String, Key As String
    
    Text = "Menu"
    If isSysMenu Then Text = "SystemMenu"
    
    If Hwnd <> -1 Then Text = Text & " from " & Hwnd
    
    SysMenuParentCount = SysMenuParentCount + 1
    
    Key = "Menu"
    If isSysMenu Then Key = "SysMenu"
    Key = Key & SysMenuParentCount
    
    Call iNode.Add(, tvwLast, Key, Text)
    
    ReDim Preserve smI(UBound(smI) + 1)
    
    SysMenuParent = UBound(smI)

    smI(UBound(smI)).Owner = Hwnd
    smI(UBound(smI)).Hwnd = 0
    smI(UBound(smI)).ID = -1
    smI(UBound(smI)).MItem = 1
    smI(UBound(smI)).SysMenu = isSysMenu
    smI(UBound(smI)).Text = Text
    
    smI(UBound(smI)).Enabled = True
    smI(UBound(smI)).Checked = False
    smI(UBound(smI)).Grayed = False

    If xHwnd <> -2 Then Call CreateNewMenu(UBound(smI))

End Sub
Public Sub SearchHwnd(Find As String, FindClass As Boolean)
    Dim X As Integer, P As Integer
    Dim t As String
    
    X = 0
    Search = True
    fC = FindClass
    
    Find = LCase(Find)
    tF = Find
    
    Call EnumCompleteWindows(Childs)
    
    Do Until X > Childs.ListCount - 1
        If FindClass Then
            If InStr(LCase(GetClass(Childs.List(X))), Find) = 0 Then _
                        Childs.RemoveItem X: _
                        X = X - 1
        Else
            If InStr(LCase(GetText(Childs.List(X))), Find) = 0 Then _
                        Childs.RemoveItem X: _
                        X = X - 1
        End If
                           
        X = X + 1
    Loop
    
    iNode.Clear
    
    For P = 0 To Childs.ListCount - 1
        Call iNode.Add(, tvwLast, , Childs.List(P))
    Next P

    Call WriteInfos(False, False)
    
    Childs.Clear
    Main.Temp.Clear
    Main.AllC.Clear
    
    Label(0).Visible = True
    Label(1).Caption = "Found " & iNode.Count & " Controls/Childs"
    Label(2).Visible = True

    t = "Searched Text"
    If FindClass Then t = "Searched Class"
    Me.Info.Caption = t & " = '" & Find & "'"

    Me.Visible = True
    
End Sub
Public Sub EnumAllWindowsMenu()
    Dim X As Integer
    
    Tree.Visible = False
    
    EnumAllMenu = True
    iNode.Clear
    
    Call EnumCompleteWindows(Childs)

    Main.Temp.Clear
    
    For X = 0 To Childs.ListCount - 1
        If GetMenuItemCount(GetMenu(Childs.List(X))) > 0 _
        Or GetMenuItemCount(GetSystemMenu(Childs.List(X), False)) > 0 Then _
                Main.Temp.AddItem Childs.List(X)
    Next X
    Childs.Clear
    
    For X = 0 To Main.Temp.ListCount - 1
        Childs.AddItem Main.Temp.List(X)
    Next
    Main.Temp.Clear
    
    ReDim smI(0)
    SysMenuParentCount = -1
    
    xHwnd = -2
            
    For X = 0 To Childs.ListCount - 1
        If GetMenuItemCount(GetMenu(Childs.List(X))) > 0 Then _
                    Call MenuEnum(Childs.List(X), False, False)
        If GetMenuItemCount(GetSystemMenu(Childs.List(X), False)) > 0 Then _
                    Call MenuEnum(Childs.List(X), True, False)
    Next X
            
    Call WriteMenuInfos
            
    Info.Caption = "All Menus in System"
    
    Label(0).Visible = True
    Label(2).Visible = True
    Label(1).Caption = iNode.Count & " Menuentries found"
            
    Call Tree_NodeClick(iNode(1))
            
    Tree.Visible = True
    Me.Visible = True
    
End Sub
