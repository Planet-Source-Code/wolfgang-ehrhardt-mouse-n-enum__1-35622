VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form EnumC 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   4  'Festes Werkzeugfenster
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command 
      Caption         =   "Refresh"
      Height          =   255
      Index           =   5
      Left            =   2040
      TabIndex        =   9
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "NoValue"
      Height          =   255
      Index           =   4
      Left            =   3000
      TabIndex        =   8
      Top             =   3360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Caption         =   "Get Value"
      Height          =   255
      Index           =   3
      Left            =   3000
      TabIndex        =   7
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Caption         =   "Close"
      Height          =   255
      Index           =   2
      Left            =   5520
      TabIndex        =   6
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton Command 
      Caption         =   "Clibboard"
      Height          =   255
      Index           =   1
      Left            =   4200
      TabIndex        =   5
      Top             =   3360
      Width           =   1215
   End
   Begin VB.ListBox W 
      Height          =   255
      Left            =   4440
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox E 
      Height          =   255
      Left            =   5160
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox cClass 
      Height          =   255
      Left            =   5160
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox cHwnd 
      Height          =   255
      Left            =   4440
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   3135
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   5530
      _Version        =   393217
      BackColor       =   4210752
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      MousePointer    =   3
      DisableNoScroll =   -1  'True
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Enum.frx":0000
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu MenuItem 
         Caption         =   "Parent"
         Index           =   0
      End
      Begin VB.Menu MenuItem 
         Caption         =   "Childs"
         Index           =   1
      End
      Begin VB.Menu MenuItem 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu MenuItem 
         Caption         =   "Close Menu"
         Index           =   3
      End
   End
   Begin VB.Menu Cmenu 
      Caption         =   "Cmenu"
      Visible         =   0   'False
      Begin VB.Menu CmenuItem 
         Caption         =   "Copy all"
         Index           =   0
      End
      Begin VB.Menu CmenuItem 
         Caption         =   "Copy only Source"
         Index           =   1
      End
      Begin VB.Menu CmenuItem 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu CmenuItem 
         Caption         =   "Close Menu"
         Index           =   3
      End
   End
End
Attribute VB_Name = "EnumC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Found As Boolean
Dim Source As String
Dim MyControl As Long, MyParent As Long, MyChild As Long
Private Sub CmenuItem_Click(Index As Integer)
    Select Case Index
        Case 0
            Clipboard.SetText Text1.Text
        Case 1
            Clipboard.SetText Source
    End Select
End Sub
Private Sub Command_Click(Index As Integer)
    Select Case Index
        Case 1
            PopupMenu Cmenu
        Case 2
            Unload Me
        Case 3
            Command(3).Visible = False
            Command(4).Visible = True
            Call ShowLB(W)
        Case 4
            Command(3).Visible = True
            Command(4).Visible = False
            Call ShowLB(E)
        Case 5
            Call GetChild(cHwnd.List(cHwnd.ListCount - 1))
    End Select
End Sub
Private Sub Form_Activate()
    Me.Refresh
    Text1.Refresh
End Sub
Private Sub Form_Load()
    Dim P As Integer
    
    Call LoadStandardForm(Me, False)
    
    Me.Top = Main.Top
    Me.Left = Main.Left
    
    For P = Command.LBound To Command.UBound
        Call SetMP(Command(P), True)
    Next P
    
End Sub
Public Sub GetChild(ByVal Child As Long)
    Dim P As Integer
    Dim Class As String, Text As String
    Dim TaskID As Long
    
    MyChild = Child
    Found = False
    
    cHwnd.Clear
    cClass.Clear
    
    E.Clear
    W.Clear
    
    Class = GetClass(Child)
    Me.Caption = "Enumierung von Control " & _
                 Child & " (" & Class & ")"
    
    cHwnd.AddItem Child
    cClass.AddItem Class
    
    Call GetWindowThreadProcessId(Child, TaskID)
    
    If TaskID = 0 Then
        E.AddItem vbRed
        E.AddItem "Control " & Child & " is unloaded => TaskID = 0"
        W.AddItem E.List(0)
        W.AddItem E.List(1)
        Command(1).Enabled = False
        Command(3).Enabled = False
        Command(4).Enabled = False
        Command(5).Enabled = False
        GoTo WriteIt
    End If
    
    MyControl = Child
    
    E.AddItem vbYellow
    E.AddItem GetCinfo(Child)
    E.AddItem vbWhite
    E.AddItem "Enumierung" & vbCrLf & "--------------------"
    
    For P = 0 To 3
        W.AddItem E.List(P)
    Next P
    
    Do Until GetParent(Child) = 0
        Child = GetParent(Child)
        cHwnd.AddItem Child, 0
        cClass.AddItem GetClass(Child), 0
    Loop
    MyParent = Child
    
    CmenuItem(1).Enabled = True
    
    For P = 0 To cHwnd.ListCount - 1
        If Not FindHwnd(P) Then Exit For
    Next P
    
WriteIt:
    Call ShowLB(E)
    
    Me.Show
    
End Sub
Private Function FindHwnd(ByVal Index As Integer) As Boolean
    Dim Class As String, Class2Find As String
    Dim t As String, Com As String, U As String, R$
    Dim Child As Long, Child2Find As Long, Parent As Long
        
    Child2Find = cHwnd.List(Index)
    Class2Find = cClass.List(Index)
        
    If Index Then
        Parent = cHwnd.List(Index - 1)
        
        Com = "Child"
        If Index < 2 Then Com = "Parent"
        
        R$ = "Child = FindChildByClass(" & Com & _
             ", " & Chr(34) & Class2Find & Chr(34) & ")"
        Source = Source & R$ & vbCrLf
        E.AddItem R$
        Child = FindChildByClass(Parent, Class2Find)
        W.AddItem "FindChildByClass(" & Parent & ", " & _
                  Chr(34) & Class2Find & Chr(34) & ") = " & _
                                                            Child
        t = "Child"
    Else
        Com = "FindWindow(" & Chr(34) & Class2Find & _
                                        Chr(34) & ", vbNullString)"
        
        R$ = "Parent = " & Com
        Source = Source & R$ & vbCrLf
        E.AddItem R$
        Child = FindWindow(Class2Find, vbNullString)
        W.AddItem Com & " = " & Child
        t = "Parent"
    End If
    
    U = t
    Do While (Child <> Child2Find) And Child
        R$ = U & " = GetWindow(" & U & ", GW_HWNDNEXT)"
        Source = Source & R$ & vbCrLf
        E.AddItem R$
        t = "GetWindow(" & Child & ", GW_HWNDNEXT) = "
        Child = GetWindow(Child, GW_HWNDNEXT)
        W.AddItem t & Child
    Loop
    
Done:
    Class = "Control " & str(MyControl) & " "
    
    If Child = 0 Then
        CmenuItem(1).Enabled = False
        FindHwnd = False
        Found = False
        t = vbCrLf & "   => " & Class & "konnte nicht mit der" & vbCrLf & _
                     "        FindChildByClass-Methode enumiert werden" & vbCrLf
        E.AddItem vbRed
        E.AddItem t
        W.AddItem vbRed
        W.AddItem t
        Win2Find = MyChild
        If EnumW(MyParent, 2, False) Then
            E.AddItem vbGreen
            W.AddItem vbGreen
            t = "   => " & Class & _
                "wurde mit der WndEnumChildProc-Methode" & _
                vbCrLf & "        erfolgreich enumiert"
            E.AddItem t
            W.AddItem t
        Else
            E.AddItem vbRed
            E.AddItem vbRed
            t = "   => " & Class & _
                "konnte nicht mit der WndEnumChildProc-Methode" & _
                vbCrLf & "        enumiert werden" & vbCrLf
            E.AddItem t
            W.AddItem t
            
            Win2Find = MyChild
            If EnumW(0, 2, True) Then
                E.AddItem vbGreen
                W.AddItem vbGreen
                t = "   => " & Class & _
                           "wurde mit der WndEnumProc-Methode" & _
                           vbCrLf & "        erfolgreich enumiert"
                E.AddItem t
                W.AddItem t
            Else
                E.AddItem vbRed
                E.AddItem vbRed
                t = "   => " & Class & _
                    "konnte nicht mit der WndEnumProc-Methode" & _
                    vbCrLf & "        enumiert werden"
                E.AddItem t
                W.AddItem t
            End If
        End If
    Else
        FindHwnd = True
        Found = True
        If Index = cHwnd.ListCount - 1 Then
            t = vbCrLf & "   => " & Class & _
                         "wurde erfolgreich mit der " & vbCrLf & _
                         "        FindChildByClass-Methode enumiert "
            E.AddItem vbGreen
            E.AddItem t
            W.AddItem vbGreen
            W.AddItem t
        End If
    End If
    
End Function
Private Sub ShowLB(LB As ListBox)
    Call WriteLB(LB, Me.Text1)
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Call FormMove(Me)
End Sub
Private Sub MenuItem_Click(Index As Integer)
    
    Select Case LCase(MenuItem(Index).Caption)
        Case "parent"
            Call LEnum(MyControl, False, False)
        Case "childs"
            Call LEnum(MyControl, False, True)
    End Select
    
End Sub
