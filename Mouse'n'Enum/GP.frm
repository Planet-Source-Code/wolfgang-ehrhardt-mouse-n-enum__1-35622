VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form GP 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Parentenumierung"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ListBox LB 
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin RichTextLib.RichTextBox Text 
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3625
      _Version        =   393217
      BackColor       =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      MousePointer    =   3
      Appearance      =   0
      TextRTF         =   $"GP.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command 
      Caption         =   "Clipboard"
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Caption         =   "Close"
      Height          =   255
      Index           =   0
      Left            =   2640
      TabIndex        =   0
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label cLabel 
      Alignment       =   1  'Rechts
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enum"
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
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   585
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu MenuItem 
         Caption         =   "Co"
         Index           =   0
      End
      Begin VB.Menu MenuItem 
         Caption         =   "Ci"
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
End
Attribute VB_Name = "GP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyHwnd As Long
Private Sub cLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then PopupMenu Menu
End Sub
Private Sub Command_Click(Index As Integer)
    Select Case Index
        Case 0
            Unload Me
        Case 1
            Clipboard.SetText Text.Text
    End Select
End Sub
Private Sub Form_Load()
    Call LoadStandardForm(Me, False)
    Call SetMP(Command(0), True)
    Call SetMP(Command(1), True)
    Call SetMP(cLabel, True)
End Sub
Public Sub GetP(ByVal Child As Long)
    Dim Xc As Long, Sx As Long, Task As Long
    Dim T As String
    
    MyHwnd = Child
    
    MenuItem(0).Caption = "Control " & MyHwnd
    MenuItem(1).Caption = "All Childs from " & MyHwnd
    
    LB.Clear
    
    Call GetWindowThreadProcessId(Child, TaskID)
    
    T = GetText(Child)
    If T <> "" Then
        T = "Text =  " & Chr(34) & T & Chr(34)
    Else
        T = "Child has no Text"
    End If
    
    LB.AddItem vbYellow
    LB.AddItem GetCinfo(Child)
    LB.AddItem vbWhite
    LB.AddItem "Enumierung" & vbCrLf & _
               "--------------------"
    
    If GetParent(Child) = 0 Then
        LB.AddItem vbRed
        LB.AddItem "Control " & Child & " is Parent" & vbCrLf
        LB.AddItem vbGreen
        LB.AddItem "GetParent(" & Child & ") = 0"
        T = GetClass(Child)
        Child = FindWindow(T, vbNullString)
        LB.AddItem "FindWindow(" & Chr(34) & T & Chr(34) & ", vbNullString)" & _
                   " = " & Child
        GoTo WriteIt
    End If
    
    Xc = Child
    Sx = Xc

    Do Until GetParent(Child) = 0
        Child = GetParent(Child)
        LB.AddItem "   - GetParent(" & Xc & ") = " & Child
        Xc = Child
    Loop
    
    Call GetWindowThreadProcessId(Child, TaskID)
    
    T = GetText(Child)
    If T <> "" Then
        T = "Text =  " & Chr(34) & T & Chr(34)
    Else
        T = "Child has no Text"
    End If

    
    LB.AddItem vbGreen
    LB.AddItem vbCrLf & "Parentinfo" & vbCrLf & _
               "-----------------" & vbCrLf & _
               "   - Hwnd = " & Child & vbCrLf & _
               "   - Classname = " & GetClass(Child) & vbCrLf & _
               "   - TaskID = " & TaskID & vbCrLf & _
               "   - " & T & vbCrLf

WriteIt:
    Call WriteLB(LB, Text)
    
    Me.Visible = True
    
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Call FormMove(Me)
End Sub
Private Sub MenuItem_Click(Index As Integer)
    Select Case Index
        Case 0
            Call LEnum(MyHwnd, True, False)
        Case 1
            Call LEnum(MyHwnd, False, True)
    End Select
End Sub
