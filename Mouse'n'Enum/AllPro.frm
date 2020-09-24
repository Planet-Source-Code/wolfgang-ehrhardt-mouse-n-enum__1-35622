VERSION 5.00
Begin VB.Form AllPro 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Form1"
   ClientHeight    =   2025
   ClientLeft      =   150
   ClientTop       =   420
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command 
      Caption         =   "SnapShot"
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin VB.ListBox lbID 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command 
      Caption         =   "Refresh"
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   1
      Top             =   1680
      Width           =   1095
   End
   Begin VB.ListBox LB 
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1425
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu MenuItem 
         Caption         =   "Open Path in Explorer"
         Index           =   0
      End
      Begin VB.Menu MenuItem 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu MenuItem 
         Caption         =   "Terminate Process"
         Index           =   2
      End
      Begin VB.Menu MenuItem 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu MenuItem 
         Caption         =   "Close Menu"
         Index           =   4
      End
   End
   Begin VB.Menu SSmenu 
      Caption         =   "SnapShot"
      Visible         =   0   'False
      Begin VB.Menu SSmenuItem 
         Caption         =   "Load SnapShot"
         Index           =   0
      End
      Begin VB.Menu SSmenuItem 
         Caption         =   "Save SnapShot"
         Index           =   1
      End
      Begin VB.Menu SSmenuItem 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu SSmenuItem 
         Caption         =   "Close Menu"
         Index           =   3
      End
   End
End
Attribute VB_Name = "AllPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public isCB As Boolean
Public lbHwnd As Long

Dim SSini As String
Private Sub Command_Click(Index As Integer)
    Select Case Index
        Case 0
            Call LBrefresh
        Case 1
            PopupMenu SSmenu
    End Select
End Sub
Private Sub Form_Load()
    Dim P As Integer

    Call LoadStandardForm(Me, False)
    
    For P = Command.LBound To Command.UBound
        Call SetMP(Command(P), True)
    Next P
    
    Call LBrefresh
    
    SSini = mY.Path & "SnapShot.ini"
    
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Call FormMove(Me)
End Sub
Private Sub LB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton And LB.ListIndex > -1 Then PopupMenu Menu
End Sub
Private Sub LB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call LBmouseMove(LB, Button, Shift, X, Y)
End Sub
Private Sub MenuItem_Click(Index As Integer)
    Dim TaskID As Long
    Dim Path As String
    Dim P As Integer
    
    Select Case LCase(MenuItem(Index).Caption)
        Case "terminate process"
            TaskID = lbID.List(LB.ListIndex)
            
            TaskID = OpenProcess(PROCESS_TERMINATE, 0&, TaskID)
            Call TerminateProcess(TaskID, 1&)
            Call CloseHandle(TaskID)
            
            Call LBrefresh
        Case "open path in explorer"
            Path = LB.List(LB.ListIndex)
            
            For P = Len(Path) To 1 Step -1
                If mID(Path, P, 1) = "\" Then Exit For
            Next P
            
            Path = mID(Path, 1, P - 1)
            
            Call OpenEXE(Path)
    End Select

End Sub
Private Sub LBrefresh()
    Dim P As Integer
    
    Main.GetHwnd.Enabled = False
    
    Do While Main.TimerWork
        DoEvents
    Loop
    
    LB.Clear
    lbID.Clear
    
    For P = 0 To Main.LB.ListCount - 1
        LB.AddItem Main.LB.List(P)
        lbID.AddItem Main.lbID.List(P)
    Next P
    
    Main.GetHwnd.Enabled = True
    
End Sub
Private Sub SSmenuItem_Click(Index As Integer)
    Dim Path As String
    Dim P As Integer
    Dim X As Long
    
    Select Case Index
        Case 0
            P = 0
            Call LBrefresh
            
            Do
                Path = GetFromINI("SnapShot", CStr(P), SSini)
                
                If Path <> "" Then
                    If SearchLB(LB.Hwnd, _
                                 LB_FINDSTRINGEXACT, -1, Path) = -1 Then _
                        Call ShellExecute(0, "Open", Path, "", "", 1)
                Else
                    Exit Do
                End If
                
                P = P + 1
                
            Loop
            
            Call LBrefresh
        Case 1
            On Local Error Resume Next
            
            Kill SSini
            
            For P = 0 To LB.ListCount - 1
                Call Write2INI("SnapShot", CStr(P), _
                               LB.List(P), SSini)
            Next P
    End Select
End Sub
