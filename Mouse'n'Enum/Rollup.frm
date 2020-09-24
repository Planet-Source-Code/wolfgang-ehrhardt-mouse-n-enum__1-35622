VERSION 5.00
Begin VB.Form Cap 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'Kein
   ClientHeight    =   3075
   ClientLeft      =   1875
   ClientTop       =   1860
   ClientWidth     =   3600
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   3075
   ScaleWidth      =   3600
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Blink 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2640
      Top             =   0
   End
   Begin VB.CommandButton Command 
      Caption         =   "Hide"
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   19
      Top             =   2760
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   3120
      Top             =   0
   End
   Begin VB.CommandButton Command 
      Caption         =   "Close"
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   0
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label LCommand 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Index           =   1
      Left            =   885
      TabIndex        =   24
      Top             =   2760
      Width           =   525
   End
   Begin VB.Label LCommand 
      Alignment       =   2  'Zentriert
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
      ForeColor       =   &H0000FF00&
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   23
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "TaskID:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   22
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   13
      Left            =   1200
      TabIndex        =   21
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   20
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   3615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label12"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   16
      Left            =   1200
      TabIndex        =   18
      Top             =   1800
      Width           =   570
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Visible:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   17
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Größe:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   16
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label12"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   15
      Left            =   1200
      TabIndex        =   15
      Top             =   1560
      Width           =   570
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Position:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   14
      Left            =   1200
      TabIndex        =   13
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "---"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   19
      Left            =   1200
      TabIndex        =   12
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ParentText:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "---"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   18
      Left            =   1200
      TabIndex        =   10
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label9"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   17
      Left            =   1200
      TabIndex        =   9
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   10
      Left            =   1200
      TabIndex        =   8
      Top             =   360
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   12
      Left            =   1200
      TabIndex        =   7
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   11
      Left            =   1200
      TabIndex        =   6
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ParentClass:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ParentHwnd:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Handle:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Class:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Text:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu MenuItem 
         Caption         =   "Hwnd"
         Index           =   0
      End
      Begin VB.Menu MenuItem 
         Caption         =   "Childs"
         Index           =   1
      End
      Begin VB.Menu MenuItem 
         Caption         =   "Parent"
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
   Begin VB.Menu View 
      Caption         =   "View"
      Visible         =   0   'False
      Begin VB.Menu ViewItem 
         Caption         =   ""
         Index           =   0
      End
      Begin VB.Menu ViewItem 
         Caption         =   ""
         Index           =   1
      End
      Begin VB.Menu ViewItem 
         Caption         =   ""
         Index           =   2
      End
      Begin VB.Menu ViewItem 
         Caption         =   ""
         Index           =   3
      End
      Begin VB.Menu ViewItem 
         Caption         =   ""
         Index           =   4
      End
      Begin VB.Menu ViewItem 
         Caption         =   ""
         Index           =   5
      End
      Begin VB.Menu ViewItem 
         Caption         =   ""
         Index           =   6
      End
      Begin VB.Menu ViewItem 
         Caption         =   ""
         Index           =   7
      End
      Begin VB.Menu ViewItem 
         Caption         =   ""
         Index           =   8
      End
      Begin VB.Menu ViewItem 
         Caption         =   ""
         Index           =   9
      End
      Begin VB.Menu ViewItem 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu ViewItem 
         Caption         =   "Close Menu"
         Index           =   11
      End
   End
End
Attribute VB_Name = "Cap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MyHwnd As Long
Dim Parent As Long
Dim CapSave(9) As Boolean
Private Sub Blink_Timer()
    Static Count As Integer
    
    Count = Count + 1
    
    If Count = 10 Then
        Label1(20).ForeColor = &HFF&
        Count = 0
        Blink.Enabled = False
    Else
        If Label1(20).ForeColor = &HFF& Then
            Label1(20).ForeColor = &HFFFFFF
        Else
            Label1(20).ForeColor = &HFF&
        End If
    End If
    
End Sub
Private Sub Command_Click(Index As Integer)
    Select Case Index
        Case 0
            Me.Hide
        Case 1
            Unload Me
    End Select
End Sub
Private Sub Form_Activate()
    Me.Refresh
End Sub
Private Sub Form_Load()
    Dim P As Integer
    
    On Local Error Resume Next
    
    Call LoadStandardForm(Me, False)
    
    For P = 0 To 1
        Call SetMP(LCommand(P), True)
        Call SetMP(Command(P), True)
    Next P
    
    Me.Height = 3165
    Me.Width = 3630
    
    For P = 0 To 9
        ViewItem(P).Caption = Left(Label1(P).Caption, Len(Label1(P)) - 1)
        CapSave(P) = CapView(P)
    Next P
        
    Parent = -1
    
    Me.Caption = ""
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call Main.FindCBitem(str(MyHwnd), True)
End Sub
Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Cap As String
    
    If Index < 20 Then
        If Index > 9 Then Index = Index - 10
        
        Cap = Label1(Index).Caption & " " & Label1(Index + 10).Caption
        Label1(Index).ToolTipText = Cap
        Label1(Index + 10).ToolTipText = Cap
    End If
    
End Sub
Public Sub GetHwnd(Hwnd As Long)
    Dim Class As String, V As String, t As String
    Dim TaskID As Long
    Dim R As RECT
    Dim P As Integer
    
    Static oParent As Long
    
    On Local Error Resume Next
    
    If GetClass(MyHwnd) = "" Then
        Command(0).Visible = False
        For P = 0 To 19
            Label1(P).Visible = False
        Next P
        For P = 0 To LCommand.UBound
            LCommand(P).Visible = False
        Next P
        Command(1).Top = 720
        Me.Height = 1035
        Label1(20).Caption = "Control " & Hwnd & vbCrLf & _
                                                     "is unloaded"
        Label1(20).Height = 615
        Label1(20).ToolTipText = Label1(20).Caption
        Timer1.Enabled = False
        
        Exit Sub
    
    End If
    
    V = "Yes"
    If IsWindowVisible(MyHwnd) = 0 Then V = "No"
    
    Call GetWindowRect(Hwnd, R)
    
    t = GetText(Hwnd)
    If t <> "" Then
        t = Chr(34) & t & Chr(34)
    Else
        t = "Control has no Text"
    End If
    
    Label1(10).Caption = t
    Label1(14).Caption = "x = " & R.Left & ", y = " & R.Top
    Label1(15).Caption = "Width: " & R.Right - R.Left & ", " & _
                         "Height: " & R.Bottom - R.Top
    Label1(16).Caption = V
    
    If oParent <> Parent Then
        Parent = GetParent(Hwnd)
        oParent = Parent
        If Parent Then
            Label1(17).ForeColor = &HFFFF&
            Label1(17).Caption = Parent
            Label1(18).Caption = GetClass(Parent)
            Label1(19).Caption = GetText(Parent)
        Else
            Label1(17).ForeColor = &HFF00&
            Label1(17).Caption = "Control is Parent"
            Label1(18).Caption = "---"
            Label1(19).Caption = "---"
        End If
    End If
    
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Call FormMove(Me)
End Sub
Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Call FormMove(Me)
End Sub
Private Sub LCommand_Click(Index As Integer)
    Select Case Index
        Case 0
            PopupMenu Menu
        Case 1
            PopupMenu View
    End Select
End Sub
Private Sub MenuItem_Click(Index As Integer)
    Select Case Index
        Case 0
            Call LEnum(MyHwnd, True, False)
        Case 1
            Call LEnum(MyHwnd, False, True)
        Case 2
            Call LEnum(MyHwnd, False, False)
    End Select
End Sub
Private Sub Timer1_Timer()
    If Me.Visible Then Call GetHwnd(MyHwnd)
End Sub
Public Sub ViewItem_Click(Index As Integer)
    Dim P As Integer, t As Integer
    
    If Index <= ViewItem.UBound - 2 Then
        If ViewItem(Index).Checked Then
            t = 0
            
            For P = 0 To ViewItem.UBound - 2
                If ViewItem(P).Checked Then t = t + 1
            Next P
            
            If t < 2 Then Exit Sub
            
            ViewItem(Index).Checked = False
        Else
            ViewItem(Index).Checked = True
        End If
        
        t = 360
        
        For P = 0 To ViewItem.UBound - 2
            If Not ViewItem(P).Checked Then
                Label1(P).Visible = False
                Label1(P + 10).Visible = False
                CapView(P) = False
            Else
                Label1(P).Top = t
                Label1(P + 10).Top = t
                Label1(P).Visible = True
                Label1(P + 10).Visible = True
                
                t = t + Label1(P).Height - 15
                CapView(P) = True
            End If
        Next P
        
        t = t + 16
        
        For P = 0 To LCommand.UBound
            LCommand(P).Top = t
        Next P
        
        For P = 0 To Command.UBound
            Command(P).Top = t
        Next P
        
        Me.Height = t + Command(0).Height + 100
        
    End If
        
End Sub
Public Sub GetStaticInfo(ByVal Hwnd As Long)
    Dim P As Integer
    Dim TaskID As Long
    
    MyHwnd = Hwnd
    Me.Tag = "Cap" & MyHwnd
    
    Call GetWindowThreadProcessId(MyHwnd, TaskID)

    Label1(11).Caption = GetClass(MyHwnd)
    Label1(12).Caption = MyHwnd
    Label1(13).Caption = TaskID

    Call GetHwnd(MyHwnd)

    Label1(20).Caption = "Capturing " & MyHwnd
    
    MenuItem(0).Caption = "Control " & MyHwnd
    MenuItem(1).Caption = "All Childs from " & MyHwnd
    MenuItem(2).Caption = "Parent from " & MyHwnd
                    
    For P = 0 To ViewItem.UBound - 2
        ViewItem(P).Checked = True
    Next P
    
    For P = 0 To ViewItem.UBound - 2
        If Not CapSave(P) Then Call ViewItem_Click(P)
    Next P
        
    For P = 0 To ViewItem.UBound - 2
        CapView(P) = CapSave(P)
    Next P
    
    Label1(20).ToolTipText = Label1(20).Caption

    Timer1.Enabled = True

    Me.Show
    
End Sub
