VERSION 5.00
Begin VB.Form CbLB 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Form1"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command 
      Caption         =   "Close"
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   0
      Top             =   1200
      Width           =   735
   End
   Begin VB.ComboBox CB 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown-Liste
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.ListBox LB 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   645
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label Label 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Box"
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
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   5
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Box"
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
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label MSG 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Reading Box....."
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu MenuItem 
         Caption         =   "Add new Item here"
         Index           =   0
      End
      Begin VB.Menu MenuItem 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu MenuItem 
         Caption         =   "Delete Item"
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
   Begin VB.Menu BoxMenu 
      Caption         =   "BoxMenu"
      Visible         =   0   'False
      Begin VB.Menu BoxMenuItem 
         Caption         =   "Add Item"
         Index           =   0
      End
      Begin VB.Menu BoxMenuItem 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu BoxMenuItem 
         Caption         =   "Disable"
         Index           =   2
      End
      Begin VB.Menu BoxMenuItem 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu BoxMenuItem 
         Caption         =   "Clear Box"
         Index           =   4
      End
      Begin VB.Menu BoxMenuItem 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu BoxMenuItem 
         Caption         =   "Refresh"
         Index           =   6
      End
      Begin VB.Menu BoxMenuItem 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu BoxMenuItem 
         Caption         =   "Close Menu"
         Index           =   8
      End
   End
End
Attribute VB_Name = "CbLB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MyHwnd As Long

Dim isLB As Boolean, isCB As Boolean, Read As Boolean
Private Sub BoxMenuItem_Click(Index As Integer)
    Dim Item As String, t As String
    Dim Answer As VbMsgBoxResult
    
    Select Case LCase(BoxMenuItem(Index).Caption)
        Case "disable"
            Call EnableWindow(MyHwnd, 0)
            BoxMenuItem(Index).Caption = "Enable"
        Case "enable"
            Call EnableWindow(MyHwnd, 1)
            BoxMenuItem(Index).Caption = "Disable"
        Case "clear box"
            If isLB Then
                Call ClearLB(MyHwnd, False)
            Else
                Call ClearLB(MyHwnd, True)
            End If
            Call GetHwnd(MyHwnd)
        Case "add item"
            Item = uInput("New Item to add ?", "Add Item to Box")
            If LB.ListCount = 0 And CB.ListCount = 0 Then
                t = "Not Items were found." & vbCrLf & vbCrLf & _
                    "Do you wish to send the String to a ListBox ?" & _
                    vbCrLf & "(Press NO to send to a ComboBox)"
                Answer = Ask(t, "Add new Item", _
                             vbQuestion + vbYesNoCancel)
                If Answer = vbCancel Then Exit Sub
            End If
            
            If isLB Or Answer = vbYes Then
                Call AddLBitem(MyHwnd, 0, Answer, False)
            Else
                Call AddLBitem(MyHwnd, 0, Answer, True)
            End If
                
            Call GetHwnd(MyHwnd)
        Case "refresh"
            Call GetHwnd(MyHwnd)
    End Select
End Sub
Private Sub CB_Click()
    If Me.Visible And Not Read Then _
                       Call CBdropdown(MyHwnd): _
                       Call SetLBlistIndex(MyHwnd, _
                                           CB.ListIndex, True): _
                       Call SendMessage(MyHwnd, _
                                        WM_LBUTTONDOWN, 0, 0&)
End Sub
Private Sub Label_Click(Index As Integer)
    Select Case Index
        Case 0
            If LB.ListCount = 0 And CB.ListCount = 0 Then
                BoxMenuItem(4).Enabled = False
            Else
                BoxMenuItem(4).Enabled = True
            End If

            PopupMenu BoxMenu
    End Select
End Sub
Private Sub LB_DblClick()
   Call SetLBlistIndex(MyHwnd, LB.ListIndex, False)
End Sub
Private Sub LB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton And LB.ListCount > 0 Then PopupMenu Menu
End Sub
Private Sub LB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call LBmouseMove(LB, Button, Shift, X, Y)
End Sub
Private Sub Command_Click(Index As Integer)
    Select Case Index
        Case 0
            Unload Me
    End Select
End Sub
Private Sub Form_Activate()
    Call FadeForm(Me, 0, 255)
End Sub
Private Sub Form_Load()
    Call LoadStandardForm(Me, True)
    
    Me.Width = 3105
    
    Call SetMP(Label(0), True)
    Call SetMP(Command(0), True)
    Call SetMP(CB, True)
    
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Call FormMove(Me)
End Sub
Private Sub MenuItem_Click(Index As Integer)
    Dim Answer As String, Item As String
    Dim Li As Integer
    
    If Index = MenuItem.UBound Then Exit Sub
    
    Li = LB.ListIndex
    Item = LB.List(Li)
    
    Select Case LCase(MenuItem(Index).Caption)
        Case "add new item here"
            Answer = uInput("New Item to add ?", "Add Item to Box")
            
            If isLB Then
                Call AddLBitem(MyHwnd, Li, Answer, False)
            Else
                Call AddLBitem(Hwnd, Li, Answer, True)
            End If
        Case "delete item"
            If isLB Then
                Call DeleteLBitem(MyHwnd, Li, False)
            Else
                Call DeleteLBitem(MyHwnd, Li, True)
            End If
    End Select
    
    Call GetHwnd(MyHwnd)

End Sub
Public Sub GetHwnd(ByVal Hwnd As Long)
        
    Read = True
        
    MyHwnd = Hwnd
    
    MSG.Caption = "Reading Box....."
    Me.Refresh
    
    CB.Visible = False
    LB.Visible = False
    
    isLB = False
    isCB = False
    
    LB.Clear
    CB.Clear
    
    If ReadLB(MyHwnd, LB) > 0 Then
        isLB = True
        LB.Visible = True
        MSG.Caption = "Found " & LB.ListCount & " Items from " & MyHwnd
        Me.Height = 1920
        Label(0).Top = 1200
        Command(0).Top = 1200
    Else
        If ReadCB(MyHwnd, CB) > 0 Then
            isCB = True
            CB.Visible = True
            MSG.Caption = "Found " & CB.ListCount & " Items from " & MyHwnd
            Me.Height = 1545
            Label(0).Top = 840
            Command(0).Top = 840
        Else
            MSG.Caption = MyHwnd & " hasn't any Item or isn't a Box"
            Me.Height = 1200
            Label(0).Top = 480
            Command(0).Top = 480
        End If
    End If
    
    Call FadeForm(Me, 0, 255)
    
    Me.Visible = True
    
    Read = False
    
End Sub
Private Sub MSG_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Call FormMove(Me)
End Sub
