VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form MnuInfo 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Form1"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ListBox LB 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin RichTextLib.RichTextBox Rtext 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   2778
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      MousePointer    =   3
      Appearance      =   0
      TextRTF         =   $"MnuInfo.frx":0000
   End
   Begin VB.Label Label 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
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
      Left            =   2880
      TabIndex        =   3
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Enumieren"
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
      Left            =   1560
      TabIndex        =   2
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu MenuItem 
         Caption         =   "Menu enumieren"
         Index           =   0
      End
      Begin VB.Menu MenuItem 
         Caption         =   "SystemMenu enumieren"
         Index           =   1
      End
      Begin VB.Menu MenuItem 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu MenuItem 
         Caption         =   "Menus von allen Childs enumieren"
         Index           =   3
      End
      Begin VB.Menu MenuItem 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu MenuItem 
         Caption         =   "Close Menu"
         Index           =   5
      End
   End
End
Attribute VB_Name = "MnuInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyHwnd As Long
Dim mItemCount As Integer, Pmenu As Integer, mSub As Integer
Private Sub Form_Load()
    Call LoadStandardForm(Me, False)
    Call SetMP(Label(0), True)
    Call SetMP(Label(1), True)
End Sub
Private Sub SMenu(mHwnd As Long, Optional iSubFlag As Boolean)
    
    Dim n As Long, i As Long
            
    On Local Error Resume Next
    
    n = GetMenuItemCount(mHwnd)
    
    For i = 0 To n - 1
        If Not iSubFlag Then
            Pmenu = Pmenu + 1
            If GetSubMenu(mHwnd, i) > 1 Then
                iSubFlag = True
                SMenu GetSubMenu(mHwnd, i), True
                iSubFlag = False
            End If
        Else
            mItemCount = mItemCount + 1
            If GetSubMenu(mHwnd, i) > 1 Then
                mSub = mSub + 1
                iSubFlag = True
                SMenu GetSubMenu(mHwnd, i), True
            End If
        End If
   Next i
    
End Sub
Public Sub GetHwnd(Hwnd As Long)
    Dim h As Long, Color As Long, miCount
    Dim P As Integer
    Dim R$

    MyHwnd = Hwnd
    
    LB.Clear
    
    LB.AddItem vbYellow
    LB.AddItem "MenuInfo from " & MyHwnd
    LB.AddItem "------------------------------------"
    
    For P = 1 To 2
        If P = 1 Then
            h = GetMenu(MyHwnd)
            R$ = "GetMenu(" & MyHwnd & ") = " & h
        Else
            h = GetSystemMenu(MyHwnd, False)
            R$ = "GetSystemMenu(" & MyHwnd & ", False) = " & h
        End If
        
        Color = vbGreen
        If h = 0 Then Color = vbRed
        
        LB.AddItem Color
        LB.AddItem R$
        
        miCount = GetMenuItemCount(h)
        R$ = "GetMenuItemCount(" & h & ") = " & miCount
        
        If miCount < 0 Then
            LB.AddItem vbRed
            LB.AddItem R$
            
            If P = 1 Then
                LB.AddItem "   => Control has no Menu"
                MenuItem(0).Enabled = False
            Else
                LB.AddItem "   => Control has no SystemMenu"
                MenuItem(1).Enabled = False
            End If
        Else
            LB.AddItem vbGreen
            LB.AddItem R$
        
            mItemCount = 0
            Pmenu = 0
            mSub = 0
        
            If P = 1 Then
                Call SMenu(GetMenu(MyHwnd), False)
            Else
                Call SMenu(GetSystemMenu(MyHwnd, False), False)
            End If
        
            LB.AddItem vbWhite
            
            If P = 1 Then
                LB.AddItem "   => Menus = " & Pmenu
                LB.AddItem "   => Menuentries = " & mItemCount
            Else
                LB.AddItem "   => Menuentries = " & Pmenu
            End If
            
            LB.AddItem "   => Submenus = " & mSub
        End If
    
        LB.AddItem " "
    
    Next P

    Call WriteLB(LB, Rtext)
            
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Call FormMove(Me)
End Sub
Private Sub Label_Click(Index As Integer)
    If Index = 1 Then Unload Me
End Sub
Private Sub Label_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 And Button = vbLeftButton Then PopupMenu Menu
End Sub
Private Sub MenuItem_Click(Index As Integer)
    Dim A As New Ac
    
    If Index < 4 Then Load A
    
    Select Case Index
        Case 0
            Call A.MenuEnum(MyHwnd, False, False)
        Case 1
            Call A.MenuEnum(MyHwnd, True, False)
        Case 3
            Call A.MenuEnum(MyHwnd, False, True)
    End Select
    
    If Index < 4 Then A.Visible = True
    
End Sub
