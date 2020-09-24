VERSION 5.00
Begin VB.Form Help 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Help"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "Yep"
      Height          =   255
      Left            =   2760
      TabIndex        =   1
      Top             =   5760
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   1230
      Left            =   480
      Picture         =   "Help.frx":0000
      ScaleHeight     =   1230
      ScaleWidth      =   2895
      TabIndex        =   3
      ToolTipText     =   "just click it ;)"
      Top             =   4320
      Width           =   2895
   End
   Begin VB.Label Label 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "woeh@gmx.de"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Bräuchte Rückmeldung zwecks Verbesserung, & Optimierung"
      Top             =   3840
      Width           =   3615
   End
   Begin VB.Label Label 
      Alignment       =   2  'Zentriert
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
      ForeColor       =   &H00FFFFFF&
      Height          =   3555
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3570
   End
End
Attribute VB_Name = "Help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Fade As Boolean
Private Sub Command1_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    Dim t As String, Cap As String
    Dim h As Integer
    
    Main.Hide
    
    Me.Top = Main.Top
    Me.Left = Main.Left
    
    Call LoadStandardForm(Me, False)
    
    Call SetMP(Label(1), True)
    Call SetMP(Picture1, True)
    Call SetMP(Command1, True)
    
    t = vbCrLf
    If Not Fade Then
        h = 6495
        
        Cap = "Über Mouse'n'Enum"
        
        t = t & "Einfach die Maus über das Control bewegen " & _
                vbCrLf & " und F12 drücken" & vbCrLf & vbCrLf
            
        t = t & "Enumiert wird über" & vbCrLf & vbCrLf & _
                "FindChildByClass," & vbCrLf & _
                "WndEnumProc" & vbCrLf & _
                "und" & vbCrLf & _
                "WndEnumChildProc" & _
                vbCrLf & vbCrLf

        t = t & "Eigentlich gibt es sonst nicht viel zu sagen..." & _
                "Außer das mich das Prog 'MouseInfo' von 'Stirb' " & _
                "inspiriert hat. Irgendwann fand ich, das ich " & _
                "gerne mehr Infos und eine kompfortablere Oberfläche " & _
                "haben wollte...Da es diese jedoch nicht gab (Oder habe " & _
                "ich es bloß nicht geschafft sie zu finden ?), habe ich " & _
                "halt selber eine geschrieben."
                
        Label(1).Caption = My.Mail
    Else
        h = 7250
        
        Cap = "Hmmm....."
        
        t = t & "Windows betrachtet jeden Button, jede" & vbCrLf & _
                "Eingabebox, jede Listbox u.s.w. als " & _
                "eigenes seperates Fenster." & _
                vbCrLf & vbCrLf
                
        t = t & "Jedes dieser Controls (Fenster) ist einem" & vbCrLf & _
                "Übergeordneten Control zugeordnet. Das Control," & vbCrLf & _
                "welches als oberstes zugeordnet ist," & vbCrLf & _
                "nennt man Parent.Alle untergeordneten Controls" & vbCrLf & _
                "nennt man Childs" & vbCrLf & _
                "(Ein Child kann aber auch wiederrum ein Parent " & vbCrLf & _
                "für ein anderes Child sein)." & _
                vbCrLf & vbCrLf
                
        t = t & "Jedes dieser Controls hat nun eine" & vbCrLf & _
                "eindeutige Nummer von Windows bei seiner" & vbCrLf & _
                "Entstehung zugewiesen bekommen (das hWnd" & vbCrLf & _
                "oder auch WindowHandle).Über diese Nummer" & vbCrLf & _
                "ist das Control ansprechbar, abfragbar ," & vbCrLf & _
                "enumierbar etc." & _
                vbCrLf & vbCrLf
    
        t = t & "Mouse'n'Enum findet die Controls der" & vbCrLf & _
                "gewünschten Fenster und kann bestimmte" & vbCrLf & _
                "Eigenschaften der Controls strukturiert" & vbCrLf & _
                "darstellen." & _
                vbCrLf & vbCrLf
        
        
        t = t & "Wenn du wissen möchtest wie man ein Control" & vbCrLf & _
                "findet, abfragt u.s.w, dann klicke bitte auf" & vbCrLf & _
                "Link unterhalb des Textes."
        
        Me.Height = h
        
        Label(0).Height = 5500
                    
        Command1.Top = Me.Height - 700
        Command1.Caption = "Hmmm..."
        
        Label(1).Caption = "ActiveVB Tutorial"
        Label(1).ToolTipText = ""
        Label(1).Top = Command1.Top - 700
        
        Picture1.Visible = False
    End If

    Me.Caption = Cap
    
    Label(0).Caption = t
    
    Me.Height = 0
    Me.Width = 0
    
    Me.Visible = True
        
    Call RollForm(Me, 0, h, 0, 3945, True, True, 0, 255)

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Call FormMove(Me)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call RollForm(Me, Me.Height, 0, Me.Width, 0, False, True, 0, 255)
        
    Main.Left = Me.Left
    Main.Top = Me.Top
    Main.Show
End Sub
Public Sub Label_Click(Index As Integer)
    If Label(1).Caption = My.Mail Then
        Call SendMail(My.Mail)
    Else
        Call PageVisit(My.VBtutor)
    End If
End Sub
Private Sub Label_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Index = 0 Or Index = 2) And Button = vbLeftButton Then _
                                        Call FormMove(Me)
End Sub
Public Sub Picture1_Click()
    Dim t As String
            
    Me.Hide
    
    t = "Ich möchte an dieser Stelle ActiveVB und dem ActiveVB-Forum " & _
        "danken..." & vbCrLf & vbCrLf & "Ich wäre heute nicht da wo " & _
        "ich bin ohne ActiveVB und kann diese Seite nur jedem " & _
        "VB-Coder empfehlen." & _
        vbCrLf & vbCrLf & _
        "Möchtest du die ActiveVB-Seite besuchen ?" & _
        vbCrLf
        
    If Ask(t, "Kurz noch...", _
           vbYesNoCancel + vbInformation) = vbYes Then _
                                            Call PageVisit(My.ActiveVB)
                            
    Me.Show
    
End Sub
