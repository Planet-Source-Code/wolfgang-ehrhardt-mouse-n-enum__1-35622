VERSION 5.00
Begin VB.Form Main 
   AutoRedraw      =   -1  'True
   ClientHeight    =   6270
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   7170
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ListBox Auto 
      Height          =   255
      Left            =   120
      TabIndex        =   69
      Top             =   5400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox Check 
      Height          =   195
      Index           =   10
      Left            =   4800
      MaskColor       =   &H8000000F&
      TabIndex        =   65
      Top             =   1080
      UseMaskColor    =   -1  'True
      Value           =   1  'Aktiviert
      Width           =   175
   End
   Begin VB.CheckBox Check 
      Height          =   195
      Index           =   9
      Left            =   4800
      MaskColor       =   &H8000000F&
      TabIndex        =   63
      Top             =   1320
      UseMaskColor    =   -1  'True
      Value           =   1  'Aktiviert
      Width           =   175
   End
   Begin VB.CheckBox Check 
      Height          =   195
      Index           =   5
      Left            =   4800
      MaskColor       =   &H8000000F&
      TabIndex        =   58
      Top             =   1560
      UseMaskColor    =   -1  'True
      Value           =   1  'Aktiviert
      Width           =   175
   End
   Begin VB.CheckBox Check 
      Height          =   195
      Index           =   8
      Left            =   4800
      MaskColor       =   &H8000000F&
      TabIndex        =   57
      Top             =   2280
      UseMaskColor    =   -1  'True
      Value           =   1  'Aktiviert
      Width           =   175
   End
   Begin VB.CheckBox Check 
      Height          =   195
      Index           =   7
      Left            =   4800
      MaskColor       =   &H8000000F&
      TabIndex        =   56
      Top             =   2040
      UseMaskColor    =   -1  'True
      Value           =   1  'Aktiviert
      Width           =   175
   End
   Begin VB.CheckBox Check 
      Height          =   195
      Index           =   6
      Left            =   4800
      MaskColor       =   &H8000000F&
      TabIndex        =   55
      Top             =   1800
      UseMaskColor    =   -1  'True
      Value           =   1  'Aktiviert
      Width           =   175
   End
   Begin VB.CheckBox Check 
      Height          =   195
      Index           =   4
      Left            =   4800
      MaskColor       =   &H8000000F&
      TabIndex        =   54
      Top             =   3120
      UseMaskColor    =   -1  'True
      Value           =   1  'Aktiviert
      Width           =   175
   End
   Begin VB.CheckBox Check 
      Height          =   195
      Index           =   3
      Left            =   4800
      MaskColor       =   &H8000000F&
      TabIndex        =   53
      Top             =   2880
      UseMaskColor    =   -1  'True
      Value           =   1  'Aktiviert
      Width           =   175
   End
   Begin VB.CheckBox Check 
      Height          =   195
      Index           =   2
      Left            =   4800
      MaskColor       =   &H8000000F&
      TabIndex        =   52
      Top             =   600
      UseMaskColor    =   -1  'True
      Value           =   1  'Aktiviert
      Width           =   175
   End
   Begin VB.CheckBox Check 
      Height          =   195
      Index           =   1
      Left            =   4800
      MaskColor       =   &H8000000F&
      TabIndex        =   51
      Top             =   360
      UseMaskColor    =   -1  'True
      Value           =   1  'Aktiviert
      Width           =   175
   End
   Begin VB.CheckBox Check 
      Height          =   195
      Index           =   0
      Left            =   4800
      MaskColor       =   &H8000000F&
      TabIndex        =   50
      Top             =   840
      UseMaskColor    =   -1  'True
      Value           =   1  'Aktiviert
      Width           =   175
   End
   Begin VB.ListBox lbID 
      Height          =   255
      Left            =   2040
      TabIndex        =   45
      Top             =   5160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox LB 
      Height          =   255
      Left            =   2040
      TabIndex        =   44
      Top             =   4920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton ShowMore 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   32
      Top             =   120
      Width           =   255
   End
   Begin VB.ListBox AllC 
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   5160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox Temp 
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   4920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox CapC 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown-Liste
      TabIndex        =   19
      ToolTipText     =   "Captured Windows"
      Top             =   3480
      Width           =   1935
   End
   Begin VB.PictureBox pichook 
      Height          =   435
      Left            =   1560
      ScaleHeight     =   375
      ScaleWidth      =   315
      TabIndex        =   12
      Top             =   4920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer GetHwnd 
      Interval        =   10
      Left            =   1080
      Top             =   4920
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Hwnd"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   10
      Left            =   1200
      TabIndex        =   68
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "TopParent:"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   67
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label LMore 
      BackStyle       =   0  'Transparent
      Caption         =   "Parent"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   3600
      TabIndex        =   66
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label LMore 
      BackStyle       =   0  'Transparent
      Caption         =   "Child"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   3600
      TabIndex        =   64
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label MenuLabel 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
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
      Height          =   195
      Index           =   3
      Left            =   1680
      TabIndex        =   62
      Top             =   120
      Width           =   645
   End
   Begin VB.Label MP 
      Caption         =   "HandMP"
      Height          =   255
      Index           =   1
      Left            =   3240
      MouseIcon       =   "Form1.frx":030A
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   61
      Top             =   5160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label MP 
      Caption         =   "MoveMP"
      Height          =   255
      Index           =   0
      Left            =   3240
      MouseIcon       =   "Form1.frx":045C
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   60
      Top             =   4920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label MouseK 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "MouseK"
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   2160
      TabIndex        =   59
      ToolTipText     =   "Mousekoordinaten"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label LMore 
      BackStyle       =   0  'Transparent
      Caption         =   "Yes"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   4680
      TabIndex        =   49
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label LMore 
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   3600
      TabIndex        =   48
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label EXEname 
      BackStyle       =   0  'Transparent
      Caption         =   "Modul:"
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   47
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label EXEname 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      ForeColor       =   &H00FFC0C0&
      Height          =   495
      Index           =   1
      Left            =   1200
      TabIndex        =   46
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label LMore 
      BackStyle       =   0  'Transparent
      Caption         =   "CloseButton"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   3600
      TabIndex        =   43
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label LMore 
      BackStyle       =   0  'Transparent
      Caption         =   "MaxButton"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   3600
      TabIndex        =   42
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label LMore 
      BackStyle       =   0  'Transparent
      Caption         =   "MinButton"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   3600
      TabIndex        =   41
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label LMore 
      BackStyle       =   0  'Transparent
      Caption         =   "WindowState"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   3600
      TabIndex        =   40
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label LMore 
      BackStyle       =   0  'Transparent
      Caption         =   "ToolWindow"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   3600
      TabIndex        =   39
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label LMore 
      BackStyle       =   0  'Transparent
      Caption         =   "AppWindow"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   38
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label LMore 
      BackStyle       =   0  'Transparent
      Caption         =   "Enabled"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   3600
      TabIndex        =   37
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label LMore 
      BackStyle       =   0  'Transparent
      Caption         =   "Topmost"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   3600
      TabIndex        =   36
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label LMore 
      BackStyle       =   0  'Transparent
      Caption         =   "MDI"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   3600
      TabIndex        =   35
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Thread:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   34
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   1200
      TabIndex        =   33
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Line Linie 
      BorderColor     =   &H00FFFFFF&
      X1              =   3480
      X2              =   3480
      Y1              =   360
      Y2              =   3480
   End
   Begin VB.Label MenuLabel 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tools"
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
      Height          =   195
      Index           =   1
      Left            =   1050
      TabIndex        =   31
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ParentText:"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   9
      Left            =   105
      TabIndex        =   10
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label12"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   12
      Left            =   1200
      TabIndex        =   29
      Top             =   1320
      Width           =   2130
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "TaskID:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   28
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label CurrentColor 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   480
      TabIndex        =   27
      Top             =   3960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "Decimal"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   1320
      TabIndex        =   26
      Top             =   3960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "Hex"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   25
      Top             =   4200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "HTML"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   1320
      TabIndex        =   24
      Top             =   4440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "Dec"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   2040
      TabIndex        =   23
      Top             =   3960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "Hex"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   2040
      TabIndex        =   22
      Top             =   4200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "HTML"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   2040
      TabIndex        =   21
      Top             =   4440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Größe:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   18
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label12"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   16
      Left            =   1200
      TabIndex        =   17
      Top             =   2280
      Width           =   2130
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Position:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   15
      Left            =   1200
      TabIndex        =   15
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label MenuLabel 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Optionen"
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
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   795
   End
   Begin VB.Label MenuLabel 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hilfe"
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
      Height          =   195
      Index           =   2
      Left            =   2520
      TabIndex        =   13
      Top             =   120
      Width           =   435
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label12"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   19
      Left            =   1200
      TabIndex        =   11
      Top             =   3000
      Width           =   2130
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label10"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   18
      Left            =   1200
      TabIndex        =   9
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label9"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   17
      Left            =   1200
      TabIndex        =   8
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   14
      Left            =   1200
      TabIndex        =   7
      Top             =   1800
      Width           =   2160
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   10
      Left            =   1200
      TabIndex        =   6
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   13
      Left            =   1200
      TabIndex        =   5
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ParentClass:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ParentHwnd:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Handle:"
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Class:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Text:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Menu ToolMenu 
      Caption         =   "Tools"
      Visible         =   0   'False
      Begin VB.Menu ToolMenuItem 
         Caption         =   "Tasks"
         Index           =   0
         Begin VB.Menu TaskMenuItem 
            Caption         =   "Alle Tasks enumieren"
            Index           =   0
         End
         Begin VB.Menu TaskMenuItem 
            Caption         =   "Alle Tasks mit Childs enumieren"
            Index           =   1
         End
         Begin VB.Menu TaskMenuItem 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu TaskMenuItem 
            Caption         =   "Manuelle Eingabe"
            Index           =   3
         End
         Begin VB.Menu TaskMenuItem 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu TaskMenuItem 
            Caption         =   "Prozesse listen"
            Index           =   5
         End
      End
      Begin VB.Menu ToolMenuItem 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu ToolMenuItem 
         Caption         =   "AutoRefresh"
         Index           =   2
         Begin VB.Menu RefreshMenuItem 
            Caption         =   "Childs immer anzeigen"
            Index           =   0
         End
         Begin VB.Menu RefreshMenuItem 
            Caption         =   "Menu immer anzeigen"
            Index           =   1
         End
      End
      Begin VB.Menu ToolMenuItem 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu ToolMenuItem 
         Caption         =   "Alle Menus im System listen"
         Index           =   4
      End
      Begin VB.Menu ToolMenuItem 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu ToolMenuItem 
         Caption         =   "Maustooltip"
         Index           =   6
      End
      Begin VB.Menu ToolMenuItem 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu ToolMenuItem 
         Caption         =   "ColorPicker"
         Index           =   8
      End
      Begin VB.Menu ToolMenuItem 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu ToolMenuItem 
         Caption         =   "Neues Fenster"
         Index           =   10
      End
      Begin VB.Menu ToolMenuItem 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu ToolMenuItem 
         Caption         =   "Close Menu"
         Index           =   12
      End
   End
   Begin VB.Menu Menu 
      Caption         =   "Optionen"
      Visible         =   0   'False
      Begin VB.Menu MenuItem 
         Caption         =   "Ansicht"
         Index           =   0
         Begin VB.Menu PosBorderItem 
            Caption         =   "Positionierungsrahmen"
            Index           =   0
         End
         Begin VB.Menu PosBorderItem 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu ViewItem 
            Caption         =   "Hwnd"
            Enabled         =   0   'False
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
            Caption         =   ""
            Index           =   10
         End
      End
      Begin VB.Menu MenuItem 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu MenuItem 
         Caption         =   "Minimiert starten"
         Index           =   2
      End
      Begin VB.Menu MenuItem 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu MenuItem 
         Caption         =   "Position speichern"
         Index           =   4
      End
      Begin VB.Menu MenuItem 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu MenuItem 
         Caption         =   "Save Menus @ Unload"
         Index           =   6
      End
      Begin VB.Menu MenuItem 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu MenuItem 
         Caption         =   "Exit"
         Index           =   8
      End
   End
   Begin VB.Menu HelpMenu 
      Caption         =   "Hilfe"
      Visible         =   0   'False
      Begin VB.Menu HelpItem 
         Caption         =   "ActiveVB"
         Index           =   0
      End
      Begin VB.Menu HelpItem 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu HelpItem 
         Caption         =   "E-Mail an den Autor"
         Index           =   2
      End
      Begin VB.Menu HelpItem 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu HelpItem 
         Caption         =   "Was sagt mir das Prog ?"
         Index           =   4
      End
      Begin VB.Menu HelpItem 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu HelpItem 
         Caption         =   "History"
         Index           =   6
      End
      Begin VB.Menu HelpItem 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu HelpItem 
         Caption         =   "Über"
         Index           =   8
      End
   End
   Begin VB.Menu SearchMenu 
      Caption         =   "SearchMenu"
      Visible         =   0   'False
      Begin VB.Menu SearchItem 
         Caption         =   "Search Windows by Class"
         Index           =   0
      End
      Begin VB.Menu SearchItem 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu SearchItem 
         Caption         =   "Search Windows by Text"
         Index           =   2
      End
      Begin VB.Menu SearchItem 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu SearchItem 
         Caption         =   "Close Menu"
         Index           =   4
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************
' * Geschrieben von                   *
' *     Wolfgang Ehrhardt             *
' *         woeh@gmx.de               *
' *                                   *
' * Der Code ist frei für Jedermann,  *
' * solange der Quellcode NICHT für   *
' * kommerzielle Zwecke verwendet     *
' * wird.                             *
' *                                   *
' * Sollten Teile oder Auszüge aus    *
' * diesem Quellcode für kommerzielle *
' * Zwecke verwendet werden, bitte    *
' * ich um Kontaktaufnahme unter      *
' *         woeh@gmx.de               *
' *                                   *
' * Vielen Dank an ActiveVB           *
' * & das ActiveVB-Forum              *
' *************************************

Private Declare Function KillTimer Lib "USER32" (ByVal Hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function SetTimer Lib "USER32" (ByVal Hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Declare Function GetPixel Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ScreenToClient Lib "USER32" (ByVal Hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function SetWindowRgn Lib "USER32" (ByVal Hwnd As Long, ByVal hRgn As Long, ByVal dRedraw As Boolean) As Long
Private Declare Function CombineRgn Lib "GDI32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateRectRgn Lib "GDI32" (ByVal XLeft As Long, ByVal YTop As Long, ByVal XRight As Long, ByVal YBottom As Long) As Long
Private Declare Function SetRect Lib "USER32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Type NOTIFYICONDATA
    cbSize As Long
    Hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Type xForm
    CPickerHeight As Integer
    MoreWide As Integer
    Height As Integer
    Wide As Integer
    More As Boolean
End Type

Public TimerWork As Boolean

Dim t As NOTIFYICONDATA

Dim MyForm As xForm

Dim OldWhwnd As Long
Dim DummyCount As Integer, ProCount As Integer
Private Sub CapC_Click()
    Dim Index As Long
    
    If CapC.ListIndex > -1 Then
        Index = FindForm("Cap", "Cap" & CapC.List(CapC.ListIndex))
        Forms(Index).Visible = True
        Call SetForegroundWindow(Forms(Index).Hwnd)
        Forms(Index).Blink.Enabled = True
    End If
    
    CapC.ListIndex = -1
End Sub
Private Sub Check_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Call FormMove(Me)
    Button = 0
End Sub
Private Sub CurrentColor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Call FormMove(Me)
End Sub
Private Sub EXEname_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    EXEname(0).ToolTipText = "Modul: " & EXEname(1).Caption
    EXEname(1).ToolTipText = EXEname(0).ToolTipText
    If Button = vbLeftButton Then Call FormMove(Me)
End Sub
Private Sub Form_Load()
    Dim P As Integer
    Dim R$, Value As String, SavedOS As String
    Dim ReadINI As Boolean, INIv As Boolean, OSwarn As Boolean
    Dim TaskID As Long, mHwnd As Long
    
    'On Local Error Resume Next
    
    mY.Name = "Mouse'n'Enum 3.60"
    
    mY.Path = App.Path
    If Right(mY.Path, 1) <> "\" Then mY.Path = mY.Path & "\"
    
    mY.EXE = mY.Path & App.EXEname & ".exe"
    mY.INI = mY.Path & "Me.ini"
    
    mY.OSstring = getOSversion
    
    Call GetWindowThreadProcessId(Me.Hwnd, TaskID)
    mY.TaskID = TaskID
    
    mY.Mail = "Woeh@gmx.de"
    mY.ActiveVB = "Http://www.goetz-reinecke.de"
    mY.VBtutor = "Http://www.activevb-archiv.net/vb/VBtutor/VBtut011.shtml"
    
    MyForm.Height = 4350
    MyForm.CPickerHeight = MyForm.Height + 900
    MyForm.MoreWide = 5520
    MyForm.Wide = 3555
    
    Me.Height = MyForm.Height
    Me.Width = MyForm.Wide

    For P = MenuItem.LBound To MenuItem.UBound
        R$ = LCase(MenuItem(P).Caption)
        If R$ = "minimiert starten" Then MainMenu.StartMin = P
        If R$ = "position speichern" Then MainMenu.SavePos = P
        If R$ = "save menus @ unload" Then MainMenu.SaveMenu = P
    Next P
    For P = ToolMenuItem.LBound To ToolMenuItem.UBound
        R$ = LCase(ToolMenuItem(P).Caption)
        If R$ = "maustooltip" Then MainMenu.ToolTip = P
        If R$ = "colorpicker" Then MainMenu.ColorP = P
        If R$ = "alle menus im system listen" Then MainMenu.AllMenu = P
    Next P
    
    Dummy.Show
    Load DummyDraw
    
    MenuItem(MainMenu.SaveMenu).Checked = True
    
    If Not FileExist(mY.INI) Then
        ReadINI = False
    Else
        If GetFromINI("mOption", "6", mY.INI) = "0" Then
            ReadINI = False
            MenuItem(MainMenu.SavePos).Checked = False
        Else
            ReadINI = True
        End If
    End If
        
    If Not ReadINI Then
        View(0) = False
        View(2) = False
        View(4) = True
        View(6) = False

        Infos(0) = True
        Infos(1) = False
        Infos(2) = False
        Infos(3) = False
        Infos(4) = False
        
        MenuItem(MainMenu.StartMin).Checked = False
        PosBorderItem(0).Checked = True
           
        For P = 0 To 10
            If P < 10 Then CapView(P) = True
            ViewItem(P).Checked = True
        Next P
        
        Cmenu(1) = False
        Cmenu(2) = False
        
        tMenu(2) = False
        tMenu(4) = False
        
        SavedOS = ""
        OSwarn = True
    Else
        For P = 2 To 4 Step 2
            MenuItem(P).Checked = GetFromINI("mOption", str(P), _
                                                             mY.INI)
        Next P
        
        If MenuItem(MainMenu.SavePos).Checked Then
            Me.Top = GetFromINI("MyPos", "Top", mY.INI)
            Me.Left = GetFromINI("MyPos", "Left", mY.INI)
        End If
        
        PosBorderItem(0).Checked = GetFromINI("PosBorder", _
                                              "PosBorder", mY.INI)
            
        If MenuItem(MainMenu.StartMin).Checked Then _
                                        Me.WindowState = vbMinimized
    
        For P = 0 To 10
            ViewItem(P).Checked = GetFromINI("MainView", str(P), _
                                                            mY.INI)
            If Not ViewItem(P).Checked Then _
                                    ViewItem(P).Checked = True: _
                                    Call ViewItem_Click(P)

            If P < 10 Then _
                CapView(P) = GetFromINI("CapView", CStr(P), mY.INI)
        
            If P = 0 Or P = 2 Or P = 4 Or P = 6 Then _
                View(P) = GetFromINI("View", str(P), mY.INI)
        
            If P < 5 Then _
                Infos(P) = GetFromINI("Infos", str(P), mY.INI)
        
            If P = 6 Or P = 8 Then _
                INIv = GetFromINI("tMenu", str(P), mY.INI): _
                If INIv Then Call ToolMenuItem_Click(P)
        
            If P = 1 Or P = 2 Then _
                Cmenu(1) = GetFromINI("cMenu", CStr(P), mY.INI)
                'Cmenu(2) = GetFromINI("cMenu", "2", My.INI)
        Next P
        
        SavedOS = GetFromINI("OSwarn", "SavedOS", mY.INI)
        OSwarn = GetFromINI("OSwarn", "OSwarn", mY.INI)
        
        
        ViewItem(0).Checked = True
    End If
    
    Select Case LCase(mY.OSstring)
        Case "windows nt", "windows 98", "windows 98 se", _
             "windows 95", "windows 32s"
            mY.OS = 1
            If OSwarn Or LCase(SavedOS) <> LCase(mY.OSstring) Then
                R$ = "Mouse'n'Enum wurde unter WindowsXP " & _
                     "entwickelt und läuft somit relativ " & _
                     "stabil und fehlerfrei unter WindowsXP, " & _
                     "Windows Millenium, und Windows 2000." & _
                     vbCrLf & vbCrLf & _
                     "Die Stabilität und Fehlerfreiheit ist " & _
                     "unter deiner Windowsversion " & _
                     "(" & mY.OSstring & ") nicht gehährleistet." & _
                     vbCrLf & vbCrLf & _
                     "Soll Mouse'n'Enum trotzdem ausgeführt " & _
                     "werden ?"
                If Ask(R$, "Hinweis", _
                       vbQuestion + vbYesNoCancel) <> vbYes Then End
            End If
        Case Else
            mY.OS = 2
    End Select
    
    Call Write2INI("OSwarn", "OSwarn", "0", mY.INI)
    If mY.OSstring <> SavedOS Then _
            Call Write2INI("OSwarn", "SavedOS", mY.OSstring, mY.INI)
    
    If GetFromINI("FirstStart", "FirstStart", mY.INI) = "" Then
        R$ = "Mouse'n'Enum wird zum erstenmal gestartet." & _
             vbCrLf & vbCrLf & _
             "Um ein Control zu manipulieren, zu enumieren etc. " & _
             vbCrLf & _
             "einfach die Mouse über das gewünschte Control " & _
             vbCrLf & "bewegen und F12 drücken." & _
             vbCrLf & vbCrLf & _
             "Bitte die gelben Menus beachten" & _
             vbCrLf & _
             "(Steckt ne Menge drin ;)"
        MsgBox R$, vbOKOnly + vbInformation, "Wilkommen"
        Call Write2INI("FirstStart", "FirstStart", _
                       CBool(False), mY.INI)
    End If
    
    Set fIcon.Move = MP(0).MouseIcon
    Set fIcon.Hand = MP(1).MouseIcon
    fIcon.Pointer = vbCustom
    
    mHwnd = GetSystemMenu(Me.Hwnd, False)
    Call RemoveMenu(mHwnd, 0, MF_BYPOSITION)
    Call RemoveMenu(mHwnd, 1, MF_BYPOSITION)
        
    Call LoadStandardForm(Me, False)

    For P = 0 To MenuLabel.UBound
        Call SetMP(MenuLabel(P), True)
    Next P
    
    Call SetMP(CapC, True)
    Call SetMP(ShowMore, True)
    
    For P = 1 To 9
        R$ = Label1(P).Caption
        ViewItem(P).Caption = Mid(R$, 1, Len(R$) - 1)
    Next P
    
    For P = 10 To ViewItem.UBound
        R$ = Label2(P - 10).Caption
        ViewItem(P).Caption = Mid(R$, 1, Len(R$) - 1)
    Next P
    
    MyForm.More = False
    Call ShowMore_Click
    
    t.cbSize = Len(t)
    t.Hwnd = pichook.Hwnd
    t.uId = 1&
    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    t.ucallbackMessage = WM_MOUSEMOVE
    t.hIcon = Me.Icon
    t.szTip = mY.Name & Chr$(0)
    Shell_NotifyIcon NIM_ADD, t
    App.TaskVisible = False
    
    ProCount = 100
    Call GetHwnd_Timer

    SetTimer Me.Hwnd, 0, 1, AddressOf TimerProc
    
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Call FormMove(Me)
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    t.cbSize = Len(t)
    t.Hwnd = pichook.Hwnd
    t.uId = 1&
    Shell_NotifyIcon NIM_DELETE, t
End Sub
Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Me.WindowState = vbNormal: _
                                         Me.Visible = False: _
                                         Me.Height = MyForm.Height: _
                                         Me.Width = MyForm.Wide
    Call FadeForm(Me, 0, 255)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim P As Integer
    Dim R$, Value As String
    
    On Local Error Resume Next
    
    KillTimer Me.Hwnd, 0
    
    GetHwnd.Enabled = False
    
    Unload DummyDraw
    Unload Dummy
    
    For P = 2 To MenuItem.UBound Step 2
        If P = MainMenu.SaveMenu Or (P < MainMenu.SaveMenu _
           And MenuItem(MainMenu.SaveMenu).Checked) Then _
                Call Write2INI("mOption", CStr(P), _
                               CBool(MenuItem(P).Checked), mY.INI)
    Next P
    
    If MenuItem(MainMenu.SaveMenu).Checked Then
        Call Write2INI("PosBorder", "PosBorder", _
                       CBool(PosBorderItem(0).Checked), mY.INI)

        If MenuItem(MainMenu.SavePos).Checked Then _
            Call Write2INI("MyPos", "Left", CStr(Me.Left), mY.INI): _
            Call Write2INI("MyPos", "Top", CStr(Me.Top), mY.INI)
                
        For P = 0 To 10
            Call Write2INI("MainView", CStr(P), _
                           CBool(ViewItem(P).Checked), mY.INI)
                           
            If P < 10 Then _
                Call Write2INI("CapView", CStr(P), _
                               CBool(CapView(P)), mY.INI)
            
            If P = 0 Or P = 2 Or P = 4 Or P = 8 Then _
                Call Write2INI("View", CStr(P), _
                               CBool(View(P)), mY.INI)
            If P < 5 Then _
                Call Write2INI("Infos", CStr(P), _
                               CBool(Infos(P)), mY.INI)
            
            If P And P < 3 Then _
                Call Write2INI("cMenu", CStr(P), _
                               CBool(Cmenu(P)), mY.INI)
            
            If P = 6 Or P = 8 Then _
                Call Write2INI("tMenu", CStr(P), _
                               CBool(ToolMenuItem(P).Checked), _
                               mY.INI)
        Next P
    End If
    
    Call ShakeForm(Me, 5, 1000)
    
    End

End Sub
Public Sub GetHwnd_Timer()
    Dim nPoint As POINTAPI
    Dim Shex As String, t$
    Dim Hdcp As Long, Color As Long, Style As Long, Index As Long
    Dim r1 As Long, r2 As Long, Prop As Long, X As Long
    Dim R As RECT
    Dim isTopMost As Boolean
    
    Static oldR As RECT
           
    TimerWork = True
    
    'On Local Error Resume Next
    
    GetCursorPos nPoint
    
    wHwnd = WindowFromPoint(nPoint.X, nPoint.Y)
    
    eWin.Hwnd = wHwnd
    eWin.ParentHwnd = GetParent(eWin.Hwnd)
    eWin.Class = GetClass(eWin.Hwnd)
    eWin.Thread = GetWindowThreadProcessId(eWin.Hwnd, eWin.TaskID)
    
    X = SearchLB(lbID.Hwnd, LB_FINDSTRINGEXACT, -1, eWin.TaskID)
    
    If X > -1 Then
        EXEname(1).Caption = LB.List(X)
    Else
        EXEname(1).Caption = "Not found"
    End If

    If ToolMenuItem(MainMenu.ToolTip).Checked Then _
            Dummy.DummyCap.Caption = eWin.Hwnd & vbCrLf & _
                                            eWin.Class: _
            Call MoveWindow(Dummy.Hwnd, nPoint.X + 15, nPoint.Y, _
                            Dummy.DummyCap.Width + 5, 27, 1): _
                   Dummy.Cls: _
                   Dummy.Print eWin.Hwnd & vbCrLf & eWin.Class
    
    If CurrentColor.Visible Then
        Hdcp = GetDC(eWin.Hwnd)
        Call ScreenToClient(eWin.Hwnd, nPoint)
        Color = GetPixel(Hdcp, nPoint.X, nPoint.Y)
    
        If Color = -1 Then
            CurrentColor.BackColor = 0
            Label(3).Caption = "Error"
            Label(4).Caption = "Error"
            Label(5).Caption = "Error"
        Else
            CurrentColor.BackColor = Color
            Label(3).Caption = Color
        
            Shex = Hex(Color)
            Label(4).Caption = Shex
        
            If Len(Shex) < 6 Then Shex = String(6 - Len(Shex), "0") & Shex
            Label(5).Caption = "#" & Right$(Shex, 2) & _
                               Mid$(Shex, 3, 2) & Left$(Shex, 2)
        End If
    End If
    
    MouseK.Caption = "x=" & nPoint.X & ", y=" & nPoint.Y
    
    Call GetWindowRect(eWin.Hwnd, R)
    
    eWin.Text = GetText(eWin.Hwnd)
    eWin.Left = R.Left
    eWin.Top = R.Top
    eWin.Width = R.Right - R.Left
    eWin.Heigth = R.Bottom - R.Top
    eWin.ParentClass = GetClass(eWin.ParentHwnd)
    eWin.ParentText = GetText(eWin.ParentHwnd)
    eWin.Visible = IsWindowVisible(eWin.Hwnd)
    
    If PosBorderItem(0).Checked Then
        If (oldR.Bottom <> R.Bottom Or oldR.Left <> R.Left _
            Or oldR.Right <> R.Right Or oldR.Top <> R.Top) _
            And wHwnd <> DummyDraw.Hwnd Then

            DummyCount = 0
            Call SetRect(oldR, R.Left, R.Top, R.Right, R.Bottom)

            Call ShowWindow(DummyDraw.Hwnd, SW_HIDE)

            Call SetWindowRgn(DummyDraw.Hwnd, 0, False)
    
            Call MoveWindow(DummyDraw.Hwnd, R.Left - 3, R.Top - 3, _
                            R.Right - R.Left + 4, eWin.Heigth + 5, 0)
        
            r1 = CreateRectRgn(1, 1, _
                    DummyDraw.Width / Screen.TwipsPerPixelX - 1, _
                    DummyDraw.Height / Screen.TwipsPerPixelY)
            r2 = CreateRectRgn(3, 3, _
                    DummyDraw.Width / Screen.TwipsPerPixelX - 3, _
                    DummyDraw.Height / Screen.TwipsPerPixelY - 2)
    
            Call CombineRgn(r2, r1, r2, RGN_XOR)
                        
            Call SetWindowRgn(DummyDraw.Hwnd, r2, True)
            
            Call ShowWindow(DummyDraw.Hwnd, SW_SHOWNOACTIVATE)
        Else
            If DummyCount = 30 Then
                Call ShowWindow(DummyDraw.Hwnd, SW_HIDE)
            Else
                If DummyCount = 60 Then
                    Call ShowWindow(DummyDraw.Hwnd, SW_SHOWNOACTIVATE)
                    DummyCount = 0
                End If
            End If
        End If
        DummyCount = DummyCount + 1
    End If
    
    Label1(10).Caption = eWin.Hwnd
    Label1(11).Caption = eWin.Thread
    Label1(12).Caption = eWin.TaskID
    Label1(13).Caption = eWin.Class
    Label1(14).Caption = eWin.Text
    Label1(15).Caption = "x = " & eWin.Left & ", y = " & eWin.Top
    Label1(16).Caption = "Width: " & eWin.Width & ", " & _
                         "Height: " & eWin.Heigth
    If eWin.ParentHwnd Then
        Label1(17).Caption = eWin.ParentHwnd
        Label1(17).ForeColor = &HFFFFFF
        Label1(18).Caption = eWin.ParentClass
        Label1(19).Caption = eWin.ParentText
        
        eWin.TopParent = eWin.ParentHwnd
        
        Do Until GetParent(eWin.TopParent) = 0
            eWin.TopParent = GetParent(eWin.TopParent)
        Loop
        
        Label2(10).Caption = eWin.TopParent
    Else
        Label1(17).Caption = "Control is Parent"
        Label1(17).ForeColor = &HFF00&
        Label1(18).Caption = "---"
        Label1(19).Caption = "---"
        Label2(10).Caption = "---"
    End If
    
    If OldWhwnd <> wHwnd Then
        If MyForm.More Then
            Style = GetWindowLongA(eWin.Hwnd, GWL_STYLE)
            Prop = GetWindowLongA(eWin.Hwnd, GWL_EXSTYLE)

            Check(0).Value = CBool(Prop And WS_EX_MDICHILD) * -1
            Check(1).Value = CBool(Prop And WS_EX_APPWINDOW) * -1
            Check(2).Value = CBool(Prop And WS_EX_TOOLWINDOW) * -1
            Check(3).Value = (Not CBool(Style And WS_DISABLED)) * -1
            Check(4).Value = CBool(Prop And WS_EX_TOPMOST) * -1
            Check(5).Value = CBool(Style And WS_DLGFRAME) * -1
            Check(6).Value = CBool(Style And WS_MINIMIZEBOX) * -1
            Check(7).Value = CBool(Style And WS_MAXIMIZEBOX) * -1
            Check(8).Value = CBool(Style And WS_SYSMENU) * -1
            'Check(9).Value = CBool(Style And WS_CHILD) * -1
            If eWin.ParentHwnd Then
                Check(9).Value = 1
                Check(10).Value = 0
            Else
                Check(9).Value = 0
                Check(10).Value = 1
            End If
        
            If Style And WS_MINIMIZE Then
                LMore(11).Caption = "Min."
            Else
                If Style And WS_MAXIMIZE Then
                    LMore(11).Caption = "Max."
                Else
                    LMore(11).Caption = "Norm."
                End If
            End If
        End If
    
        OldWhwnd = wHwnd
        
        If Cchilds.Show And Not Cchilds.Stopped Then _
            Index = FindForm("Ac", "cRefresh"): _
            Call Forms(Index).GetHwnd(eWin.Hwnd, False)
          
        If MyMenu.Show And Not MyMenu.Stopped Then _
            Index = FindForm("Ac", "mRefresh"): _
            Call Forms(Index).MenuEnum(eWin.Hwnd, False, False)
    End If

    ProCount = ProCount + 1
    If ProCount > 49 Then ProCount = 0: _
                          Call GetAllProcess
        
    TimerWork = False
    
End Sub
Private Sub HelpItem_Click(Index As Integer)
    Dim Item As String
    
    Item = LCase(HelpItem(Index).Caption)

    Select Case Item
        Case "activevb"
            Call PageVisit(mY.ActiveVB)
        Case "e-mail an den autor"
            Call SendMail(mY.Mail)
        Case "was sagt mir das prog ?", "über"
            Help.Fade = False
            If Item = "was sagt mir das prog ?" Then Help.Fade = True
            Help.Show
        Case "history"
            Call OpenEXE(mY.Path & "History.doc")
    End Select
    
End Sub
Private Sub Label_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Call FormMove(Me)
End Sub
Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Call FormMove(Me)
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
Private Sub Label2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Call FormMove(Me)
End Sub
Private Sub Label2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Cap As String
    
    If Index > 9 Then Index = Index - 10
    
    Cap = Label2(Index).Caption & " " & Label2(Index + 10).Caption
    
    Label2(Index).ToolTipText = Cap
    Label2(Index + 10).ToolTipText = Cap
    
End Sub
Private Sub LMore_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Call FormMove(Me)
End Sub
Private Sub MenuItem_Click(Index As Integer)

    Select Case LCase(MenuItem(Index).Caption)
        Case "save menus @ unload", _
             "minimiert starten", _
             "position speichern"
            MenuItem(Index).Checked = Not (MenuItem(Index).Checked)
        Case "exit"
            Unload Me
    End Select
    
End Sub
Private Sub MenuLabel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = vbLeftButton Then
        Select Case Index
            Case 0
                PopupMenu Menu
            Case 1
                PopupMenu ToolMenu
            Case 2
                PopupMenu HelpMenu
            Case 3
                PopupMenu SearchMenu
        End Select
    End If

End Sub
Private Sub MouseK_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Call FormMove(Me)
End Sub
Private Sub pichook_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static Rec As Boolean
    Static MSG As Long
    
    MSG = X / Screen.TwipsPerPixelX
    
    If Not Rec Then
        Rec = True
        If MSG = WM_LBUTTONDOWN Then
            If Me.Visible Then
                Me.Visible = False
            Else
                Me.Width = MyForm.Wide
                If MyForm.More Then Me.Width = 5520
                
                Me.Height = MyForm.Height
                If CurrentColor.Visible Then Me.Height = MyForm.CPickerHeight
                
                Me.Visible = True
            End If
        End If
        Rec = False
    End If
    
End Sub
Public Function FindCBitem(Item As String, Del As Boolean) _
                                                        As Integer
    Dim P As Integer
    
    Item = Trim(Item)
        
    For P = 0 To CapC.ListCount - 1
        If CapC.List(P) = Item Then
            If Del Then
                CapC.RemoveItem P
                CapC.ListIndex = -1
                If CapC.ListCount = 0 Then CapC.Enabled = False
            End If
            FindCBitem = P
            Exit Function
        End If
    Next P
    
    FindCBitem = -1
    
End Function
Private Sub PosBorderItem_Click(Index As Integer)
    PosBorderItem(0).Checked = Not (PosBorderItem(0).Checked)
    
    If Not PosBorderItem(0).Checked Then
        DummyDraw.Hide
    Else
        DummyDraw.Show
    End If
    
    DummyCount = 0

End Sub
Private Sub RefreshMenuItem_Click(Index As Integer)
    Dim A As New Ac
    
    Select Case Index
        Case 0
            If RefreshMenuItem(Index).Checked Then
                Cchilds.Show = False
                RefreshMenuItem(Index).Checked = False
                Unload Forms(FindForm("Ac", "cRefresh"))
            Else
                RefreshMenuItem(Index).Checked = True
                Load A
                Call A.AutoRefresh(True)
            End If
        Case 1
            If RefreshMenuItem(Index).Checked Then
                RefreshMenuItem(Index).Checked = False
                MyMenu.Show = False
                Unload Forms(FindForm("Ac", "mRefresh"))
            Else
                RefreshMenuItem(Index).Checked = True
                Load A
                A.MenuE = True
                Call A.AutoRefresh(True)
            End If

    End Select

End Sub
Private Sub SearchItem_Click(Index As Integer)
    Dim Prompt As String, Title As String, Find As String
    Dim FindClass As Boolean
    Dim A As New Ac
    
    Title = SearchItem(Index).Caption
    
    Select Case LCase(Title)
        Case "search windows by class"
            Prompt = "Type the String of Classname to find"
            FindClass = True
        Case "search windows by text"
            Prompt = "Type the String of Text to find"
            FindClass = False
    End Select
    
    If Prompt <> "" Then
        Find = uInput(Prompt, Title)
        If Find <> "" Then
            Load A
            Call A.SearchHwnd(Find, FindClass)
        End If
    End If
    
End Sub

Private Sub ShowMore_Click()
    Dim P As Integer, X As Integer
        
    If ShowMore.Caption = "<" Then
        ShowMore.Caption = ">"
        ShowMore.ToolTipText = "Erweiterte Ansicht"
        MyForm.More = False
        
        Call RollForm(Me, Me.Height, Me.Height, Me.Width, _
                      MyForm.Wide, False, True, 0, 255)
        
        Linie.Visible = False

        For P = 0 To LMore.UBound
            LMore(P).Visible = False
            If P <= Check.UBound Then Check(P).Visible = False
        Next P
        
    Else
        ShowMore.Caption = "<"
        ShowMore.ToolTipText = "Einfache Ansicht"
        MyForm.More = True
        
        Linie.Visible = True
        
        For P = 0 To LMore.UBound
            LMore(P).Visible = True
            If P <= Check.UBound Then Check(P).Visible = True
        Next P
        
        OldWhwnd = 0
        Call GetHwnd_Timer
        
        Call RollForm(Me, Me.Height, Me.Height, Me.Width, _
                      MyForm.MoreWide, True, True, 0, 255)
    End If
    
End Sub
Private Sub TaskMenuItem_Click(Index As Integer)
    Dim A As New Ac, C As New Cap
    Dim mHwnd As Long
    Dim R$
    
    On Local Error Resume Next
    
    Select Case LCase(TaskMenuItem(Index).Caption)
        Case "alle tasks enumieren"
            Load A
            Call A.GetHwnd(0, True)
        Case "alle tasks mit childs enumieren"
            Load A
            Call A.GetHwnd(-1, False)
        Case "manuelle eingabe"
            Err.Clear
            R$ = "Bitte gebe das gewünschte hWnd (vom Typ Long) ein."
            mHwnd = CLng(uInput(R$, "Manuelle Eingabe"))
            If Err.Number = 13 Then
                R$ = "Die Eingabe von dir war keine Zahl vom Typ Long."
                Call uMsg(R$, "Ungültige Eingabe", vbCritical + vbOKOnly)
            Else
                Load C
                Call C.GetStaticInfo(mHwnd)
                Call C.GetHwnd(mHwnd)
                C.Visible = True
            End If
        Case "prozesse listen"
            AllPro.Show
    End Select
    
End Sub
Private Sub ToolMenuItem_Click(Index As Integer)
    Dim mHwnd As Long
    Dim P As Integer
    Dim A As New Ac
    
    On Local Error Resume Next
    
    Select Case LCase(ToolMenuItem(Index).Caption)
        Case "alle menus im system listen"
            If ToolMenuItem(Index).Checked Then
                Unload Forms(FindForm("Ac", "Allmenu"))
            Else
                ToolMenuItem(Index).Checked = True
                Load A
                A.Tag = "AllMenu"
                'Call A.GetHwnd(-2, True)
                Call A.EnumAllWindowsMenu
            End If
        Case "combo- + listboxen"
            ToolMenuItem(Index).Checked = Not (ToolMenuItem(Index).Checked)
            If ToolMenuItem(Index).Checked Then
                CbLB.Show
            Else
                Unload CbLB
            End If
        Case "maustooltip"
            If ToolMenuItem(Index).Checked Then
                ToolMenuItem(Index).Checked = False
                Dummy.Hide
            Else
                ToolMenuItem(Index).Checked = True
                Dummy.Show
            End If
        Case "colorpicker"
            If ToolMenuItem(Index).Checked Then
                ToolMenuItem(Index).Checked = False
                Call RollForm(Me, Me.Height, MyForm.Height, _
                              Me.Width, Me.Width, False, False, 0, 0)
                CurrentColor.Visible = False
                For P = 0 To 5
                    Label(P).Visible = False
                Next P
            Else
                ToolMenuItem(Index).Checked = True
                CurrentColor.Visible = True
                For P = 0 To 5
                    Label(P).Visible = True
                Next P
                
                Call GetHwnd_Timer
        
                Call RollForm(Me, Me.Height, MyForm.CPickerHeight, _
                              Me.Width, Me.Width, True, False, 0, 0)
            End If
        Case "neues fenster"
            Call OpenEXE(mY.EXE)
    End Select
    
End Sub
Private Sub ViewItem_Click(Index As Integer)
    Dim t As Integer, OldH As Integer
    
    If ViewItem(Index).Checked Then
        ViewItem(Index).Checked = False
        If Index < 10 Then
            Label1(Index).Visible = False
            Label1(Index + 10).Visible = False
        Else
            Label2(Index - 10).Visible = False
            Label2(Index).Visible = False
        End If
    Else
        ViewItem(Index).Checked = True
        If Index < 10 Then
            Label1(Index).Visible = True
            Label1(Index + 10).Visible = True
        Else
            Label2(Index - 10).Visible = True
            Label2(Index).Visible = True
        End If
    End If
    
    OldH = Me.Height
    t = Label1(1).Top
    
    For P = 1 To ViewItem.UBound - 1
        If Label1(P).Visible Then Label1(P).Top = t: _
                                  Label1(P + 10).Top = t: _
                                  t = t + Label1(P).Height - 1
    Next P
    
    For P = ViewItem.UBound To ViewItem.UBound
        If Label2(P - 10).Visible Then Label2(P - 10).Top = t: _
                                       Label2(P).Top = t: _
                                       t = t + Label2(P).Height - 1
    Next P

    
    CapC.Top = t + 3
    MouseK.Top = t + 3
    
    t = CapC.Top + CapC.Height + 100
    
    CurrentColor.Top = t
    
    For P = 0 To 2
        Label(P).Top = t
        Label(P + 3).Top = t
        t = t + Label(P).Height - 1
    Next P
    
    If Not CurrentColor.Visible Then
        t = CapC.Top + CapC.Height + 600
    Else
        t = CurrentColor.Top + CurrentColor.Height + 600
    End If
    
    MyForm.CPickerHeight = CurrentColor.Top + _
                           CurrentColor.Height + 600
    MyForm.Height = t
    Me.Height = t
    
End Sub
