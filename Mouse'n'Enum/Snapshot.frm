VERSION 5.00
Begin VB.Form SnapShot 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Festes Werkzeugfenster
   ClientHeight    =   4080
   ClientLeft      =   1125
   ClientTop       =   1380
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   4080
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command 
      Caption         =   "Print"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "Save"
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   4
      Top             =   3720
      Width           =   855
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   4335
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3255
      Left            =   4440
      TabIndex        =   2
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'Kein
      Height          =   3255
      Left            =   120
      ScaleHeight     =   3255
      ScaleWidth      =   4335
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'Kein
         Height          =   4335
         Left            =   120
         ScaleHeight     =   4335
         ScaleWidth      =   6015
         TabIndex        =   1
         Top             =   240
         Width           =   6015
      End
   End
End
Attribute VB_Name = "SnapShot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type PALETTEENTRY
    peRed   As Byte
    peGreen As Byte
    peBlue  As Byte
    peFlags As Byte
End Type

Private Type LOGPALETTE
    palVersion       As Integer
    palNumEntries    As Integer
    palPalEntry(255) As PALETTEENTRY
End Type

Private Type GUID
    Data1    As Long
    Data2    As Integer
    Data3    As Integer
    Data4(7) As Byte
End Type

Private Const RASTERCAPS  As Long = 38
Private Const RC_PALETTE  As Long = &H100
Private Const SIZEPALETTE As Long = 104

Private Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "GDI32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDeviceCaps Lib "GDI32" (ByVal hDC As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "GDI32" (ByVal hDC As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function CreatePalette Lib "GDI32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function SelectObject Lib "GDI32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "GDI32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function GetForegroundWindow Lib "USER32" () As Long
Private Declare Function SelectPalette Lib "GDI32" (ByVal hDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function GetWindowDC Lib "USER32" (ByVal Hwnd As Long) As Long
Private Declare Function GetDC Lib "USER32" (ByVal Hwnd As Long) As Long
Private Declare Function GetWindowRect Lib "USER32" (ByVal Hwnd As Long, lpRect As RECT) As Long
Private Declare Function ReleaseDC Lib "USER32" (ByVal Hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDesktopWindow Lib "USER32" () As Long

Private Type PicBmp
    Size     As Long
    Type     As Long
    hBmp     As Long
    hPal     As Long
    Reserved As Long
End Type

Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Public Function CreateBitmapPicture(ByVal hBmp As Long, _
                                    ByVal hPal As Long) As Picture
    Dim R As Long
    Dim Pic As PicBmp
    Dim IPic As IPicture
    Dim IID_IDispatch As GUID

    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With

    With Pic
        .Size = Len(Pic)
        .Type = vbPicTypeBitmap
        .hBmp = hBmp
        .hPal = hPal
    End With

    R = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)

    Set CreateBitmapPicture = IPic
    
End Function
Public Function CaptureWindow(ByVal hWndSrc As Long, _
                              ByVal Client As Boolean, _
                              ByVal LeftSrc As Long, _
                              ByVal TopSrc As Long, _
                              ByVal WidthSrc As Long, _
                              ByVal HeightSrc As Long) As Picture
    Dim hDCMemory As Long, hBmp As Long, hBmpPrev As Long
    Dim R As Long, hDCSrc As Long, hPal As Long, hPalPrev As Long
    Dim RasterCapsScrn As Long, HasPaletteScrn As Long
    Dim PaletteSizeScrn As Long

    Dim LogPal As LOGPALETTE

    hDCSrc = GetWindowDC(hWndSrc)
    
    hDCMemory = CreateCompatibleDC(hDCSrc)
    hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
    hBmpPrev = SelectObject(hDCMemory, hBmp)

    RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS)
    HasPaletteScrn = RasterCapsScrn And RC_PALETTE
    PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE)
    
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        LogPal.palVersion = &H300
        LogPal.palNumEntries = 256
        R = GetSystemPaletteEntries(hDCSrc, 0, 256, _
                                    LogPal.palPalEntry(0))
        hPal = CreatePalette(LogPal)
        hPalPrev = SelectPalette(hDCMemory, hPal, 0)
        R = RealizePalette(hDCMemory)
    End If

    R = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, _
               hDCSrc, LeftSrc, TopSrc, vbSrcCopy)

    hBmp = SelectObject(hDCMemory, hBmpPrev)

    If HasPaletteScrn And (PaletteSizeScrn = 256) Then _
            hPal = SelectPalette(hDCMemory, hPalPrev, 0)
    
    R = DeleteDC(hDCMemory)
    R = ReleaseDC(hWndSrc, hDCSrc)

    Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)

End Function
Public Function CaptureActiveWindow(ByVal Hwnd As Long) As Picture
    Dim hWndActive As Long
    Dim R As Long
    Dim RectActive As RECT

    hWndActive = Hwnd

    R = GetWindowRect(hWndActive, RectActive)

    Set CaptureActiveWindow = _
              CaptureWindow(hWndActive, False, 0, 0, _
                            RectActive.Right - RectActive.Left, _
                            RectActive.Bottom - RectActive.Top)
                            
End Function
Public Sub PrintPictureToFitPage(Prn As Printer, Pic As Picture)
    Dim PicRatio As Double, PrnWidth As Double, PrnHeight As Double
    Dim PrnRatio As Double, PrnPicWidth As Double
    Dim PrnPicHeight As Double
    
    Const vbHiMetric As Integer = 8

    If Pic.Height >= Pic.Width Then
        Prn.Orientation = vbPRORPortrait
    Else
        Prn.Orientation = vbPRORLandscape
    End If

    PicRatio = Pic.Width / Pic.Height

    PrnWidth = Prn.ScaleX(Prn.ScaleWidth, Prn.ScaleMode, vbHiMetric)
    PrnHeight = Prn.ScaleY(Prn.ScaleHeight, Prn.ScaleMode, vbHiMetric)

    PrnRatio = PrnWidth / PrnHeight

    If PicRatio >= PrnRatio Then
        PrnPicWidth = Prn.ScaleX(PrnWidth, vbHiMetric, Prn.ScaleMode)
        PrnPicHeight = Prn.ScaleY(PrnWidth / PicRatio, vbHiMetric, _
                                  Prn.ScaleMode)
    Else
        PrnPicHeight = Prn.ScaleY(PrnHeight, vbHiMetric, Prn.ScaleMode)
        PrnPicWidth = Prn.ScaleX(PrnHeight * PicRatio, vbHiMetric, _
                                 Prn.ScaleMode)
    End If

    Prn.PaintPicture Pic, 0, 0, PrnPicWidth, PrnPicHeight

End Sub
Public Sub GetWindowHwnd(ByVal Hwnd As Long)
    Dim R As RECT
    
    Picture2.AutoRedraw = True

    Set Picture2.Picture = CaptureActiveWindow(Hwnd)
    
    Call GetWindowRect(Hwnd, R)
    
    Picture2.Width = (R.Right - R.Left) * 15
    Picture2.Height = (R.Bottom - R.Top) * 15

    VScroll1.LargeChange = Picture1.Height / 4
    VScroll1.SmallChange = 120
    HScroll1.LargeChange = Picture1.Width / 4
    HScroll1.SmallChange = 120

    VScroll1.Max = Picture2.Height - Picture1.Height + 15
    HScroll1.Max = Picture2.Width - Picture1.Width + 15

    Me.Visible = True
    
End Sub
Private Sub Command_Click(Index As Integer)
    Dim fName As String
    Dim X As Integer
    
    Select Case Index
        Case 0
            X = 1
            fName = My.Path & "SnapShot_"
            
            Do Until Not FileExist(fName & X & ".bmp")
                X = X + 1
            Loop
            
            SavePicture Picture2.Image, fName & X & ".bmp"
        Case 1
            PrintPictureToFitPage Printer, Picture2.Picture
            Printer.EndDoc
    End Select
    
End Sub
Private Sub Form_Activate()
    Call FadeForm(Me, 0, 255)
End Sub
Private Sub Form_Load()
    Dim P As Integer
    
    Call LoadStandardForm(Me, True)
    Me.Caption = My.Name & " - " & Me.Name
    
    Call SetMP(VScroll1, True)
    Call SetMP(HScroll1, True)
        
    For P = Command.LBound To Command.UBound
        Call SetMP(Command(P), True)
    Next P
    
    Picture2.Top = 0
    Picture2.Left = 0
  
    VScroll1.LargeChange = Picture2.Height / 4
    VScroll1.SmallChange = 120
    
    HScroll1.LargeChange = Picture2.Width / 4
    HScroll1.SmallChange = 120
    
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Call FormMove(Me)
End Sub
Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Call FormMove(Me)
End Sub
Private Sub HScroll1_Change()
    Picture2.Left = -HScroll1.Value
End Sub
Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Call FormMove(Me)
End Sub
Private Sub VScroll1_Change()
    Picture2.Top = -VScroll1.Value
End Sub
