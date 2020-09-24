Attribute VB_Name = "BoxModul"
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal MSG As Long, wParam As Any, lParam As Any) As Long

Private Const CB_ADDSTRING = &H143
Private Const CB_GETCURSEL = &H147
Private Const CB_GETDROPPEDSTATE = &H157
Private Const CB_GETLBTEXT = &H148
Private Const CB_GETLBTEXTLEN = &H149
Private Const CB_GETCOUNT = &H146
Private Const CB_INSERTSTRING = &H14A
Private Const CB_SETCURSEL = &H14E
Private Const CB_RESETCONTENT = &H14B
Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_DELETESTRING = &H144

Private Const LB_ADDSTRING = &H180
Private Const LB_GETCOUNT = &H18B
Private Const LB_GETTEXTLEN = &H18A
Private Const LB_GETTEXT = &H189
Private Const LB_GETCURSEL = &H188
Private Const LB_INSERTSTRING = &H181
Private Const LB_SETCURSEL = &H186
Private Const LB_RESETCONTENT = &H184
Private Const LB_DELETESTRING = &H182
Public Function ReadCB(ByVal Hwnd As Long, Box2Drope As ComboBox) _
                                                            As Long
    Dim TextLen As Long, Count As Long, i As Long
    Dim ItemText As String
    
    Count = SendMessage(Hwnd, CB_GETCOUNT, ByVal CLng(0), ByVal CLng(0))
    
    If Count Then
        For i = 0 To Count - 1
            TextLen = SendMessage(Hwnd, CB_GETLBTEXTLEN, ByVal CLng(i), ByVal CLng(0))
            ItemText = Space(TextLen) & vbNullChar
            TextLen = SendMessage(Hwnd, CB_GETLBTEXT, ByVal CLng(i), ByVal ItemText)
            ItemText = Left(ItemText, TextLen)
            Box2Drope.AddItem ItemText
        Next i
        Box2Drope.ListIndex = GetLBlistIndex(Hwnd, True)
    End If
    
    ReadCB = Count

End Function
Public Function ReadLB(ByVal Hwnd As Long, Box2Drope As ListBox) _
                                                            As Long
    Dim TextLen As Long, Count As Long, i As Long
    Dim ItemText As String
    
    Count = SendMessage(Hwnd, LB_GETCOUNT, ByVal CLng(0), _
                                            ByVal CLng(0))
    
    If Count Then
        For i = 0 To Count - 1
            TextLen = SendMessage(Hwnd, LB_GETTEXTLEN, ByVal CLng(i), ByVal CLng(0))
            ItemText = Space(TextLen) & vbNullChar
            TextLen = SendMessage(Hwnd, LB_GETTEXT, ByVal CLng(i), ByVal ItemText)
            ItemText = Left(ItemText, TextLen)
            Box2Drope.AddItem ItemText
        Next i
        Box2Drope.ListIndex = GetLBlistIndex(Hwnd, False)
    End If
    
    ReadLB = Count
    
End Function
Public Function GetLBlistIndex(ByVal Hwnd As Long, Combo As Boolean) _
                                                            As Long
    If Combo Then
        GetLBlistIndex = SendMessage(Hwnd, CB_GETCURSEL, _
                                     ByVal CLng(0), ByVal CLng(0))
    Else
        GetLBlistIndex = SendMessage(Hwnd, LB_GETCURSEL, _
                                     ByVal CLng(0), ByVal CLng(0))
    End If
    
End Function
Public Function SetLBlistIndex(ByVal Hwnd, ByVal Index As Integer, _
                               Combo As Boolean) As Long
    If Combo Then
        Call SendMessage(Hwnd, CB_SETCURSEL, _
                         ByVal CLng(Index), ByVal CLng(1))
    Else
        Call SendMessage(Hwnd, LB_SETCURSEL, _
                         ByVal CLng(Index), ByVal CLng(1))
    End If
End Function
Public Function AddLBitem(ByVal Hwnd As Long, ByVal Index As Integer, _
                          ByVal TXT As String, Combo As Boolean) _
                                                             As Long
    If Combo Then
        Call SendMessage(Hwnd, CB_INSERTSTRING, ByVal CLng(Index), _
                                                         ByVal TXT)
    Else
        Call SendMessage(Hwnd, LB_INSERTSTRING, ByVal CLng(Index), _
                                                         ByVal TXT)
    End If
End Function
Public Function DeleteLBitem(ByVal Hwnd As Long, ByVal Index As Integer, _
                                                   Combo As Boolean)
    If Combo Then
        Call SendMessage(Hwnd, CB_DELETESTRING, ByVal CLng(Index), _
                                                      ByVal CLng(0))
    Else
        Call SendMessage(Hwnd, LB_DELETESTRING, ByVal CLng(Index), _
                                                      ByVal CLng(0))
    End If
End Function
Public Function ClearLB(ByVal Hwnd As Long, Combo As Boolean)
    If Combo Then
        Call SendMessage(Hwnd, CB_RESETCONTENT, _
                        ByVal CLng(0), ByVal CLng(0))
    Else
        Call SendMessage(Hwnd, LB_RESETCONTENT, _
                         ByVal CLng(0), ByVal CLng(0))
    End If
End Function
Public Function CBisDropped(ByVal Hwnd As Long) As Long
    CBisDropped = SendMessage(Hwnd, CB_GETDROPPEDSTATE, _
                              ByVal CLng(0), ByVal CLng(0))
End Function
Public Sub CBdropdown(ByVal Hwnd As Long)
    
    Call SendMessage(Hwnd, CB_SHOWDROPDOWN, ByVal CLng(1), 0)

    Do Until CBisDropped(Hwnd) = 1
        DoEvents
    Loop
    
End Sub
