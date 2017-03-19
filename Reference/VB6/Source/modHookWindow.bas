Attribute VB_Name = "modHookWindow"
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal Length As Long)

Public Enum WM
     WM_NULL = &H0
     WM_CREATE = &H1
     WM_DESTROY = &H2
     WM_MOVE = &H3
     WM_SIZE = &H5
     WM_ACTIVATE = &H6
     WM_SETFOCUS = &H7
     WM_KILLFOCUS = &H8
     WM_ENABLE = &HA
     WM_SETREDRAW = &HB
     WM_SETTEXT = &HC
     WM_GETTEXT = &HD
     WM_GETTEXTLENGTH = &HE
     WM_PAINT = &HF
     WM_CLOSE = &H10
     WM_QUERYENDSESSION = &H11
     WM_QUIT = &H12
     WM_QUERYOPEN = &H13
     WM_ERASEBKGND = &H14
     WM_SYSCOLORCHANGE = &H15
     WM_ENDSESSION = &H16
     WM_SHOWWINDOW = &H18
     WM_WININICHANGE = &H1A
     WM_DEVMODECHANGE = &H1B
     WM_ACTIVATEAPP = &H1C
     WM_FONTCHANGE = &H1D
     WM_TIMECHANGE = &H1E
     WM_CANCELMODE = &H1F
     WM_SETCURSOR = &H20
     WM_MOUSEACTIVATE = &H21
     WM_CHILDACTIVATE = &H22
     WM_QUEUESYNC = &H23
     WM_GETMINMAXINFO = &H24
     WM_PAINTICON = &H26
     WM_ICONERASEBKGND = &H27
     WM_NEXTDLGCTL = &H28
     WM_SPOOLERSTATUS = &H2A
     WM_DRAWITEM = &H2B
     WM_MEASUREITEM = &H2C
     WM_DELETEITEM = &H2D
     WM_VKEYTOITEM = &H2E
     WM_CHARTOITEM = &H2F
     WM_SETFONT = &H30
     WM_GETFONT = &H31
     WM_SETHOTKEY = &H32
     WM_GETHOTKEY = &H33
     WM_QUERYDRAGICON = &H37
     WM_COMPAREITEM = &H39
     WM_COMPACTING = &H41
     WM_WINDOWPOSCHANGING = &H46
     WM_WINDOWPOSCHANGED = &H47
     WM_POWER = &H48
     WM_COPYDATA = &H4A
     WM_CANCELJOURNAL = &H4B
     WM_NCCREATE = &H81
     WM_NCDESTROY = &H82
     WM_NCCALCSIZE = &H83
     WM_NCHITTEST = &H84
     WM_NCPAINT = &H85
     WM_NCACTIVATE = &H86
     WM_GETDLGCODE = &H87
     WM_NCMOUSEMOVE = &HA0
     WM_NCLBUTTONDOWN = &HA1
     WM_NCLBUTTONUP = &HA2
     WM_NCLBUTTONDBLCLK = &HA3
     WM_NCRBUTTONDOWN = &HA4
     WM_NCRBUTTONUP = &HA5
     WM_NCRBUTTONDBLCLK = &HA6
     WM_NCMBUTTONDOWN = &HA7
     WM_NCMBUTTONUP = &HA8
     WM_NCMBUTTONDBLCLK = &HA9
     WM_KEYFIRST = &H100
     WM_KEYDOWN = &H100
     WM_KEYUP = &H101
     WM_CHAR = &H102
     WM_DEADCHAR = &H103
     WM_SYSKEYDOWN = &H104
     WM_SYSKEYUP = &H105
     WM_SYSCHAR = &H106
     WM_SYSDEADCHAR = &H107
     WM_KEYLAST = &H108
     WM_INITDIALOG = &H110
     WM_COMMAND = &H111
     WM_SYSCOMMAND = &H112
     WM_TIMER = &H113
     WM_HSCROLL = &H114
     WM_VSCROLL = &H115
     WM_INITMENU = &H116
     WM_INITMENUPOPUP = &H117
     WM_MENUSELECT = &H11F
     WM_MENUCHAR = &H120
     WM_ENTERIDLE = &H121
     WM_CTLCOLORMSGBOX = &H132
     WM_CTLCOLOREDIT = &H133
     WM_CTLCOLORLISTBOX = &H134
     WM_CTLCOLORBTN = &H135
     WM_CTLCOLORDLG = &H136
     WM_CTLCOLORSCROLLBAR = &H137
     WM_CTLCOLORSTATIC = &H138
     WM_MOUSEFIRST = &H200
     WM_MOUSEMOVE = &H200
     WM_LBUTTONDOWN = &H201
     WM_LBUTTONUP = &H202
     WM_LBUTTONDBLCLK = &H203
     WM_RBUTTONDOWN = &H204
     WM_RBUTTONUP = &H205
     WM_RBUTTONDBLCLK = &H206
     WM_MBUTTONDOWN = &H207
     WM_MBUTTONUP = &H208
     WM_MBUTTONDBLCLK = &H209
     WM_MOUSELAST = &H209
     WM_PARENTNOTIFY = &H210
     WM_ENTERMENULOOP = &H211
     WM_EXITMENULOOP = &H212
     WM_CAPTURECHANGED = &H215
     WM_MDICREATE = &H220
     WM_MDIDESTROY = &H221
     WM_MDIACTIVATE = &H222
     WM_MDIRESTORE = &H223
     WM_MDINEXT = &H224
     WM_MDIMAXIMIZE = &H225
     WM_MDITILE = &H226
     WM_MDICASCADE = &H227
     WM_MDIICONARRANGE = &H228
     WM_MDIGETACTIVE = &H229
     WM_MDISETMENU = &H230
     WM_DROPFILES = &H233
     WM_MDIREFRESHMENU = &H234
     WM_CUT = &H300
     WM_COPY = &H301
     WM_PASTE = &H302
     WM_CLEAR = &H303
     WM_UNDO = &H304
     WM_RENDERFORMAT = &H305
     WM_RENDERALLFORMATS = &H306
     WM_DESTROYCLIPBOARD = &H307
     WM_DRAWCLIPBOARD = &H308
     WM_PAINTCLIPBOARD = &H309
     WM_VSCROLLCLIPBOARD = &H30A
     WM_SIZECLIPBOARD = &H30B
     WM_ASKCBFORMATNAME = &H30C
     WM_CHANGECBCHAIN = &H30D
     WM_HSCROLLCLIPBOARD = &H30E
     WM_QUERYNEWPALETTE = &H30F
     WM_PALETTEISCHANGING = &H310
     WM_PALETTECHANGED = &H311
     WM_HOTKEY = &H312
     WM_PENWINFIRST = &H380
     WM_PENWINLAST = &H38F
End Enum

Private Const GWL_WNDPROC = (-4)
Public OldWindowProc As Long
Public ctlActiveColorbox As ctlColorBox

Public Function HookFunc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo EH

If Not ctlActiveColorbox Is Nothing Then
    Select Case Msg
    Case WM_CAPTURECHANGED
        ctlActiveColorbox.LosingCapture
        HookFunc = 0
        Exit Function
    
    Case WM_MOUSEMOVE
        ctlActiveColorbox.MouseMoved LOWORD(lParam), HIWORD(lParam)
    
    Case WM_KEYDOWN
        ctlActiveColorbox.Cancel
        'HookFunc = 0
        'Exit Function
    End Select
End If

HookFunc = CallWindowProc(OldWindowProc, hWnd, Msg, wParam, lParam)
Exit Function

EH:
UnhookWindow hWnd
End Function

Public Sub HookWindow(hWnd As Long)
On Local Error GoTo EH
If OldWindowProc Then Exit Sub
OldWindowProc = GetWindowLong(hWnd, GWL_WNDPROC)
If OldWindowProc <> 0 Then SetWindowLong hWnd, GWL_WNDPROC, AddressOf HookFunc
EH:
End Sub

Public Sub UnhookWindow(hWnd As Long)
On Local Error GoTo EH
If (OldWindowProc <> 0) Then SetWindowLong hWnd, GWL_WNDPROC, OldWindowProc
OldWindowProc = 0
EH:
End Sub
