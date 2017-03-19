Attribute VB_Name = "MainModule"
Option Explicit

#Const conUseGLAbout = 0
#Const conBlackDefaultInterface = 0
#Const conProtected = 0
'#Const conTips = False
'#Const conDebug = False
Public Const conLanguage = 0 'Languages.langEnglish
Public Const EMailRSA = "rakov_s@ukr.net"
Public Const EMailOK = "dg@osenkov.com"
Public Const EMailCommon = "dg@osenkov.com"

Public CanSave As Boolean
Public MenuHandle As Long
Public Const PerversionBase = 2147483611
Public Const RegTxt = "dg_regk.txt"

'==========================================================
'API declarations:
'==========================================================

'==========================================================
'Constants
'==========================================================

Public Const CC_FULLOPEN = &H2
Public Const BITSPIXEL = 12
Public Const CC_RGBINIT = &H1
Public Const FW_NORMAL = 400
Public Const FW_BOLD = 700
Public Const FW_MEDIUM = 500

Public Const HELP_COMMAND = &H102&
Public Const HELP_CONTENTS = &H3&
Public Const HELP_CONTEXT = &H1
Public Const HELP_CONTEXTPOPUP = &H8&
Public Const HELP_FORCEFILE = &H9&
Public Const HELP_HELPONHELP = &H4
Public Const HELP_INDEX = &H3
Public Const HELP_KEY = &H101
Public Const HELP_MULTIKEY = &H201&
Public Const HELP_PARTIALKEY = &H105&
Public Const HELP_QUIT = &H2
Public Const HELP_SETCONTENTS = &H5&
Public Const HELP_SETINDEX = &H5
Public Const HELP_SETWINPOS = &H203&

Public Const NULL_BRUSH = 5
Public Const NULL_PEN = 8
Public Const BF_BOTTOM = &H8
Public Const BF_FLAT = &H4000
Public Const BF_LEFT = &H1
Public Const BF_MONO = &H8000
Public Const BF_RIGHT = &H4
Public Const BF_SOFT = &H1000
Public Const BF_TOP = &H2
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_RAISEDOUTER = &H1
Public Const PS_NULL = 5
Public Const PS_SOLID = 0
Public Const DT_NOPREFIX = &H800
Public Const DT_NOCLIP = &H100
Public Const DT_EXPANDTABS = &H40

Public Const WS_VISIBLE = &H10000000
Public Const WS_THICKFRAME = &H40000
Public Const WS_MAXIMIZE = &H1000000
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOOWNERZORDER = &H200
Public Const GMEM_MOVEABLE = &H2
Public Const GMEM_ZEROINIT = &H40
Public Const BDR_SUNKENINNER = &H8

Public Const DSS_DISABLED = &H20
Public Const DSS_MONO = &H80
Public Const DSS_NORMAL = &H0
Public Const DSS_RIGHT = &H8000&
Public Const DSS_UNION = &H10
Public Const DST_BITMAP = &H4
Public Const DST_COMPLEX = &H0
Public Const DST_ICON = &H3
Public Const DST_PREFIXTEXT = &H2
Public Const DST_TEXT = &H1

Public Const BDR_SUNKENOUTER = &H2
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const TA_BOTTOM = 8
Public Const TA_LEFT = 0
Public Const TA_TOP = 0
Public Const TA_RIGHT = 2
Public Const TA_CENTER = 6
Public Const TA_BOTTOMCENTER = TA_BOTTOM Or TA_CENTER
Public Const TA_LEFTTOP = TA_LEFT Or TA_TOP
Public Const TA_TOPRIGHT = TA_TOP Or TA_RIGHT
Public Const VTA_BOTTOM = TA_RIGHT
Public Const VTA_CENTER = TA_CENTER
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_NORMAL = 1
Public Const TA_TOPCENTER = TA_TOP Or TA_CENTER
Public Const SND_ALIAS = &H10000
Public Const SND_ALIAS_ID = &H110000
Public Const SND_ALIAS_START = 0
Public Const SND_APPLICATION = &H80
Public Const SND_ASYNC = &H1
Public Const OPAQUE = 2
Public Const SND_FILENAME = &H20000
Public Const DIB_RGB_COLORS = 0
Public Const BI_RGB = 0&
Public Const SWP_FRAMECHANGED = &H20
Public Const SND_LOOP = &H8
Public Const SND_MEMORY = &H4
Public Const SND_NODEFAULT = &H2
Public Const SHOW_OPENNOACTIVATE = 4
Public Const SND_NOSTOP = &H10
Public Const SND_NOWAIT = &H2000
Public Const ALTERNATE = 1
Public Const HWND_NOTOPMOST = -2
Public Const ERROR_NO_MORE_FILES = 18&
Public Const SND_PURGE = &H40
Public Const SND_RESERVED = &HFF000000
Public Const MM_ANISOTROPIC = 8
Public Const HKL_NEXT = 1
Public Const HKL_PREV = 0
Public Const KL_NAMELENGTH = 9
Public Const Transparent = 1
Public Const SND_RESOURCE = &H40004
Public Const SND_SYNC = &H0
Public Const SND_TYPE_MASK = &H170007
Public Const SND_VALID = &H1F
Public Const SND_VALIDFLAGS = &H17201F
Public Const Max_Path = 260
Public Const DT_EDITCONTROL = &H2000
Public Const DT_BOTTOM = &H8
Public Const DT_CALCRECT = &H400
Public Const DT_CENTER = &H1
Public Const DT_LEFT = &H0
Public Const DT_RIGHT = &H2
Public Const DT_SINGLELINE = &H20
Public Const DT_TOP = &H0
Public Const DT_VCENTER = &H4
Public Const DT_WORDBREAK = &H10
Public Const SM_CXSCREEN = 0
Public Const SM_CYSCREEN = 1
Public Const SM_CXHSCROLL = 21
Public Const SM_CXVSCROLL = 2
Public Const SM_CYHSCROLL = 3
Public Const SM_CYVSCROLL = 20
Public Const SM_CYVTHUMB = 9
Public Const DEFAULT_QUALITY = 0
Public Const DEFAULT_PITCH = 0
Public Const ANSI_CHARSET = 0
Public Const OUT_DEFAULT_PRECIS = 0
Public Const CLIP_DEFAULT_PRECIS = 0
Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)
Public Const GWL_HINSTANCE = (-6)
Public Const HWND_TOPMOST = -1
Public Const WS_EX_TOPMOST = &H8&
Public Const LOCALE_ILANGUAGE = &H1
Public Const LOCALE_SLANGUAGE = &H2
Public Const LOCALE_SENGLANGUAGE = &H1001
Public Const LOCALE_SABBREVLANGNAME = &H3
Public Const LOCALE_SNATIVELANGNAME = &H4
Public Const LR_SHARED = &H8000&
Public Const MF_BYPOSITION = &H400&
Public Const DI_MASK = 1
Public Const DI_IMAGE = 2
Public Const DI_NORMAL = 3
Public Const DI_COMPAT = 4
Public Const DI_DEFAULTSIZE = 8
Public Const IMAGE_BITMAP = 0
Public Const IMAGE_ICON = 1
Public Const IMAGE_CURSOR = 2

'==========================================================
'Types
'==========================================================

Public Type Size
        cx As Long
        cy As Long
End Type

Public Type POINTAPI
        X As Long
        Y As Long
End Type

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName(0 To 31) As Byte
        'lfFaceName As String * 32
End Type

Public Type TEXTMETRIC
        tmHeight As Long
        tmAscent As Long
        tmDescent As Long
        tmInternalLeading As Long
        tmExternalLeading As Long
        tmAveCharWidth As Long
        tmMaxCharWidth As Long
        tmWeight As Long
        tmOverhang As Long
        tmDigitizedAspectX As Long
        tmDigitizedAspectY As Long
        tmFirstChar As Byte
        tmLastChar As Byte
        tmDefaultChar As Byte
        tmBreakChar As Byte
        tmItalic As Byte
        tmUnderlined As Byte
        tmStruckOut As Byte
        tmPitchAndFamily As Byte
        tmCharSet As Byte
End Type

Private Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

Private Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Public Type METAFILEPICT
        mm As Long
        xExt As Long
        yExt As Long
        hMF As Long
End Type

Private Type TRIRGB
    Red As Byte
    Green As Byte
    Blue As Byte
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * Max_Path
        cAlternate As String * 14
End Type

Public Type MMTIME
        wType As Long
        u As Long
End Type

Public Type TIMECAPS
        wPeriodMin As Long
        wPeriodMax As Long
End Type

Public Type BROWSEINFO
     hWndOwner As Long
     pidlRoot As Long
     pszDisplayName As Long
     lpszTitle As String
     ulFlags As Long
     lpfnCallback As Long
     lParam As Long
     iImage As Long
End Type

Public Type RECTS
        Left As Integer
        Top As Integer
        Right As Integer
        Bottom As Integer
End Type

Public Type METAFILEHEADER
    Key As Long
    hMF As Integer
    bbox As RECTS
    inch As Integer
    reserved As Long
    checksum As Integer
End Type

Public Type LARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type
'        LargeNum = TotalSpace.LowPart
'        If LargeNum < 0 Then LargeNum = LargeNum + 2 ^ 32
'        GetDriveSpace = TotalSpace.HighPart * 2 ^ 32 + LargeNum


'==========================================================
'Function declarations
'==========================================================

Public Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Public Declare Function SetMenu Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long) As Long
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Public Declare Function GetWinMetaFileBitsLen Lib "gdi32" Alias "GetWinMetaFileBits" (ByVal hemf As Long, ByVal cbBuffer As Long, ByVal lpbBuffer As Long, ByVal fnMapMode As Long, ByVal hdcRef As Long)

Public Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long

Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWinMetaFileBits Lib "gdi32" (ByVal cbBuffer As Long, lpbBuffer As Byte, ByVal hdcRef As Long, lpmfp As METAFILEPICT) As Long
Public Declare Function GetWinMetaFileBits Lib "gdi32" (ByVal hemf As Long, ByVal cbBuffer As Long, lpbBuffer As Byte, ByVal fnMapMode As Long, ByVal hdcRef As Long) As Long
Public Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
Public Declare Function ActivateKeyboardLayout Lib "user32" (ByVal HKL As Long, ByVal Flags As Long) As Long
Public Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Public Declare Function GetKeyboardLayoutName Lib "user32" Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long

Public Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Public Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal N3 As Long, ByVal n4 As Long, ByVal un As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long

Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As Long
Public Declare Function Polyline Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function PolyBezier Lib "gdi32" (ByVal hDC As Long, lppt As POINTAPI, ByVal cPoints As Long) As Long
Public Declare Function PolyBezierTo Lib "gdi32" (ByVal hDC As Long, lppt As POINTAPI, ByVal cCount As Long) As Long
Public Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function PolyDraw Lib "gdi32" (ByVal hDC As Long, lppt As POINTAPI, lpbTypes As Byte, ByVal cCount As Long) As Long
Public Declare Function PolyPolygon Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long) As Long
Public Declare Function PolylineTo Lib "gdi32" (ByVal hDC As Long, lppt As POINTAPI, ByVal cCount As Long) As Long
Public Declare Function Arc Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Public Declare Function Pie Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Public Declare Function PolyPolyline Lib "gdi32" (ByVal hDC As Long, lppt As POINTAPI, lpdwPolyPoints As Long, ByVal cCount As Long) As Long
Public Declare Function GetCurrentPositionEx Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI) As Long
Public Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Public Declare Function SetROP2 Lib "gdi32" (ByVal hDC As Long, ByVal nDrawMode As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Public Declare Function GetNearestColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long

Public Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function PlayMetaFile Lib "gdi32" (ByVal hDC As Long, ByVal hMF As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function EnumFonts Lib "gdi32" Alias "EnumFontsA" (ByVal hDC As Long, ByVal lpsz As String, ByVal lpFontEnumProc As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowExtEx Lib "gdi32" (ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, lpSize As Size) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetMetaFileBitsEx Lib "gdi32" (ByVal nSize As Long, lpData As Byte) As Long
Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Public Declare Function CloseEnhMetaFile Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function DeleteEnhMetaFile Lib "gdi32" (ByVal hemf As Long) As Long
Public Declare Function DeleteMetaFile Lib "gdi32" (ByVal hMF As Long) As Long
Public Declare Function CloseMetaFile Lib "gdi32" (ByVal hMF As Long) As Long
Public Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function CreateEnhMetaFile Lib "gdi32" Alias "CreateEnhMetaFileA" (ByVal hdcRef As Long, ByVal lpFileName As String, lpRect As RECT, ByVal lpDescription As String) As Long
Public Declare Function CreateMetaFile Lib "gdi32" Alias "CreateMetaFileA" (ByVal lpString As String) As Long
Public Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Declare Function GetMetaFileBitsEx Lib "gdi32" (ByVal hMF As Long, ByVal nSize As Long, lpvData As Any) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function SetMapMode Lib "gdi32" (ByVal hDC As Long, ByVal nMapMode As Long) As Long
Public Declare Function SetTextAlign Lib "gdi32" (ByVal hDC As Long, ByVal wFlags As Long) As Long
Public Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Public Declare Function GetLogicalDrives Lib "kernel32" () As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Public Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Char As Byte)
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public Declare Function LoadCursorBynum& Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long)
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function LoadImageBynum Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long

Declare Function lopen Lib "kernel32" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Long) As Long
Declare Function lclose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long
Declare Function lcreat Lib "kernel32" Alias "_lcreat" (ByVal lpPathName As String, ByVal iAttribute As Long) As Long
Declare Function llseek Lib "kernel32" Alias "_llseek" (ByVal hFile As Long, ByVal lOffset As Long, ByVal iOrigin As Long) As Long
Declare Function lread Lib "kernel32" Alias "_lread" (ByVal hFile As Long, lpBuffer As Any, ByVal wBytes As Long) As Long
Declare Function lwrite Lib "kernel32" Alias "_lwrite" (ByVal hFile As Long, lpBuffer As Any, ByVal wBytes As Long) As Long

Declare Function hread Lib "kernel32" Alias "_hread" (ByVal hFile As Long, lpBuffer As Any, ByVal lBytes As Long) As Long
Declare Function hwrite Lib "kernel32" Alias "_hwrite" (ByVal hFile As Long, lpBuffer As Any, ByVal lBytes As Long) As Long
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Public Declare Function SHBrowseForFolder Lib "Shell32" (lpBI As BROWSEINFO) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function SHGetPathFromIDList Lib "Shell32" (ByVal PidList As Long, ByVal lpBuffer As String) As Long


Public Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef PC As LARGE_INTEGER) As Long
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef PC As LARGE_INTEGER) As Long
Public Declare Function timeGetTime Lib "WinMM.dll" () As Long
Public Declare Function timeKillEvent Lib "WinMM.dll" (ByVal uID As Long) As Long
Public Declare Function timeSetEvent Lib "WinMM.dll" (ByVal uDelay As Long, ByVal uResolution As Long, ByVal lpFunction As Long, ByVal dwUser As Long, ByVal uFlags As Long) As Long
Public Declare Function timeGetSystemTime Lib "WinMM.dll" (lpTime As MMTIME, ByVal uSize As Long) As Long
Public Declare Function timeGetDevCaps Lib "WinMM.dll" (lpTimeCaps As TIMECAPS, ByVal uSize As Long) As Long
Public Declare Function timeEndPeriod Lib "WinMM.dll" (ByVal uPeriod As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function timeBeginPeriod Lib "WinMM.dll" (ByVal uPeriod As Long) As Long
Public Declare Function PlaySound Lib "WinMM.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
' Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpfilename As String) As Long
Public Declare Function GetPrivateProfileStringByKeyName Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName$, ByVal lpszKey$, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$) As Long
Public Declare Function GetPrivateProfileStringKeys Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName$, ByVal lpszKey&, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$) As Long
Public Declare Function GetPrivateProfileStringSections Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName&, ByVal lpszKey&, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$) As Long
' Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, lpString As Any, ByVal lplFileName As String) As Long
Public Declare Function WritePrivateProfileStringByKeyName Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lplFileName As String) As Long
Public Declare Function WritePrivateProfileStringToDeleteKey Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Long, ByVal lplFileName As String) As Long
Public Declare Function WritePrivateProfileStringToDeleteSection Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Long, ByVal lpString As Long, ByVal lplFileName As String) As Long

'=========================================================
' /API declarations finished
'=========================================================



'=========================================================
' DG enumerations
'=========================================================

Public Enum CursorState
    curStateArrow
    curStateCross
    curStateDrag
    curStateNo
    curStateQuestion
    curStateHourglass
    curStateAdd
    curStateRemove
    curStateSelect
End Enum

Public Enum Languages
    langEnglish
    langRussian
    langGerman = 2000
    langUkrainian
End Enum

Public Enum MenuBarState
    mbsToolBar
    mbsSelectObjectsFinish
    mbsMacroGivens
    mbsMacroResults
    mbsMacroRun
    mbsAnimation
    mbsDemo
    mbsCancel
End Enum

Public Enum DynamicLabelType
    StaticString
    DynamicString
End Enum

Public Enum TextBorderStyle
    tbsNone
    tbs2D
    tbs3DThin
    tbs3DRaised
    tbs3DSunken
    tbs3DEtched
    tbs3DBump
    tbs3DShadow
    tbs3DTransparentShadow
    tbs3DGlass
End Enum

Public Enum MacroErrors
    meResultsNotSelected
    meErrorCreatingMacro
    meNotAllGivensSelected
    meAtLeastOneFigureNeeded
End Enum

Public Enum ObjectSelectionType
    ostShowHideObjects
    ostCalcPoints
    ostMacroGivens
    ostMacroResults
End Enum

Public Enum ObjectSelectionCaller
    oscButton
    oscCalcLabels
    oscCalcAnPoint
    oscCalcWE
    oscCalcFunction
    oscCalculator
End Enum

Public Enum StaticGraphicType
    sgPolygon
    sgBezier
    sgVector
End Enum

Public Enum ButtonType
    butShowHide
    butMsgBox
    butPlaySound
    butLaunchFile
End Enum

Public Enum GeometryObjectType
    gotGeneric
    gotPoint
    gotFigure
    gotSG
    gotLocus
    gotLabel
    gotWE
    gotButton
End Enum

Public Enum DrawState
    dsSelect = -1
    dsPoint
    dsSegment
    dsRay
    dsLine_2Points
    dsLine_PointAndParallelLine
    dsLine_PointAndPerpendicularLine
    dsCircle_CenterAndCircumPoint
    dsCircle_CenterAndTwoPoints
    dsCircle_ArcCenterAndRadiusAndTwoPoints
    dsMiddlePoint
    dsSimmPoint
    dsSimmPointByLine
    dsInvert
    dsIntersect
    dsPointOnFigure
    dsMeasureDistance
    dsMeasureAngle
    dsAnLineGeneral
    dsAnLineCanonic
    dsAnLineNormal
    dsAnLineNormalPoint
    dsAnCircle
    dsAnPoint
    dsDynamicLocus
    dsBisector
    dsMeasureArea
    dsPolygon
End Enum

Public Type Transposition
    Element1() As Variant
    Element2() As Variant
    Count As Long
End Type

'=========================================================
'DG global constants
'=========================================================

Public Const AppName = "DGeometry"
Public Const extFIG = "DGF"
Public Const extMAC = "DGM"
Public Const extBMP = "BMP"
Public Const extWMF = "WMF"
Public Const extEMF = "EMF"
Public Const extHTM = "HTM"
Public Const extWAV = "WAV"
Public Const HelpFileRussian = "dg_ru.hlp"
Public Const HelpFileEnglish = "dg_en.hlp"
Public Const HelpFileGerman = "dg_de.hlp"
Public Const HelpFileUkrainian = "dg_uk.hlp"

#If conBlackDefaultInterface = 1 Then
    Public Const DefaultAlign = vbAlignLeft
#Else
    Public Const DefaultAlign = vbAlignTop
#End If

Public Const IconSize16 = 200
Public Const IconSize32 = 300

Public Const MaxFigureCount = 4000
Public Const MaxPointCount = 4000
Public Const MaxLabelCount = 200
Public Const MaxButtonCount = 200
Public Const MaxCoord As Double = 1000

Public Const MaxPrecision = 14
Public Const defAngleMarkRadius = 16
Public Const defAngleMarkDist = 25
Public Const FrameWidth = 8, FrameHeight = 4
Public Const NumDecimalDigits = 6
Public Const StatusBarAutoRedraw = False
Public Const DynamicLocusPointCount = 64
Public Const HighQualityDynamicLocusPointCount = 640
Public Const MaxDynamicLocusDetails = 511
Public Const MaxDynamicLocusDetailsHigh = 4095
Public Const IconSize = 16
Public Const MenuHorizSpace = 10
Public Const MenuVerticSpace = 4
Public Const Margin = 4
Public Const TTTMargin = 2
Public Const TTOffset = 16
Public Const Shadow = 6
Public Const IconIndent = 4
Public Const ArrowLength = 8
Public Const RulerSize = 10
Public Const RulerFontSize = 7
Public Const RulerFontName = "Arial"
Public Const FullScreenLineMargin = 0
Public Const EmptyVar = -2 ^ 31 + 3
Public Const SysColorTranslationBase = 2 ^ 31
Public Const MRUMax = 6

Public Const MaxDrawWidth = 16
Public Const MinPointSize = 2
Public Const MaxPointSize = 30

Public Const MenuCheckDrawMode = DrawModeConstants.vbCopyPen
Public Const SelectionThickness = 8

'=========================================================
'Color constants
'=========================================================
#If conBlackDefaultInterface = 1 Then
    Public Const colPaperGradient1 = &H404040
    Public Const colPaperGradient2 = &H808080
    Public Const colAxes = &H707070
    Public Const colGridColor = &H505050
    
    Public Const colFigureForeColor = vbCyan
    Public Const colPointColor = vbWhite
    Public Const colPointFillColor = vbBlack
    Public Const colSemiDependentForeColor = vbYellow
    Public Const colDependent = vbCyan
    Public Const colLocusColor = vbGreen
    Public Const colTextColor = vbWhite
    Public Const colPlotColor = vbGreen
    Public Const colPolygonFillColor = &H808080
    Public Const colBezierColor = vbCyan
    Public Const colVectorColor = colFigureForeColor
    
    'Public Const colFormBackGround = &HC0C0C0 'vbApplicationWorkspace
    Public Const colFormBackGround = vbApplicationWorkspace
    Public Const colRulerBackColor = &HF0E000
    Public Const colRulerForeColor = &H0
    Public Const colRulerGradient = vbWhite
    Public Const colStatusGradient = &HA0A090 ' &HD0A0A0 '
    
    Public Const colMenuBackGround = SystemColorConstants.vbMenuBar '&H808080
    Public Const colMenuForeGround = vbCyan
    Public Const colMenuGradient = colStatusGradient
    Public Const colMenuCheck = vbWhite '&HE0E0E0
    Public Const colInActiveTab = &H808080
    Public Const colActiveTab = vbWhite
    Public Const colWatchExpressions = colStatusGradient
#Else
    Public Const colPaperGradient1 = vbWhite
    Public Const colPaperGradient2 = &HFFE0E0
    Public Const colAxes = &HF0A0A0
    Public Const colGridColor = &HF0F0F0
    
    Public Const colFigureForeColor = vbBlack
    Public Const colFigureFillColor = &HE0FFFF
    Public Const colPointColor = &H400000
    Public Const colPointFillColor = vbWhite
    Public Const colSemiDependentForeColor = vbRed
    Public Const colFigurePointFillColor = vbYellow
    Public Const colDependent = vbBlue
    Public Const colDependentPointFillColor = &HFFEEEE
    
    Public Const colLocusColor = vbBlue
    Public Const colTextColor = vbBlack
    Public Const colPlotColor = vbBlue
    Public Const colPolygonFillColor = &HFFD0D0
    Public Const colBezierColor = vbBlue
    Public Const colVectorColor = colFigureForeColor
    
    Public Const colFormBackGround = vbApplicationWorkspace
    Public Const colRulerBackColor = &H10E0FF
    Public Const colRulerForeColor = &H0
    Public Const colRulerGradient = vbWhite
    Public Const colStatusGradient = 12640496 '&HD0D0D0 '&HA0A090 ' &HD0A0A0 '
    
    Public Const colMenuBackGround = SystemColorConstants.vbMenuBar '&H808080
    Public Const colMenuForeGround = vbCyan
    Public Const colMenuGradient = SystemColorConstants.vbButtonFace
    Public Const colMenuCheck = vbWhite '&HE0E0E0
    Public Const colInActiveTab = &H808080
    Public Const colActiveTab = vbWhite
    Public Const colWatchExpressions = colStatusGradient
#End If

'=========================================================
'Default value constants
'=========================================================
Public Const defSLabelFontName = "Arial"
Public Const defSLabelFontSize = 10
Public Const defPointSize = 7
Public Const defPointDrawWidth = 1
Public Const defPointShape = vbShapeCircle
Public Const defFigurePointShape = vbShapeCircle
Public Const defDependentPointShape = vbShapeCircle

Public Const defFigureDrawWidth = 1
Public Const defFigureDrawMode = 13
Public Const defFigureDrawStyle = 0
Public Const defInternalScaleMode = ScaleModeConstants.vbCentimeters
Public Const defDemoInterval = 1000
Public Const defLabelFontName = "Arial"
Public Const defLabelFontSize = 10
Public Const defLabelFontBold As Boolean = False
Public Const defLabelFontitalic As Boolean = False
Public Const defLabelFontUnderline As Boolean = False


'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'=========================================================
'Module level complete.
'=========================================================

Public Sub Main()
FormMain.Show
End Sub

Public Function RetrieveName(ByVal FName As String) As String
If InStr(FName, "\") = 0 Then RetrieveName = FName: Exit Function
RetrieveName = Right(FName, Len(FName) - InStrRev(FName, "\"))
End Function

Public Function RetrieveDir(ByVal FName As String) As String
Dim Z As Long
On Local Error Resume Next
If FName = "" Then Exit Function
Z = InStrRev(FName, "\")
If Z = 0 Then Exit Function
RetrieveDir = AddDirSep(Left(FName, Z - 1))
End Function

'Public Function VBGetPrivateProfileString(ByVal Section As String, ByVal Key As String, ByVal File As String, Optional ByVal DefaultStr As String = "") As String
'Dim Characters As Long, KeyValue As String
'
'Characters = GetPrivateProfileStringByKeyName(Section, Key, DefaultStr, TempShortStringBuffer, LenTempShortStringBuffer, File)
'If Characters >= LenTempShortStringBuffer - 1 Then
'    Characters = GetPrivateProfileStringByKeyName(Section, Key, DefaultStr, TempLongStringBuffer, LenTempLongStringBuffer, File)
'    KeyValue = Left(TempLongStringBuffer, Characters)
'Else
'    KeyValue = Left(TempShortStringBuffer, Characters)
'End If
'VBGetPrivateProfileString = KeyValue
'End Function

'
'Public Function GetTextSize(ByVal lpString As String, Optional ByVal tFontName As String = "", Optional ByVal tFontSize As Long = 0, Optional ByVal tFontBold As Boolean = False, Optional ByVal tFontItalic As Boolean = False, Optional ByVal tFontUnderline As Boolean = False, Optional ByVal tFontStrikeThru As Boolean = False, Optional ByVal Ang As Single = 0) As Size
'Dim LF As LOGFONT, I As Long, NF As Long, tS As String, lpRect As RECT
'LF.lfEscapement = CLng(Ang * 10)
'LF.lfOrientation = LF.lfEscapement
'LF.lfWeight = IIf(tFontBold, 700, 400)
'LF.lfItalic = IIf(tFontItalic, 255, 0)
'LF.lfUnderline = IIf(tFontUnderline, 255, 0)
'LF.lfStrikeOut = IIf(tFontStrikeThru, 255, 0)
'LF.lfQuality = 2
'''LF.lfWidth = 0'LF.lfCharSet = 0'LF.lfOutPrecision = 0'LF.lfClipPrecision = 0'LF.lfPitchAndFamily = 0
'tS = tFontName
'If Len(tS) > 31 Then tS = "Arial"
'For Q = 1 To Len(tS): LF.lfFaceName(Q) = Asc(Mid$(tS, Q, 1)): Next
'LF.lfFaceName(Len(tS) + 1) = 0
'LF.lfHeight = tFontSize * -20 / Screen.TwipsPerPixelY
'I = CreateFontIndirect(LF)
'NF = SelectObject(ServiceHDC, I)
'lpRect.Left = 0
'lpRect.Top = 0
'SetTextAlign Paper.hDC, TA_TOP Or TA_LEFT
'DrawText ServiceHDC, lpString, Len(lpString), lpRect, DT_CALCRECT Or DT_NOPREFIX Or DT_INTERNAL
'GetTextSize.CX = lpRect.Right - lpRect.Left
'GetTextSize.CY = lpRect.Bottom - lpRect.Top
'SelectObject ServiceHDC, NF
'DeleteObject I
'End Function

'Public Function GetTextSizeMultiLine(ByVal lpString As String, Optional ByVal tFontName As String = "", Optional ByVal tFontSize As Long = 0, Optional ByVal tFontBold As Boolean = False, Optional ByVal tFontItalic As Boolean = False, Optional ByVal tFontUnderline As Boolean = False, Optional ByVal tFontStrikeThru As Boolean = False, Optional ByVal Ang As Single = 0) As Size
'Dim LF As LOGFONT, I As Long, NF As Long, TS As String, lpRect As RECT
'LF.lfEscapement = CLng(Ang * 10)
'LF.lfOrientation = LF.lfEscapement
'LF.lfWeight = IIf(tFontBold, 700, 400)
'LF.lfItalic = IIf(tFontItalic, 255, 0)
'LF.lfUnderline = IIf(tFontUnderline, 255, 0)
'LF.lfStrikeOut = IIf(tFontStrikeThru, 255, 0)
'LF.lfQuality = 2
'''LF.lfWidth = 0'LF.lfCharSet = 0'LF.lfOutPrecision = 0'LF.lfClipPrecision = 0'LF.lfPitchAndFamily = 0
'TS = tFontName
'If Len(TS) > 31 Then TS = "Arial"
'For Q = 1 To Len(TS): LF.lfFaceName(Q) = Asc(Mid$(TS, Q, 1)): Next
'LF.lfFaceName(Len(TS) + 1) = 0
'LF.lfHeight = tFontSize * -20 / Screen.TwipsPerPixelY
'I = CreateFontIndirect(LF)
'NF = SelectObject(ServiceHDC, I)
'lpRect.Left = 0
'lpRect.Top = 0
'DrawText ServiceHDC, lpString, Len(lpString), lpRect, DT_CALCRECT Or DT_NOPREFIX
'GetTextSize.CX = lpRect.Right - lpRect.Left
'GetTextSize.CY = lpRect.Bottom - lpRect.Top
'SelectObject ServiceHDC, NF
'DeleteObject I
'End Function

Public Sub CreateTempFontDC()
'TempFontDC = CreateCompatibleDC(Paper.hDC)
DefaultFontCharset = Paper.Font.Charset
ServiceHDC = CreateCompatibleDC(Paper.hDC)
ServiceHBitmap = CreateCompatibleBitmap(Paper.hDC, GetSystemMetrics(SM_CXSCREEN), GetSystemMetrics(SM_CYSCREEN))
OldServiceHBitmap = SelectObject(ServiceHDC, ServiceHBitmap)
'Dim LF As LOGFONT, I As Long, tS As String

'LF.lfWeight = 400
'LF.lfItalic = 0
'LF.lfUnderline = 0
'LF.lfStrikeOut = 0
'LF.lfQuality = 2
'tS = defSLabelFontName
'If Len(tS) > 31 Then tS = "Arial"
'For Q = 1 To Len(tS): LF.lfFaceName(Q) = Asc(Mid$(tS, Q, 1)): Next
'LF.lfFaceName(Len(tS) + 1) = 0
'LF.lfHeight = defSLabelFontSize * -20 / Screen.TwipsPerPixelY

'I = CreateFontIndirect(LF)
'TempFontHandle = SelectObject(TempFontDC, I)
End Sub

Public Sub DeleteTempFontDC()
SelectObject ServiceHDC, OldServiceHBitmap
DeleteObject ServiceHBitmap
DeleteDC ServiceHDC
'DeleteEnhMetaFile BackgroundMetafile
End Sub

'Public Sub CreateBackgroundMetafile()
'Dim hDC As Long, lpRect As RECT, UnitMM As Double
'
'If Not nShowAxes And Not nShowGrid Then Exit Sub
'
'UnitMM = Paper.ScaleX(1, vbPixels, vbMillimeters) * 100
'lpRect.Right = PaperScaleWidth * UnitMM
'lpRect.Bottom = PaperScaleHeight * UnitMM
'
'If BackgroundMetafile <> 0 Then DeleteEnhMetaFile BackgroundMetafile
'hDC = CreateEnhMetaFile(GetWindowDC(Paper.hWnd), vbNullString, lpRect, GetString(ResCaption) & vbNullChar & RetrieveName(DrawingName) & vbNullChar & vbNullChar)
'SetBkMode hDC, Transparent
'If nShowGrid Then ShowGrid hDC
'If nShowAxes Then ShowAxes hDC
'BackgroundMetafile = CloseEnhMetaFile(hDC)
'End Sub

'Public Function GetTempFontSize(ByVal sStr As String) As Size
'Dim lpSize As Size
'sStr = GetTextExtentPoint32(TempFontDC, sStr, Len(sStr), lpSize)
'GetTempFontSize = lpSize
'End Function

Public Sub InitGraphics()
Dim Q As Long

Paper.DrawMode = 13
Paper.DrawStyle = 0
Paper.DrawWidth = 1
Paper.FillStyle = 1
SetBkMode Paper.hDC, Transparent
SetTextAlign Paper.hDC, TA_BOTTOM Or TA_CENTER

Paper.FontName = setdefPointFontName
Paper.FontSize = setdefPointFontSize
Paper.FontBold = setdefPointFontBold
Paper.FontItalic = setdefPointFontItalic
Paper.FontUnderline = setdefPointFontUnderline
Paper.Font.Charset = setdefPointFontCharset

For Q = 0 To Len(Paper.FontName) - 1
    ByteFontName(Q) = Asc(Mid$(Paper.FontName, Q + 1, 1))
Next
ByteFontName(Len(Paper.FontName)) = 0
'ByteFontName = Paper.FontName & vbNullChar

If setWallpaper <> "" Then PrepareWallPaper
End Sub

Public Sub ImitateMouseMove()
Dim lpPoint As POINTAPI
GetCursorPos lpPoint
'SetCursorPos lpPoint.X, lpPoint.Y + 1 '????? <<-------- will this work???
SetCursorPos lpPoint.X, lpPoint.Y
End Sub

Public Function BrowseForFolder(sPrompt As String) As String
Const BIF_RETURNONLYFSDIRS = 1, Max_Path = 260
Dim intNull As Integer, lngIdList As Long
Dim strPath As String, udtBI As BROWSEINFO

With udtBI
    .hWndOwner = 0
    .lpszTitle = sPrompt
    .ulFlags = BIF_RETURNONLYFSDIRS
End With
lngIdList = SHBrowseForFolder(udtBI)
If lngIdList Then
    strPath = String$(Max_Path, 0)
    SHGetPathFromIDList lngIdList, strPath
    CoTaskMemFree lngIdList
    intNull = InStr(strPath, vbNullChar)
    If intNull Then strPath = Left$(strPath, intNull - 1)
End If
strPath = AddDirSep(strPath)
BrowseForFolder = strPath
End Function

Public Function PrepareFileList(ByVal strPath As String, ByVal strPattern As String) As String()
On Local Error GoTo EH:
Dim tFileList() As String, tDirList() As String, tTempList() As String
Dim FCount As Long, DCount As Long, tCount As Long, Z As Long, Q As Long
Dim sHandle As Long, WFD As WIN32_FIND_DATA, tFilename As String, OldFileName As String
strPath = AddDirSep(strPath)
strPattern = LCase(strPattern)
ReDim tFileList(1 To 1)
ReDim tDirList(1 To 1)
ReDim tTempList(1 To 1)

sHandle = FindFirstFile(strPath & "*.*", WFD)
tFilename = Left(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1)
Do While tFilename <> ""
    If tFilename <> "." And tFilename <> ".." Then
        If (WFD.dwFileAttributes And vbDirectory) = vbDirectory Then
            DCount = DCount + 1
            ReDim Preserve tDirList(1 To DCount)
            tDirList(DCount) = AddDirSep(strPath & tFilename)
        Else
            If LCase(tFilename) Like strPattern Then
                FCount = FCount + 1
                ReDim Preserve tFileList(1 To FCount)
                tFileList(FCount) = strPath & tFilename
            End If
        End If
    End If
    FindNextFile sHandle, WFD
    If GetLastError = ERROR_NO_MORE_FILES Then Exit Do
    OldFileName = tFilename
    tFilename = Left(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1)
    If OldFileName = tFilename Then Exit Do
Loop
FindClose sHandle

If DCount > 0 Then
    tCount = 0
    For Z = 1 To DCount
        tTempList = PrepareFileList(tDirList(Z), strPattern)
        tCount = UBound(tTempList)
        If (tCount <> 0) And Not (tCount = 1 And tTempList(1) = "") Then
            FCount = FCount + tCount
            ReDim Preserve tFileList(1 To FCount)
            For Q = FCount - tCount + 1 To FCount
                tFileList(Q) = tTempList(Q - FCount + tCount)
            Next
        End If
    Next Z
End If
PrepareFileList = tFileList
Exit Function

EH:
MsgBox ERR.Description, vbExclamation
End Function

Public Sub BringToTop(ByVal hWnd As Long, Optional ByVal OnTop As Boolean = True)
SetWindowPos hWnd, IIf(OnTop, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Public Function ConvertEMF2WMF(ByVal hemf As Long) As Long
Dim mfData() As Byte, cbSize As Long
ReDim mfData(1 To 1)
cbSize = GetWinMetaFileBitsLen(hemf, 1, 0, MM_ANISOTROPIC, Paper.hDC)
If cbSize = 0 Then Exit Function
ReDim mfData(1 To cbSize)
If cbSize > 1 Then
    GetWinMetaFileBits hemf, cbSize, mfData(1), MM_ANISOTROPIC, Paper.hDC
End If
ConvertEMF2WMF = SetMetaFileBitsEx(cbSize, mfData(1))
End Function

Public Sub ShowToolTip(ByVal TTT As String, Optional ByVal X As Long = -1000, Optional ByVal Y As Long = -1000, Optional ByVal TimeOut As Long = 10)
Dim lpPoint As POINTAPI
Dim W As Long, H As Long

If CurrentToolTipText = TTT Or Not setShowTooltips Then Exit Sub

If X = -1000 Or Y = -1000 Then
    GetCursorPos lpPoint
    lpPoint.X = lpPoint.X + TTOffset
    lpPoint.Y = lpPoint.Y + TTOffset
End If
X = lpPoint.X * Screen.TwipsPerPixelX
Y = lpPoint.Y * Screen.TwipsPerPixelY

If ToolTipShown Then
    frmToolTip.Visible = False
Else
    Load frmToolTip
End If

frmToolTip.lblText = TTT
W = (frmToolTip.lblText.Width + TTTMargin * 2 + 2) * Screen.TwipsPerPixelX
H = (frmToolTip.lblText.Height + TTTMargin * 2 + 2) * Screen.TwipsPerPixelY
If X + W > GetSystemMetrics(SM_CXSCREEN) * Screen.TwipsPerPixelX Then
    X = X - W - 2 * TTOffset * Screen.TwipsPerPixelX
End If
If Y + H > GetSystemMetrics(SM_CYSCREEN) * Screen.TwipsPerPixelY Then
    Y = Y - H - 2 * TTOffset * Screen.TwipsPerPixelY
End If
frmToolTip.Move X, Y, W, H
'frmToolTip.Cls
'If setGradientFill Then
'    Gradient frmToolTip.hDC, vbWhite, frmToolTip.BackColor, 0, 0, frmToolTip.ScaleWidth, frmToolTip.ScaleHeight, False
'End If
'frmToolTip.CurrentX = TTTMargin + 1
'frmToolTip.CurrentY = TTTMargin + 1
'frmToolTip.Print TTT
'frmToolTip.Line (0, 0)-(frmToolTip.ScaleWidth - 1, frmToolTip.ScaleHeight - 1), 0, B
frmToolTip.StartTime = Timer
ShowWindow frmToolTip.hWnd, SHOW_OPENNOACTIVATE

ToolTipShown = True
CurrentToolTipText = TTT
If TimeOut <> 0 Then
    frmToolTip.tmrTimer1.Tag = TimeOut
    frmToolTip.tmrTimer1.Enabled = True
End If

End Sub

Public Sub ShowAbout()
On Local Error GoTo EH
Dim PreviousTime As Long
Const DelayTime As Long = 40

#If conUseGLAbout = 1 Then
    frmAbout.Show
    
    If InitializeOpenGL(frmAbout.hWnd) <> 0 Then
        InitializeFirework
        
        Do While Not frmAbout.WasResponce
            AdvanceFirework
            CheckForInitialize
            Do
                DoEvents
            Loop Until timeGetTime >= PreviousTime + DelayTime
            PreviousTime = timeGetTime
        Loop
        
        TerminateOpenGL
        Unload frmAbout
    Else
        frmAboutSimple.Show vbModal
    End If
#Else
    frmAboutSimple.Show vbModal
#End If

EH:
End Sub

Public Sub MakeFullScreen(ByVal hWnd As Long, Optional ByVal OnOff As Boolean = True)
If OnOff Then
    MenuHandle = GetMenu(hWnd)
    WindowStyle = GetWindowLong(hWnd, GWL_STYLE)
    SetMenu hWnd, 0
    SetWindowLong hWnd, GWL_STYLE, WS_VISIBLE Or WS_MAXIMIZE 'Or WS_THICKFRAME
    SetWindowPos hWnd, 0, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_FRAMECHANGED
Else
    SetMenu hWnd, MenuHandle
    SetWindowLong hWnd, GWL_STYLE, WindowStyle
    SetWindowPos hWnd, 0, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_FRAMECHANGED
End If
End Sub

Public Sub PrintFile()
If Printers.Count < 1 Then Exit Sub
Printer.ScaleMode = vbPixels
ShowAll Printer.hDC
Printer.EndDoc
End Sub

Public Sub ParseCommandLine()
Dim Cmd As String
Cmd = Trim(LCase(Command))
If Cmd <> "" Then
    If InStr(Cmd, "debug") Then InDebugMode = True Else InDebugMode = False
    Cmd = Trim(Replace(Cmd, Chr(34), ""))
    If Dir(Cmd) <> "" And Right(Cmd, 3) = LCase(extFIG) Then CommandLineFile = Cmd Else CommandLineFile = ""
End If
End Sub

Public Sub OpenCommandLineFile()
DrawingName = CommandLineFile
LastFigurePath = AddDirSep(RetrieveDir(DrawingName))
If Not IsValidPath(LastFigurePath) Then LastFigurePath = ProgramPath
ClearPrivileges
OpenFile DrawingName
AddMRUItem DrawingName
End Sub

Public Sub PreparePaths()
On Local Error Resume Next
ProgramPath = AddDirSep(App.Path)

If Dir(ProgramPath & HelpFileRussian) <> "" Then HelpPathRussian = ProgramPath & HelpFileRussian
If Dir(ProgramPath & HelpFileEnglish) <> "" Then HelpPathEnglish = ProgramPath & HelpFileEnglish
If Dir(ProgramPath & HelpFileGerman) <> "" Then HelpPathGerman = ProgramPath & HelpFileGerman
If Dir(ProgramPath & HelpFileUkrainian) <> "" Then HelpPathUkrainian = ProgramPath & HelpFileUkrainian

PrepareHelpFile
LoadLastPaths
End Sub

Public Sub PrepareHelpFile()
App.HelpFile = ""
If setLanguage = langEnglish And HelpPathEnglish <> "" Then App.HelpFile = HelpPathEnglish
If setLanguage = langRussian And HelpPathRussian <> "" Then App.HelpFile = HelpPathRussian
If setLanguage = langGerman And HelpPathGerman <> "" Then App.HelpFile = HelpPathGerman
If setLanguage = langUkrainian And HelpPathUkrainian <> "" Then App.HelpFile = HelpPathUkrainian
End Sub

Public Sub ProcessAutoloadMacros()
If setMacroAutoloadPath <> "" Then
    If Not IsValidPath(setMacroAutoloadPath) Then Exit Sub
    FormMain.Enabled = False
    i_ShowStatus GetString(ResWorkingPleaseWait)
    FormMain.Refresh
    i_SetMousePointer curStateHourglass
    AutoLoadMacros
    i_SetMousePointer curStateArrow
    FormMain.Enabled = True
    If FormMain.Visible Then FormMain.SetFocus: i_ShowStatus
End If
End Sub

Public Function IsComment(ByVal S1 As String) As Boolean
Dim RemChar As String
RemChar = Left(S1, 1)
If S1 = "" Or RemChar = "/" Or RemChar = "'" Or RemChar = ";" Then IsComment = True
End Function

Public Function IsSectionHeader(ByVal S1 As String) As Boolean
IsSectionHeader = Left(S1, 1) = "[" And Right(S1, 1) = "]"
End Function

Public Function GetSectionHeader(ByVal S1 As String) As String
GetSectionHeader = Mid(S1, 2, Len(S1) - 2)
End Function

Public Function IsINIProperty(ByVal S1 As String) As Boolean
If InStr(S1, "=") Then If Not IsComment(S1) And Not IsSectionHeader(S1) Then IsINIProperty = True
End Function

Public Function GetINIPropertyName(ByVal S1 As String) As String
GetINIPropertyName = Left(S1, InStr(S1, "=") - 1)
End Function

Public Function GetINIPropertyValue(ByVal S1 As String) As String
GetINIPropertyValue = Right(S1, Len(S1) - InStr(S1, "="))
End Function

Public Function ToBoolean(ByVal B1 As Boolean) As String
ToBoolean = IIf(B1, "True", "False")
End Function

Public Function FromBoolean(ByVal S1 As String) As Boolean
On Local Error GoTo EH:
FromBoolean = CBool(S1)
Exit Function

EH:
ERR.Clear
FromBoolean = S1 = "True" Or S1 = "Истина" Or S1 = "#TRUE#"
End Function

Public Function IsValidPath(ByVal FName As String) As Boolean
Dim D, Z As Long, A As String, S As String
Dim OldDrive As String

On Local Error Resume Next
ERR.Clear

If Len(FName) < 3 Then IsValidPath = False: Exit Function

FName = RemoveDirSep(FName)
D = Split(FName, "\")

S = ""
OldDrive = Left(CurDir, 3)

A = UCase(Left(FName, 1))
If A < "A" Or A > "Z" Then IsValidPath = False: Exit Function
If (2 ^ (Asc(A) - Asc("A")) And GetLogicalDrives) = 0 Then IsValidPath = False: Exit Function

ChDrive Left(FName, 3)
If ERR <> 0 Then
    ERR.Clear
    IsValidPath = False
    Exit Function
End If

For Z = LBound(D) To UBound(D)
    S = S & D(Z) & "\"
    If Dir(S, 23) = "" Then
        IsValidPath = False
        Exit Function
    End If
Next

ChDrive OldDrive

IsValidPath = True
End Function

Public Sub EnableTextBox(obj As Object, Enable As Boolean)
obj.Enabled = Enable
obj.BackColor = IIf(Enable, vbWindowBackground, vbButtonFace)
End Sub

Public Sub ReportError(ByVal S1 As String)
MsgBox "Передайте, пожалуйста, Кириллу:" & vbCrLf & S1 & vbCrLf & vbCrLf & "(По возможности, сохраните этот файл)", vbOKOnly + vbExclamation, GetString(ResError)
End Sub

Public Function FontExists(ByVal FontName As String) As Boolean
FontExists = FontList.FindItem(FontName) <> 0
End Function

Public Function PickCharset(ByVal Str1 As String) As Long
Dim Z As Long
For Z = 1 To Len(Str1)
    If Asc(Str1) > 127 Then PickCharset = DefaultFontCharset: Exit Function
Next
End Function

Public Function ToSingleLine(ByVal S1 As String) As String
ToSingleLine = IIf(InStr(S1, vbCrLf) = 0, S1, Replace(S1, vbCrLf, "~~"))
End Function

Public Function ToMultiLine(ByVal S1 As String) As String
ToMultiLine = IIf(InStr(S1, "~~") = 0, S1, Replace(S1, "~~", vbCrLf))
End Function

Public Function ColorToHex(ByVal Col1 As Long) As String
ColorToHex = "#" & Leading0(Hex(Red(Col1)), 2) & Leading0(Hex(Green(Col1)), 2) & Leading0(Hex(Blue(Col1)), 2)
End Function

Public Function Leading0(ByVal S1 As String, ByVal NumOf0 As Integer) As String
If Len(S1) >= NumOf0 Then Leading0 = S1 Else Leading0 = String(NumOf0 - Len(S1), "0") & S1
End Function

Public Function Str2Color(ByVal S1 As String) As Long
If Left(S1, 1) = "#" Then
    Str2Color = RGB(Val("&H" & Mid(S1, 2, 2)), Val("&H" & Mid(S1, 4, 2)), Val("&H" & Mid(S1, 6, 2)))
Else
    Str2Color = Val(S1)
End If
End Function

Public Sub FillFonts()
FontList.Clear
EnumFonts GetWindowDC(GetDesktopWindow), vbNullString, AddressOf EnumFontsProc, 0
End Sub

Public Function EnumFontsProc(hLogFont As LOGFONT, hTextMetric As TEXTMETRIC, lType As Long, lParam As Long) As Integer
On Local Error Resume Next
'Dim S As String * 32
'S = hLogFont.lfFaceName
FontList.Add ByteArrayToString(hLogFont.lfFaceName)
EnumFontsProc = 1
End Function

Public Function ByteArrayToString(abBytes() As Byte) As String
Dim lBytePoint As Long
Dim lByteVal As Long
Dim sOut As String

'init array pointer
lBytePoint = LBound(abBytes)

'fill sOut with characters in array
While lBytePoint <= UBound(abBytes)
    
    lByteVal = abBytes(lBytePoint)
    
    'return sOut and stop if Chr$(0) is encountered
    If lByteVal = 0 Then
        ByteArrayToString = sOut
        Exit Function
    Else
        sOut = sOut & Chr$(lByteVal)
    End If
    
    lBytePoint = lBytePoint + 1

Wend

'return sOut if Chr$(0) wasn't encountered
ByteArrayToString = sOut
End Function

Public Sub ProgressShow(Optional ByVal S As String, Optional ByVal Quiet As Boolean = False)
If S <> "" Then ProgressMsg = S
ProgressQuiet = Quiet
If Not Quiet Then
    frmProgressWnd.Show
End If
End Sub

Public Sub ProgressUpdate(Optional ByVal Percentage As Double)
If ProgressQuiet Then
    i_ShowStatus ProgressMsg & " " & Format(Percentage, "#0%")
    FormMain.StatusBar.Refresh
Else
    frmProgressWnd.Progress Percentage
End If
End Sub

Public Sub ProgressClose()
If ProgressQuiet Then
    i_ShowStatus ""
Else
    Unload frmProgressWnd
End If
End Sub

Public Sub ProgressRefresh()
If ProgressQuiet Then
    FormMain.Status(1).Refresh
Else
    frmProgressWnd.Refresh
End If
End Sub

Public Function StripNull(ByVal S As String) As String
Dim Z As Long
Z = InStr(S, vbNullChar)
If Z = 0 Then StripNull = S Else StripNull = Left(S, Z - 1)
End Function

'Public Sub LaunchHelp(ByVal Index As Long)
'If App.HelpFile <> "" Then
'    CD.HelpFile = App.HelpFile
'    CD.HelpCommand = Index
'    CD.ShowHelp
'Else
'    MsgBox GetString(ResHelpFileNotFound) & " " & ProgramPath
'End If
'End Sub

Public Function GetAbsolutePath(ByVal RelativePath As String, ByVal CurrentPath As String) As String
Dim DirsToProcess, Z As Long, A As String
Dim DF As String
Dim SD As String, DD As String

DF = RetrieveName(RelativePath)
DD = RemoveDirSep(RetrieveDir(RelativePath))
SD = RemoveDirSep(RetrieveDir(AddDirSep(CurrentPath)))

DirsToProcess = Split(DD, "\")

For Z = LBound(DirsToProcess) To UBound(DirsToProcess)
    A = DirsToProcess(Z)
    If A = "." Then
        'do nothing
    ElseIf A = ".." Then
        SD = RemoveDirSep(RetrieveParentDir(SD))
    Else
        SD = SD & "\" & A
    End If
Next

If DF <> "" Then
    SD = SD & "\" & DF
Else
    SD = SD & "\"
End If

GetAbsolutePath = SD
End Function

Public Function RetrieveParentDir(ByVal Path As String) As String
Path = RemoveDirSep(Path)
RetrieveParentDir = RetrieveDir(Path)
End Function

Public Function RemoveDirSep(ByVal S As String) As String
If S <> "" Then
    If Right(S, 1) = "\" Then
        RemoveDirSep = Left(S, Len(S) - 1)
    Else
        RemoveDirSep = S
    End If
Else
    RemoveDirSep = ""
End If
End Function

Public Function AddDirSep(ByVal sStr As String) As String
If sStr <> "" Then If Right(sStr, 1) <> "\" Then sStr = sStr & "\"
AddDirSep = sStr
End Function

Public Function IsOKVersion() As Boolean
On Local Error GoTo EH
Dim A As String, Z As Long
Dim K As Double, D As Double
Dim DGRegTxt As String

#If conProtected = 1 Then
    DGRegTxt = ProgramPath & RegTxt
    If Dir(DGRegTxt, 23) = "" Then IsOKVersion = False: Exit Function
    
    Open DGRegTxt For Input As #1
        D = PerversionBase \ (2 ^ 4 + 1)
        For Z = 1 To 22
            Line Input #1, A
            K = Val(A)
            If K > PerversionBase Then K = K - (K \ PerversionBase) * PerversionBase
            D = D Xor K
        Next
        Line Input #1, A
        If Val(A) <> D Then IsOKVersion = False: Close #1: Exit Function
    Close #1
    
    Open DGRegTxt For Input As #1
        For Z = 1 To 1 + 2 + 4
            Line Input #1, A
        Next Z
        Line Input #1, A
    Close #1
    
    If Val(A) = Pervert(Val(Trim(Oct(Int((Val(Trim(Str(Pervert(GetS)))) / 17)))))) Then
        IsOKVersion = True
    Else
        IsOKVersion = False
    End If
    
#Else
    IsOKVersion = True
#End If
Exit Function

EH:
IsOKVersion = False
End Function

Public Function GetS(Optional ByVal Path As String = "") As String
Dim VolumeNameBuffer As String
Dim RootPathName As String
Dim VolumeSerialNumber As Long
Dim VolumeNameSize As Long
Dim MaximumComponentLength As Long
Dim FileSystemFlags As Long
Dim FileSystemNameBuffer As String
Dim FileSystemNameSize As Long
Dim A As String

If Path <> "" Then A = Path Else A = ProgramPath
RootPathName = UCase(Left(A, 3))
VolumeNameSize = 11
GetVolumeInformation RootPathName, VolumeNameBuffer, VolumeNameSize, VolumeSerialNumber, MaximumComponentLength, FileSystemFlags, FileSystemNameBuffer, FileSystemNameSize

GetS = Trim(Str(VolumeSerialNumber))
End Function

Public Function Pervert(ByVal N As Double) As Double
Pervert = PerversionBase - N
End Function

Public Sub AntiAliasHDC(ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)
On Local Error GoTo EH

Dim i As Long, j As Long, I1 As Long, J1 As Long
Dim n1 As Long, n2 As Long, N3 As Long
Dim W As Long, H As Long
Dim AvRow(-1 To 1) As Long
Dim BInfo As BITMAPINFO
Dim CompDC As Long
Dim Addr As Long
Dim DIBSectionHandle As Long
Dim OldCompDCBM As Long
Dim BytesPerScanLine As Long
Dim DI As Long
Dim RGB1 As TRIRGB, RGB2 As TRIRGB, RGB3 As TRIRGB

W = X2 - X1 + 1
H = Y2 - Y1 + 1

CompDC = CreateCompatibleDC(hDC)

With BInfo.bmiHeader
    .biSize = 40
    .biWidth = W
    .biHeight = -H
    .biPlanes = 1
    .biBitCount = 24
    .biCompression = BI_RGB
    BytesPerScanLine = (.biWidth * 3 + 3) And &HFFFFFFFC
    .biSizeImage = BytesPerScanLine * .biHeight
End With
DIBSectionHandle = CreateDIBSection(CompDC, BInfo, DIB_RGB_COLORS, Addr, 0, 0)
OldCompDCBM = SelectObject(CompDC, DIBSectionHandle)

PrepareHDC CompDC
PaperCls CompDC
ShowAll CompDC

For i = 2 To W - 1
    n1 = 0
    n2 = 0
    N3 = 0
    CopyMemory n1, ByVal (Addr + (i - 2 + 1) * 3 - 3), 3
    CopyMemory n2, ByVal (Addr + (i - 2 + 2) * 3 - 3), 3
    CopyMemory N3, ByVal (Addr + (i - 2 + 3) * 3 - 3), 3
    AvRow(0) = (n1 + n2 + N3) \ 3
    n1 = 0
    n2 = 0
    N3 = 0
    CopyMemory n1, ByVal (Addr + (i - 2 + W + 1) * 3 - 3), 3
    CopyMemory n2, ByVal (Addr + (i - 2 + W + 2) * 3 - 3), 3
    CopyMemory N3, ByVal (Addr + (i - 2 + W + 3) * 3 - 3), 3
    AvRow(1) = (n1 + n2 + N3) \ 3
    For j = 2 To H - 1
        AvRow(-1) = AvRow(0)
        AvRow(0) = AvRow(1)
        n1 = 0
        n2 = 0
        N3 = 0
        'CopyMemory N1, ByVal (Addr + ((J - 1) * W + I + 1) * 3 - 3), 3
        'CopyMemory N2, ByVal (Addr + (J * W + I + 1) * 3 - 3), 3
        CopyMemory RGB2, ByVal (Addr + (j * W + i + 3) * 3 - 3), 3

        'CopyMemory N3, ByVal (Addr + ((J + 1) * W + I + 1) * 3 - 3), 3
        AvRow(1) = (n1 + n2 + N3) \ 3
        SetPixelV hDC, i, j, RGB(RGB2.Red, RGB2.Green, RGB2.Blue) '(AvRow(-1) + AvRow(0) + AvRow(1)) \ 3
    Next
Next

DI = SelectObject(CompDC, OldCompDCBM)
DI = DeleteObject(DIBSectionHandle)
DI = DeleteDC(CompDC)
Exit Sub

EH:
MsgBox "Unable to antialias", vbCritical
End Sub

Public Function LOWORD(ByVal N As Long) As Long
Dim K As Long
K = N And 65535
If K < 32768 Then LOWORD = K Else LOWORD = K - 65536
End Function

Public Function HIWORD(ByVal N As Long) As Long
Dim K As Long
K = N \ 65536 And 65535
If K < 32768 Then HIWORD = K Else HIWORD = K - 65536
End Function

Public Sub CenterForm(frm As Form)
frm.Move (Screen.Width - frm.Width) \ 2, (Screen.Height - frm.Height) \ 2
End Sub

Public Sub AddListboxScrollbar(Lst1 As ListBox)
Dim Z As Long, W As Long, m As Long
If Lst1.ListCount = 0 Then Exit Sub

For Z = 0 To Lst1.ListCount - 1
    W = Lst1.Parent.ScaleX(Lst1.Parent.TextWidth(Lst1.List(Z)), Lst1.Parent.ScaleMode, vbPixels) + IIf(Lst1.Style = 1, 20, 8)
    If W > m Then m = W
Next

SendMessage Lst1.hWnd, &H194, m, 0
End Sub

Public Function IsAlpha(ByVal S As String) As Boolean
IsAlpha = Len(S) = 1 And S >= "A" And S <= "Z"
End Function

Public Function IsAlphaNumeric(ByVal S As String) As Boolean
IsAlphaNumeric = Len(S) = 1 And ((S >= "A" And S <= "Z") Or (S >= "0" And S <= "9"))
End Function

Public Function IsAlphaNumericEx(ByVal S As String) As Boolean
IsAlphaNumericEx = Len(S) = 1 And ((S >= "A" And S <= "Z") Or (S >= "0" And S <= "9")) Or S = "_"
End Function

Public Function IsLegalVariableName(ByVal S As String) As Boolean
Dim Z As Long

If S = "" Then Exit Function
If Not IsAlpha(Left(S, 1)) Then Exit Function

For Z = 1 To Len(S)
    If Not IsAlphaNumeric(Mid(S, Z, 1)) Then Exit Function
Next

IsLegalVariableName = True
End Function

Public Function Proper(ByVal S As String) As String
Proper = StrConv(S, vbProperCase)
End Function

Public Function GetFormatString(Optional ByVal Precision As Long = 0) As String
If Precision = 0 Then GetFormatString = "0" Else GetFormatString = "0." & String(Precision, "0")
End Function

Public Sub RedimBasePoint(ByVal LB As Long, ByVal UB As Long)
If UB < LB Then
    ReDim BasePoint(1 To 1)
    'ReDim OrderedPoints(1 To 1)
    Exit Sub
End If
ReDim BasePoint(LB To UB)
'ReDim OrderedPoints(LB To UB)
End Sub

Public Sub RedimPreserveBasePoint(ByVal LB As Long, ByVal UB As Long)
If UB < LB Then
    ReDim BasePoint(1 To 1)
    'ReDim OrderedPoints(1 To 1)
    Exit Sub
End If
ReDim Preserve BasePoint(LB To UB)
'ReDim OrderedPoints(LB To UB)
End Sub

Public Sub RedimFigures(ByVal LB As Long, ByVal UB As Long)
If UB < LB Then
    ReDim Figures(0 To 0)
    ReDim OrderedFigures(0 To 0)
    Exit Sub
End If
ReDim Figures(LB To UB)
ReDim OrderedFigures(LB To UB)
End Sub

Public Sub RedimPreserveFigures(ByVal LB As Long, ByVal UB As Long)
If UB < LB Then
    ReDim Figures(1 To 1)
    ReDim OrderedFigures(1 To 1)
    Exit Sub
End If
ReDim Preserve Figures(LB To UB)
ReDim OrderedFigures(LB To UB)
End Sub

'==============================================
' DEBUG section
'==============================================

#If conDebug = 1 Then

    Public Sub AddManyPoints(Optional ByVal HowMany As Long = 1000)
    Dim Z As Long
    For Z = 1 To HowMany
        AddBasePoint Random(CanvasBorders.P1.X, CanvasBorders.P2.X), Random(CanvasBorders.P1.Y, CanvasBorders.P2.Y)
    Next
    End Sub

#End If

Public Function LargeInteger(L As LARGE_INTEGER) As Variant
Dim LargeNum
LargeNum = L.LowPart
If LargeNum < 0 Then LargeNum = LargeNum + 2 ^ 32
LargeInteger = CDec(L.HighPart * 2 ^ 32 + LargeNum)
End Function
