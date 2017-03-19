Attribute VB_Name = "modDrawAux"
Option Explicit

Public Sub PrepareHDC(ByVal hDC As Long)
DefaultFontCharset = Paper.Font.Charset
SetROP2 hDC, vbCopyPen
SetBkColor hDC, Paper.BackColor
SetBkMode hDC, Transparent
SetTextAlign hDC, TA_BOTTOM Or TA_CENTER
End Sub

Public Function WriteMetafile(ByVal FName As String, MFHnd As Long, xExt As Long, yExt As Long)
Dim fHnd As Long
Dim DI As Long, dl As Long
Dim mfglbhnd As Long
Dim NewMF As Long
Dim DC&
Dim mFile As METAFILEHEADER
Dim mfinfosize&
Dim currentfileloc&
Dim gptr&
Dim oldsize As Size

' Open the file to write
fHnd = lcreat(FName, 0)
If fHnd >= 0 Then Call lclose(fHnd) ' Close the open handle
fHnd = lopen(FName, 2)
If fHnd < 0 Then Exit Function
If MFHnd = 0 Then Exit Function

' First write a placeable header file header
mFile.Key = &H9AC6CDD7  ' The key - required
mFile.hMF = 0           ' Must be 0
mFile.bbox.Left = 0
mFile.bbox.Top = 0
' These should be calculated using GetDeviceCaps
mFile.bbox.Right = xExt + 1 ' Size in metafile units of bounding area
mFile.bbox.Bottom = yExt + 1
mFile.inch = Paper.ScaleX(1, vbInches, vbPixels) '1000 ' Number of metafile units per inch

mFile.reserved = 0
' Build the checksum
mFile.checksum = &H9AC6 Xor &HCDD7 ' 9ac6 xor cdd7
mFile.checksum = mFile.checksum Xor mFile.bbox.Right
mFile.checksum = mFile.checksum Xor mFile.bbox.Bottom
mFile.checksum = mFile.checksum Xor mFile.inch

' Write the buffer
DI = lwrite(fHnd, mFile, Len(mFile))

' Now we retrieve a handle that will contain the
' metafile  - We make a copy, but first we set the
' extents so that it can be properly displayed
DC = CreateMetaFile(vbNullString)
dl = SetWindowExtEx(DC, xExt, yExt, oldsize)
DI = SetMapMode(DC, MM_ANISOTROPIC)
DI = PlayMetaFile(DC, MFHnd)
NewMF = CloseMetaFile(DC)

' Find out how bit the buffer needs to be
mfinfosize = GetMetaFileBitsEx(NewMF, 0, ByVal 0)
If mfinfosize = 0 Then
    DI = lclose(fHnd)
    Exit Function
End If
mfglbhnd = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, mfinfosize)
gptr = GlobalLock(mfglbhnd)
dl = GetMetaFileBitsEx(NewMF, mfinfosize, ByVal gptr)

dl = hwrite(fHnd, ByVal gptr, mfinfosize)

DI = GlobalUnlock(mfglbhnd)
DI = GlobalFree(mfglbhnd)
DI = lclose(fHnd)
End Function

Public Sub Gradient(ByVal DC As Long, ByVal SCol As Long, ByVal DCol As Long, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, Optional ByVal Vertical As Boolean = True)
On Local Error GoTo EH:

Const Details = 92

Dim DrawColorR1 As Integer
Dim DrawColorG1 As Integer
Dim DrawColorB1 As Integer
Dim DrawColorR2 As Integer
Dim DrawColorG2 As Integer
Dim DrawColorB2 As Integer
Dim hBrush As Long
Dim hOldBrush As Long
Dim InvHeight As Double
Dim InvWidth As Double
Dim i As Long, St As Long, Coeff As Double
Dim lpR As RECT

If Height < 1 Or Width < 1 Then Exit Sub
InvHeight = 1 / Height
InvWidth = 1 / Width

If SCol < 0 Then SCol = GetSysColor(SCol + SysColorTranslationBase)
If DCol < 0 Then DCol = GetSysColor(DCol + SysColorTranslationBase)

DrawColorR1 = SCol And 255
DrawColorG1 = (SCol And 65535) \ 256
DrawColorB1 = (SCol \ 65536)
DrawColorR2 = (DCol And 255) - DrawColorR1
DrawColorG2 = ((DCol And 65535) \ 256) - DrawColorG1
DrawColorB2 = (DCol \ 65536) - DrawColorB1

If Vertical Then
    St = Width \ Details
    If St < 1 Then St = 1
    lpR.Top = Top
    lpR.Bottom = Top + Height
    
    For i = Left To Left + Width Step St
        Coeff = (i - Left) * InvWidth
        hBrush = CreateSolidBrush(RGB(DrawColorR1 + DrawColorR2 * Coeff, DrawColorG1 + DrawColorG2 * Coeff, DrawColorB1 + DrawColorB2 * Coeff))
        lpR.Left = i
        lpR.Right = i + St
        FillRect DC, lpR, hBrush
        DeleteObject hBrush
    Next i
    
Else
    
    St = Height \ Details
    If St < 1 Then St = 1
    lpR.Left = Left
    lpR.Right = Left + Width
    
    For i = Top To Top + Height Step St
        Coeff = (i - Top) * InvHeight
        hBrush = CreateSolidBrush(RGB(DrawColorR1 + DrawColorR2 * Coeff, DrawColorG1 + DrawColorG2 * Coeff, DrawColorB1 + DrawColorB2 * Coeff))
        lpR.Top = i
        lpR.Bottom = i + St
        FillRect DC, lpR, hBrush
        DeleteObject hBrush
    Next i
    
End If

EH:
End Sub

Public Function Red(ByVal RGB As Long) As Long
Red = RGB And 255
End Function

Public Function Green(ByVal RGB As Long) As Long
Green = (RGB And 65535) \ 256
End Function

Public Function Blue(ByVal RGB As Long) As Long
Blue = RGB \ 65536
End Function

Public Sub FillRGB(ByVal Col As Long, R As Long, G As Long, B As Long)
Col = EnsureRGB(Col)
R = Red(Col)
G = Green(Col)
B = Blue(Col)
End Sub

Public Function DarkenColor(ByVal Col As Long, ByVal Percentage As Single) As Long
Col = EnsureRGB(Col)
DarkenColor = RGB(Red(Col) * Percentage, Green(Col) * Percentage, Blue(Col) * Percentage)
End Function

Public Function IsColorDark(ByVal Col As Long) As Boolean
Dim R As Long, G As Long, B As Long, m As Long
Col = EnsureRGB(Col)
m = (Red(Col) + Green(Col) + Blue(Col)) / 3
IsColorDark = m < 128
End Function

Public Function GetGrayedColor(ByVal Col As Long, Optional ByVal Weak As Boolean = False) As Long
Const Deviation = 24

Dim R1 As Long, G1 As Long, B1 As Long
Dim R2 As Long, G2 As Long, B2 As Long
Dim R As Long, B As Long
Dim PB As Long
Dim Dev As Long
Dim Span As Single

Dev = Deviation
If Weak Then Dev = Dev / 2
Span = Dev - 4

PB = EnsureRGB(Paper.BackColor)
Col = EnsureRGB(Col)

R1 = Red(PB)
G1 = Green(PB)
B1 = Blue(PB)

B = (R1 + G1 + B1) \ 3
If B < 56 Then B = 56
If B = 128 Then B = 129
B = B + Dev * Sgn(128 - B)

R2 = B + (Red(Col) - 128) * Span / 256#
G2 = B + (Green(Col) - 128) * Span / 256#
B2 = B + (Blue(Col) - 128) * Span / 256#

If R2 < 0 Then R2 = 0
If G2 < 0 Then G2 = 0
If B2 < 0 Then B2 = 0
If R2 > 255 Then R2 = 255
If G2 > 255 Then G2 = 255
If B2 > 255 Then B2 = 255

GetGrayedColor = RGB(R2, G2, B2)
End Function

Public Sub ShadowRect(ByVal hDC As Long, ptRect As RECT, ByVal Col As Long, Optional ByVal ShadowSize As Long = Shadow, Optional ByVal TransparentShadow As Boolean = False)
Dim ptRect3 As RECT, hBrush As Long, hOldBrush As Long
Dim Ratio As Single, tRed As Long, tGreen As Long, tBlue As Long, Z As Long
Dim hPen As Long, hOldPen As Long
ptRect3 = ptRect
If Not setGradientFill Then Exit Sub

Col = EnsureRGB(Col)
If TransparentShadow Then SetROP2 hDC, vbMaskPen

tRed = Red(Col)
tGreen = Green(Col)
tBlue = Blue(Col)
OffsetRect ptRect3, ShadowSize + 1, ShadowSize + 1

For Z = ShadowSize To 1 Step -1
    OffsetRect ptRect3, -1, -1
    Ratio = Sin(Z / ShadowSize + PI / 2 - 1)
    Col = RGB(tRed * Ratio, tGreen * Ratio, tBlue * Ratio)
    
    hPen = CreatePen(PS_SOLID, 1, Col)
    hOldPen = SelectObject(hDC, hPen)
    hOldBrush = SelectObject(hDC, GetStockObject(NULL_BRUSH))
    
    RoundRect hDC, ptRect3.Left, ptRect3.Top, ptRect3.Right, ptRect3.Bottom, Z * 2, Z * 2
    'Rectangle hDC, ptRect3.Left, ptRect3.Top, ptRect3.Right, ptRect3.Bottom
    SetPixelV hDC, ptRect3.Right - Z, ptRect3.Bottom - Z, Col
    
    SelectObject hDC, hOldBrush
    SelectObject hDC, hOldPen
    DeleteObject hPen
Next

If TransparentShadow Then SetROP2 hDC, vbCopyPen
End Sub

Public Sub ShadowControl(pControl As Object, Optional ByVal Col As Long = -1, Optional ByVal ShadowSize As Long = Shadow, Optional ByVal TransparentShadow As Boolean = False)
On Local Error GoTo EH
Dim ptRect As RECT
If Not setGradientFill Then Exit Sub
If Col = -1 Then Col = pControl.Container.BackColor
ptRect.Left = pControl.Left
ptRect.Top = pControl.Top
ptRect.Right = ptRect.Left + pControl.Width
ptRect.Bottom = ptRect.Top + pControl.Height
ShadowRect pControl.Container.hDC, ptRect, Col, ShadowSize, TransparentShadow
EH:
End Sub

Public Sub PrepareWallPaper()
Dim Pict As IPictureDisp, tX As Long, tY As Long, X As Long, Y As Long
Dim SW As Long, SH As Long
Dim OldRPR As Boolean

Set Pict = LoadPicture(setWallpaper)

SW = GetSystemMetrics(SM_CXSCREEN)
SH = GetSystemMetrics(SM_CYSCREEN)

OldRPR = RestrictPaperResize
RestrictPaperResize = True
Paper.Move 0, 0, SW, SH
RestrictPaperResize = OldRPR

tX = Paper.ScaleX(Pict.Width, vbHimetric, vbPixels)
tY = Paper.ScaleY(Pict.Height, vbHimetric, vbPixels)
For X = 0 To SW \ tX
    For Y = 0 To SH \ tY
        Paper.PaintPicture Pict, tX * X, tY * Y
    Next Y
Next X
BitBlt ServiceHDC, 0, 0, SW, SH, Paper.hDC, 0, 0, SRCCOPY
End Sub

Public Sub DrawWallPaper(ByVal hDC As Long)
BitBlt hDC, 0, 0, PaperScaleWidth, PaperScaleHeight, ServiceHDC, 0, 0, SRCCOPY
End Sub

Public Function EnsureRGB(ByVal Col As Long) As Long
If Col < 0 Then Col = GetSysColor(Col + SysColorTranslationBase)
EnsureRGB = Col
End Function

Public Function PrintText(ByVal Canv As Object, ByVal szS As String, Optional ByVal X As Long = EmptyVar, Optional ByVal Y As Long = EmptyVar, Optional ByVal Ang As Single) As Boolean
Dim LF As LOGFONT, i As Long, NF As Long, tS As String, Q As Long

LF.lfWidth = 0
LF.lfEscapement = CLng(Ang * 10)
LF.lfOrientation = LF.lfEscapement
LF.lfWeight = Canv.Font.Weight
LF.lfItalic = IIf(Canv.FontItalic, 255, 0)
LF.lfUnderline = IIf(Canv.FontUnderline, 255, 0)
LF.lfStrikeOut = IIf(Canv.FontStrikethru, 255, 0)
LF.lfCharSet = 0 'Paper.Font.Charset
LF.lfOutPrecision = 0
LF.lfClipPrecision = 0
LF.lfQuality = 2
LF.lfPitchAndFamily = 0
tS = Canv.FontName
'LF.lfFaceName = tS & vbNullChar
If Len(tS) > 31 Then tS = "Arial"
For Q = 1 To Len(tS)
    LF.lfFaceName(Q) = Asc(Mid$(tS, Q, 1))
Next
LF.lfFaceName(Len(tS) + 1) = 0
LF.lfHeight = Canv.FontSize * -20 / Screen.TwipsPerPixelY

i = CreateFontIndirect(LF)
NF = SelectObject(Canv.hDC, i)
If X = EmptyVar Then X = Canv.CurrentX
If Y = EmptyVar Then Y = Canv.CurrentY
SetTextAlign Canv.hDC, TA_BOTTOM Or TA_CENTER
TextOut Canv.hDC, X, Y, szS, Len(szS)
SelectObject Canv.hDC, NF
DeleteObject i

End Function

Public Function PrintTextAPI(ByVal hDC As Long, ByVal szS As String, Optional ByVal ForeColor As Long, Optional ByVal X As Long = EmptyVar, Optional ByVal Y As Long = EmptyVar, Optional ByVal FontName As String, Optional ByVal FontSize As Long = 8, Optional ByVal Bold As Boolean = False, Optional ByVal Italic As Boolean = False, Optional ByVal Underline As Boolean = False, Optional ByVal Ang As Single, Optional ByVal Charset As Long = 0) As Boolean
Dim LF As LOGFONT, i As Long, NF As Long, tS As String, lpPoint As POINTAPI, Q As Long

LF.lfWidth = 0
LF.lfEscapement = CLng(Ang * 10)
LF.lfOrientation = LF.lfEscapement
LF.lfWeight = IIf(Bold, 700, 400)
LF.lfItalic = IIf(Italic, 255, 0)
LF.lfUnderline = IIf(Underline, 255, 0)
LF.lfStrikeOut = 0
LF.lfCharSet = Charset
LF.lfOutPrecision = 0
LF.lfClipPrecision = 0
LF.lfQuality = 2
LF.lfPitchAndFamily = 0
tS = FontName
'LF.lfFaceName = tS & vbNullChar
If Len(tS) > 31 Then tS = "Arial"
For Q = 1 To Len(tS)
    LF.lfFaceName(Q) = Asc(Mid$(tS, Q, 1))
Next
LF.lfFaceName(Len(tS) + 1) = 0
LF.lfHeight = FontSize * -20 / Screen.TwipsPerPixelY

i = CreateFontIndirect(LF)
NF = SelectObject(hDC, i)
If X = EmptyVar Or Y = EmptyVar Then
    GetCurrentPositionEx hDC, lpPoint
    X = lpPoint.X
    Y = lpPoint.Y
End If
TextOut hDC, X, Y, szS, Len(szS)
SelectObject hDC, NF
DeleteObject i
End Function

Public Sub DrawAngleMark(ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal Col As Long, Optional ByVal Count As Long = 1, Optional ByVal DrawWidth As Long = 1, Optional ByVal Radius As Long = defAngleMarkRadius)
Const PiMinusToRad2 = (PI - ToRad) * 2 '6.2482787221397

Dim A0 As Double, A1 As Double, A2 As Double, Rad As Double
Dim ShouldSwap As Boolean
Dim hPen As Long, hOldPen As Long
Dim Z As Long

A1 = GetAngle(X2, Y2, X1, Y1)
A2 = GetAngle(X2, Y2, X3, Y3)
If Abs(A2 - A1) < 2 * ToRad Or Abs(A2 - A1) > PiMinusToRad2 Then Exit Sub
If A2 < A1 Then A2 = A2 + PI2
If PI - A2 + A1 <= 0 Then
    Swap X1, X3
    Swap Y1, Y3
End If

hPen = CreatePen(PS_SOLID, DrawWidth, Col)
hOldPen = SelectObject(hDC, hPen)

For Z = 1 To Count
    Rad = Radius + (Z - 1) * (DrawWidth + 2)
    Arc hDC, X2 - Rad - 1, Y2 - Rad - 1, X2 + Rad, Y2 + Rad, X1, Y1, X3, Y3
Next

SelectObject hDC, hOldPen
DeleteObject hPen
End Sub

Public Sub DrawRightAngleMark(ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal Col As Long, Optional ByVal Count As Long = 1, Optional ByVal DrawWidth As Long = 1, Optional ByVal Radius As Long = defAngleMarkRadius)
Dim hPen As Long, hOldPen As Long, lpPoint As POINTAPI
Dim A As Double, B As Double
Dim Xa As Double, Ya As Double, Xb As Double, Yb As Double
Dim P1 As OnePoint, P2 As OnePoint, P3 As OnePoint

A = Distance(X1, Y1, X2, Y2)
B = Distance(X3, Y3, X2, Y2)

Xa = (X3 - X2) * Radius / B
Ya = (Y3 - Y2) * Radius / B
Xb = (X1 - X2) * Radius / A
Yb = (Y1 - Y2) * Radius / A

P1.X = X2 + Xa
P1.Y = Y2 + Ya
P2.X = X2 + Xa + Xb
P2.Y = Y2 + Ya + Yb
P3.X = X2 + Xb
P3.Y = Y2 + Yb

hPen = CreatePen(PS_SOLID, DrawWidth, Col)
hOldPen = SelectObject(hDC, hPen)

MoveToEx hDC, P1.X, P1.Y, lpPoint
LineTo hDC, P2.X, P2.Y
LineTo hDC, P3.X, P3.Y

SelectObject hDC, hOldPen
DeleteObject hPen
End Sub

Public Function GetSystemColor(ByVal Col As Long) As Long
If Col < 0 Then GetSystemColor = GetSysColor(Col + SysColorTranslationBase) Else GetSystemColor = Col
End Function

Public Sub DrawSelectionShadow(ByVal hDC As Long, lpRect As RECT, Optional ByVal Col As Long = 0, Optional ByVal ShadowWidth As Long = 8, Optional ByVal TransparentShadow As Boolean = True)
Dim hPen As Long, hOldPen As Long

If setGradientFill Then
    ShadowRect hDC, lpRect, nPaperColor1, ShadowWidth, TransparentShadow
Else
    hPen = CreatePen(0, 1, Col)
    hOldPen = SelectObject(hDC, hPen)
    Rectangle hDC, lpRect.Left - 2, lpRect.Top - 2, lpRect.Right + 2, lpRect.Bottom + 3
    SelectObject hDC, hOldPen
    DeleteObject hPen
End If

End Sub
