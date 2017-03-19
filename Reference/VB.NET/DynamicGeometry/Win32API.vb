Friend Module Win32API

#Region " Constants "
	Public Const SRCCOPY As Int32 = &HCC0020
#End Region

#Region " Types "
	Public Structure RECT
		Public Left As Integer
		Public Top As Integer
		Public Right As Integer
		Public Bottom As Integer
	End Structure

	Public Structure POINTAPI
		Public x As Integer
		Public y As Integer
	End Structure
#End Region

#Region " Functions "
#Region " GDI "
	'#Region " Objects "
	Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Integer, ByVal hObject As Integer) As Integer
	Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Integer) As Integer
	'#End Region

	'#Region " DC "
	Public Declare Function GetDesktopWindow Lib "user32" () As Integer
	Public Declare Function GetDC Lib "user32" (ByVal hWnd As Integer) As Integer
	Public Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Integer) As Integer
	Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Integer) As Integer
	Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Integer) As Integer
	Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Integer, ByVal hDC As Integer) As Integer
	'#End Region

	'#Region " Pens "
	Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Integer, ByVal nWidth As Integer, ByVal crColor As Integer) As Integer
	'#End Region

	'#Region " Brushes "
	Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Integer) As Integer
	'#End Region

	'#Region " Bitmaps "
	Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal xSrc As Integer, ByVal ySrc As Integer, ByVal dwRop As Integer) As Integer
	Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer) As Integer
	'#End Region

	'#Region " Filled shapes "
	Public Declare Function FillRect Lib "user32" (ByVal hDC As Integer, ByRef lpRect As RECT, ByVal hBrush As Integer) As Integer
	'#End Region

	'#Region " Lines and curves "
	Public Declare Function LineTo Lib "gdi32" (ByVal hDC As Integer, ByVal X As Integer, ByVal Y As Integer) As Integer
	Public Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Integer, ByVal X As Integer, ByVal Y As Integer, ByRef lpPoint As POINTAPI) As Integer
	Public Declare Function Ellipse Lib "gdi32" (ByVal hDC As Integer, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer) As Integer
	Public Declare Function Rectangle Lib "gdi32" (ByVal hDC As Integer, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer) As Integer
	'#End Region

#End Region

#Region " Window functions "
	Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Int32) As Int32
#End Region

#Region " Time "
	Public Declare Function timeGetTime Lib "WinMM.dll" () As Integer
	Public Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef PC As Int64) As Integer
	Public Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef PC As Int64) As Integer
#End Region
#End Region

#Region " Macros "

#Region " GDI "
	Public Function RGB(ByVal R As Integer, ByVal G As Integer, ByVal B As Integer) As Integer
		RGB = R Or (G * 256) Or (B * 65536)
	End Function

	Public Function ToRECT(ByVal r As System.Drawing.Rectangle) As RECT
		With ToRECT
			.Left = r.Left
			.Top = r.Top
			.Right = r.Right
			.Bottom = r.Bottom
		End With
	End Function

	Public Sub FillRectangle(ByVal DC As Integer, ByVal Col As System.Drawing.Color, ByVal r As System.Drawing.Rectangle)
		Dim hBrush As Integer = CreateSolidBrush(System.Drawing.ColorTranslator.ToWin32(Col))
		FillRect(DC, ToRECT(r), hBrush)
		DeleteObject(hBrush)
	End Sub
#End Region

#Region " Time "
	Public Function Ticks() As Int64
		' TODO: SOMETIME: replace with timeGetTime on Win98...
		Dim t As Int64
		QueryPerformanceCounter(t)
		Return t
	End Function

	Public Function PerformanceCounterFrequency() As Int64
		Dim t As Int64
		QueryPerformanceFrequency(t)
		Return t
	End Function
#End Region

#Region " VB "

	Public Sub Swap(ByRef x1 As Double, ByRef x2 As Double)
		Dim t As Double = x1
		x1 = x2
		x2 = t
	End Sub

	Public Sub Order(ByRef x1 As Double, ByRef x2 As Double)
		If x1 > x2 Then
			Dim t As Double = x1
			x1 = x2
			x2 = t
		End If
	End Sub

#End Region

#End Region

End Module
