Attribute VB_Name = "modMenubar"
'========================================================
'Menu service module
'========================================================

Option Explicit

'========================================================
'Some API declarations required for TransparentBlt (see below)
'========================================================
Public Const SRCCOPY = &HCC0020
Public Const NOTSRCCOPY = &H330008
Public Const SRCAND = &H8800C6
Public Const SRCPAINT = &HEE0086

Public Type BITMAP
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type

Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function DPtoLP Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function GetMapMode Lib "gdi32" (ByVal hDC As Long) As Long

'========================================================

Public Enum MenuItemStyle
    TextOnly
    GraphicsOnly
    TextAndGraphics
End Enum

Public Type MenuItem
    AuxIndex As Variant
    BeginAGroup As Boolean
    Bounds As RECT
    Caption As String
    Checked As Boolean
    Enabled As Boolean
    Icon As StdPicture
    PopUp As Boolean
    ShortcutKey As Long
    Style As MenuItemStyle
    SubLevel As Integer
    ToolTipText As String
    Visible As Boolean
End Type

Public Type MenuType 'Global data type to hold menu information
    Items() As MenuItem
End Type

Public Menus() As MenuType 'global menu array
Public TotalMenuCount As Long 'number of open menu controls

'========================================================
'Adds a global array for a new menu
'========================================================

Public Function AddNewMenu()
TotalMenuCount = TotalMenuCount + 1
ReDim Preserve Menus(1 To TotalMenuCount)
ReDim Menus(TotalMenuCount).Items(1 To 1)
AddNewMenu = TotalMenuCount
End Function

'========================================================
'Removes an array when a menu is deleted
'========================================================

Public Sub RemoveMenu(ByVal Index As Long)
Dim Z As Long

If Index < 1 Or Index > TotalMenuCount Then Exit Sub

If Index < TotalMenuCount Then
    For Z = Index To TotalMenuCount - 1
        Menus(Z) = Menus(Z + 1)
    Next
End If

TotalMenuCount = TotalMenuCount - 1
If TotalMenuCount > 0 Then ReDim Preserve Menus(1 To TotalMenuCount)
End Sub

'============================================================
'Draws a specified bitmap on a specified hDC using specified transparent color
'Left-top corner has coordinates (XStart, YStart)
'============================================================

Public Sub TransparentBlt(ByVal hDC As Long, ByVal hBitmap As Long, ByVal XStart As Long, ByVal YStart As Long, ByVal cTransparentColor As Long, Optional ByVal DrawDisabled As Boolean = False)
Dim bm As BITMAP
Dim cColor As Long
Dim bmAndBack As Long, bmAndObject As Long, bmAndMem As Long, bmSave As Long
Dim bmBackOld As Long, bmObjectOld As Long, bmMemOld As Long, bmSaveOld As Long
Dim hdcMem As Long, hdcBack As Long, hdcObject As Long, hdcTemp As Long, hdcSave As Long
Dim ptSize As POINTAPI

'============================================================
'// 9 BitBlt calls;
'// 5 DCs declared;
'// 4 Bitmaps created;
'============================================================

hdcTemp = CreateCompatibleDC(hDC)
SelectObject hdcTemp, hBitmap   ' Select the bitmap

GetObjectAPI hBitmap, Len(bm), bm
ptSize.X = bm.bmWidth            ' Get width of bitmap
ptSize.Y = bm.bmHeight           ' Get height of bitmap
DPtoLP hdcTemp, ptSize, 1      ' Convert from device
                                              ' to logical points

'DrawState hdcTemp, 0, 0, hBitmap, 0, 0, 0, ptSize.X, ptSize.Y, DSS_DISABLED Or DST_BITMAP

'============================================================

'// Create some DCs to hold temporary data.
hdcBack = CreateCompatibleDC(hDC)
hdcObject = CreateCompatibleDC(hDC)
hdcMem = CreateCompatibleDC(hDC)
hdcSave = CreateCompatibleDC(hDC)

'// Create a bitmap for each DC. DCs are required for a number of
'// GDI functions.

'// Monochrome DC
bmAndBack = CreateBitmap(ptSize.X, ptSize.Y, 1, 1, 0)

'// Monochrome DC
bmAndObject = CreateBitmap(ptSize.X, ptSize.Y, 1, 1, 0)

bmAndMem = CreateCompatibleBitmap(hDC, ptSize.X, ptSize.Y)
bmSave = CreateCompatibleBitmap(hDC, ptSize.X, ptSize.Y)

'// Each DC must select a bitmap object to store pixel data.
bmBackOld = SelectObject(hdcBack, bmAndBack)
bmObjectOld = SelectObject(hdcObject, bmAndObject)
bmMemOld = SelectObject(hdcMem, bmAndMem)
bmSaveOld = SelectObject(hdcSave, bmSave)

'// Set proper mapping mode.
SetMapMode hdcTemp, GetMapMode(hDC)

'// Save the bitmap sent here, because it will be overwritten.
BitBlt hdcSave, 0, 0, ptSize.X, ptSize.Y, hdcTemp, 0, 0, SRCCOPY

'============================================================

'// Set the background color of the source DC to the color.
'// contained in the parts of the bitmap that should be transparent
cColor = SetBkColor(hdcTemp, cTransparentColor)

'// Create the object mask for the bitmap by performing a BitBlt
'// from the source bitmap to a monochrome bitmap.
BitBlt hdcObject, 0, 0, ptSize.X, ptSize.Y, hdcTemp, 0, 0, SRCCOPY

'// Set the background color of the source DC back to the original
'// color.
SetBkColor hdcTemp, cColor

'============================================================

'// Create the inverse of the object mask.
BitBlt hdcBack, 0, 0, ptSize.X, ptSize.Y, hdcObject, 0, 0, NOTSRCCOPY

'============================================================
'//Here we form the resulting bitmap!!!!!!!!!!!! (4 steps)

'1. //Copy the background of the main DC to the destination.
BitBlt hdcMem, 0, 0, ptSize.X, ptSize.Y, hDC, XStart, YStart, SRCCOPY

'2. // Mask out the places where the bitmap will be placed.
BitBlt hdcMem, 0, 0, ptSize.X, ptSize.Y, hdcObject, 0, 0, SRCAND

'3. // Mask out the transparent colored pixels on the bitmap.
BitBlt hdcTemp, 0, 0, ptSize.X, ptSize.Y, hdcBack, 0, 0, SRCAND

'4. //XOR the bitmap with the background on the destination DC.
BitBlt hdcMem, 0, 0, ptSize.X, ptSize.Y, hdcTemp, 0, 0, SRCPAINT

'============================================================

'// Copy the destination to the screen.
BitBlt hDC, XStart, YStart, ptSize.X, ptSize.Y, hdcMem, 0, 0, SRCCOPY

'DrawState hDC, 0, 0, hBitmap, 0, XStart, YStart, ptSize.X, ptSize.Y, DSS_DISABLED Or DST_BITMAP

'============================================================

'// Place the original bitmap back into the bitmap sent here.
BitBlt hdcTemp, 0, 0, ptSize.X, ptSize.Y, hdcSave, 0, 0, SRCCOPY

'============================================================

'// Delete the memory bitmaps.
DeleteObject SelectObject(hdcBack, bmBackOld)
DeleteObject SelectObject(hdcObject, bmObjectOld)
DeleteObject SelectObject(hdcMem, bmMemOld)
DeleteObject SelectObject(hdcSave, bmSaveOld)

'// Delete the memory DCs.
DeleteDC hdcMem
DeleteDC hdcBack
DeleteDC hdcObject
DeleteDC hdcSave
DeleteDC hdcTemp
End Sub

