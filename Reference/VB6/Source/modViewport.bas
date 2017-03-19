Attribute VB_Name = "modViewport"
'Contains procedures and functions specific to coordinate transformations,
'logical viewport parameters, etc.
Option Explicit

Public Const ScrollAmount = 0.389 ' how much to scroll when arrow keys are pressed
Public Const ZoomAmount = 1.1 ' how much to zoom when arrow keys are pressed

Public Sub ScrollLeft(Optional ByVal Large As Boolean = False)
ScrollPaper 1, 0, Large
End Sub

Public Sub ScrollRight(Optional ByVal Large As Boolean = False)
ScrollPaper -1, 0, Large
End Sub

Public Sub ScrollUp(Optional ByVal Large As Boolean = False)
ScrollPaper 0, 1, Large
End Sub

Public Sub ScrollDown(Optional ByVal Large As Boolean = False)
ScrollPaper 0, -1, Large
End Sub

Public Sub ZoomIn()
Dim TOX As Double, TOY As Double, OX As Double, OY As Double

If Abs(WorldTransform.XScalar) > 32 Then Exit Sub
PaperCls
TOX = WorldTransform.XOffset - WorldTransform.XScrOffset
TOY = WorldTransform.YOffset - WorldTransform.YScrOffset
OX = TOX * ZoomAmount - TOX
OY = TOY * ZoomAmount - TOY
ScrollCanvas OX, OY, ZoomAmount
ShowProperAll
i_Scrolled
End Sub

Public Sub ZoomOut()
Dim TOX As Double, TOY As Double, OX As Double, OY As Double

If Abs(WorldTransform.XScalar) <= 0.05 Then Exit Sub
PaperCls
TOX = WorldTransform.XOffset - WorldTransform.XScrOffset
TOY = WorldTransform.YOffset - WorldTransform.YScrOffset
OX = TOX / ZoomAmount - TOX
OY = TOY / ZoomAmount - TOY
ScrollCanvas OX, OY, 1 / ZoomAmount
ShowProperAll
i_Scrolled
End Sub

Public Sub ScrollHome()
If DragS.State = dscNormalState Then
    PaperCls
    LoadIdentity
    ScrollCanvas
    ShowProperAll
    i_Scrolled
End If
End Sub

Public Sub ScrollMouse(ByVal tX As Double, ByVal tY As Double)
tX = Paper.ScaleX(tX, vbPixels, defInternalScaleMode)
tY = Paper.ScaleY(tY, vbPixels, defInternalScaleMode)
PaperCls
ScrollCanvas tX, tY, , False
ShowAll
i_Scrolled tX <> 0, tY <> 0
End Sub

Public Sub LoadIdentity()
WorldTransform.XOffset = Paper.ScaleX(Paper.ScaleWidth / 2, vbPixels, defInternalScaleMode)
WorldTransform.YOffset = Paper.ScaleY(Paper.ScaleHeight / 2, vbPixels, defInternalScaleMode)
WorldTransform.XScrOffset = WorldTransform.XOffset
WorldTransform.YScrOffset = WorldTransform.YOffset
WorldTransform.XScalar = 1
WorldTransform.YScalar = -1
WorldTransform.XUnit = Paper.ScaleX(1, defInternalScaleMode, vbPixels)
WorldTransform.YUnit = Paper.ScaleY(1, defInternalScaleMode, vbPixels)
WorldTransform.Epsilon = Epsilon '/ WorldTransform.XScalar

WorldTransform.XCache = WorldTransform.XUnit * WorldTransform.XScalar
WorldTransform.YCache = WorldTransform.YUnit * WorldTransform.YScalar
WorldTransform.XCachedOffset = WorldTransform.XOffset * WorldTransform.XUnit
WorldTransform.YCachedOffset = WorldTransform.YOffset * WorldTransform.YUnit
WorldTransform.XDiv = WorldTransform.XOffset / WorldTransform.XScalar
WorldTransform.YDiv = WorldTransform.YOffset / WorldTransform.YScalar

Sensitivity = setCursorSensitivity
AngleMarkDist = defAngleMarkDist
ToLogicalLength Sensitivity
ToLogicalLength AngleMarkDist

RefreshCanvasBorders
FormMain.XRuler_Resize
FormMain.YRuler_Resize
End Sub

Public Sub ScrollCanvas(Optional ByVal XOffset As Double = 0, Optional ByVal YOffset As Double = 0, Optional ByVal ZoomFactor As Double = 1, Optional ByVal ShouldImitateMouseMove As Boolean = True)
On Local Error GoTo EH
Dim OldWorldTransform As TransformType, R As TwoPoints

OldWorldTransform = WorldTransform
WorldTransform.XOffset = WorldTransform.XOffset + XOffset
WorldTransform.YOffset = WorldTransform.YOffset + YOffset
WorldTransform.XScalar = WorldTransform.XScalar * ZoomFactor
WorldTransform.YScalar = WorldTransform.YScalar * ZoomFactor
WorldTransform.XCache = WorldTransform.XUnit * WorldTransform.XScalar
WorldTransform.YCache = WorldTransform.YUnit * WorldTransform.YScalar
WorldTransform.XCachedOffset = WorldTransform.XOffset * WorldTransform.XUnit
WorldTransform.YCachedOffset = WorldTransform.YOffset * WorldTransform.YUnit
WorldTransform.XDiv = WorldTransform.XOffset / WorldTransform.XScalar
WorldTransform.YDiv = WorldTransform.YOffset / WorldTransform.YScalar
WorldTransform.Epsilon = Epsilon

R.P1.X = 0 + FullScreenLineMargin
R.P1.Y = 0 + FullScreenLineMargin
ToLogical R.P1.X, R.P1.Y
R.P2.X = Paper.ScaleWidth - FullScreenLineMargin - 1
R.P2.Y = Paper.ScaleHeight - FullScreenLineMargin - 1
ToLogical R.P2.X, R.P2.Y

If Abs(R.P1.X) > MaxCoord Or Abs(R.P1.Y) > MaxCoord Or Abs(R.P2.X) > MaxCoord Or Abs(R.P2.Y) > MaxCoord Then
    WorldTransform = OldWorldTransform
    Exit Sub
End If

Sensitivity = setCursorSensitivity
AngleMarkDist = defAngleMarkDist
ToLogicalLength Sensitivity
ToLogicalLength AngleMarkDist

RefreshCanvasBorders
RecalcScrollAssociatedInfo
RecalcAllAuxInfo

If ShouldImitateMouseMove Then ImitateMouseMove
Exit Sub

EH:
End Sub

Public Sub ScrollPaper(ByVal X As Double, ByVal Y As Double, Optional ByVal Large As Boolean = False)
Const LargeScroll As Long = 10
PaperCls
ScrollCanvas X * ScrollAmount * IIf(Large, LargeScroll, 1), Y * ScrollAmount * IIf(Large, LargeScroll, 1)
ShowProperAll
i_Scrolled X <> 0, Y <> 0
End Sub

Public Sub ToLogical(ByRef X As Double, ByRef Y As Double)
X = X / WorldTransform.XCache - WorldTransform.XDiv
Y = Y / WorldTransform.YCache - WorldTransform.YDiv
End Sub

Public Sub ToPhysical(ByRef X As Double, ByRef Y As Double)
X = X * WorldTransform.XCache + WorldTransform.XCachedOffset
Y = Y * WorldTransform.YCache + WorldTransform.YCachedOffset
End Sub

Public Sub ToLogicalLength(ByRef D As Double)
D = D / WorldTransform.XCache
End Sub

Public Sub ToPhysicalLength(ByRef D As Double)
D = D * WorldTransform.XCache
End Sub

'================================================================================

Public Sub CalculateDefaultTransform(WorkArea As TwoPoints)
On Local Error Resume Next
Dim ZoomFactor As Double, X As Double, Y As Double
DefaultTransform.XOffset = Paper.ScaleX(Paper.ScaleWidth / 2, vbPixels, defInternalScaleMode)
DefaultTransform.YOffset = Paper.ScaleY(Paper.ScaleHeight / 2, vbPixels, defInternalScaleMode)

With WorkArea
    ZoomFactor = Minimum((CanvasBorders.P2.X - CanvasBorders.P1.X) / (.P2.X - .P1.X), (CanvasBorders.P2.Y - CanvasBorders.P1.Y) / (.P2.Y - .P1.Y))
    X = -(.P1.X + .P2.X) / 2 * ZoomFactor
    Y = (.P1.Y + .P2.Y) / 2 * ZoomFactor
End With

ScrollCanvas X, Y, ZoomFactor
FormMain.XRuler_Resize
FormMain.YRuler_Resize
End Sub

Public Sub RefreshCanvasBorders()
Dim X As Long, Y As Long, tX As Double, tY As Double, Q As Long
Dim CntX As Long, CntY As Long, SW As Long, SH As Long

PaperBorders.Left = 0 + FullScreenLineMargin
PaperBorders.Top = 0 + FullScreenLineMargin
PaperBorders.Right = Paper.ScaleWidth - FullScreenLineMargin - 1
PaperBorders.Bottom = Paper.ScaleHeight - FullScreenLineMargin - 1

CanvasBorders.P1.X = 0 + FullScreenLineMargin
CanvasBorders.P1.Y = 0 + FullScreenLineMargin
ToLogical CanvasBorders.P1.X, CanvasBorders.P1.Y
CanvasBorders.P2.X = Paper.ScaleWidth - FullScreenLineMargin - 1
CanvasBorders.P2.Y = Paper.ScaleHeight - FullScreenLineMargin - 1
ToLogical CanvasBorders.P2.X, CanvasBorders.P2.Y
If CanvasBorders.P2.X < CanvasBorders.P1.X Then Swap CanvasBorders.P1.X, CanvasBorders.P2.X
If CanvasBorders.P2.Y < CanvasBorders.P1.Y Then Swap CanvasBorders.P1.Y, CanvasBorders.P2.Y

SW = Paper.ScaleWidth
SH = Paper.ScaleHeight
PaperScaleWidth = SW
PaperScaleHeight = SH
ToPhysical tX, tY
OriginX = tX
OriginY = tY

CntX = Round(CanvasBorders.P2.X) - Round(CanvasBorders.P1.X) + 1
CntY = Round(CanvasBorders.P2.Y) - Round(CanvasBorders.P1.Y) + 1
ReDim Pts(1 To 2 * (CntX + CntY))
ReDim PtNums(1 To CntX + CntY)
UBoundPtNums = UBound(PtNums)

Q = Round(CanvasBorders.P1.X)
For X = Round(CanvasBorders.P1.X) To Round(CanvasBorders.P2.X)
    tX = X
    tY = 0
    ToPhysical tX, tY
    Pts((X - Q) * 2 + 1).X = tX
    Pts((X - Q) * 2 + 1).Y = 0
    Pts((X - Q) * 2 + 2).X = tX
    Pts((X - Q) * 2 + 2).Y = SH
    PtNums(X - Q + 1) = 2
Next

Q = Round(CanvasBorders.P1.Y)
For Y = Round(CanvasBorders.P1.Y) To Round(CanvasBorders.P2.Y)
    tX = 0
    tY = Y
    ToPhysical tX, tY
    Pts((Y - Q) * 2 + 1 + CntX * 2).X = 0
    Pts((Y - Q) * 2 + 1 + CntX * 2).Y = tY
    Pts((Y - Q) * 2 + 2 + CntX * 2).X = SW
    Pts((Y - Q) * 2 + 2 + CntX * 2).Y = tY
    PtNums(Y - Q + CntX + 1) = 2
Next

End Sub

Public Sub RecalcScrollAssociatedInfo()
RecalcPhysicalPoints
RecalcPointWidths
RecalcLocuses
RecalcLabels
RecalcButtons
End Sub

Public Sub RecalcPhysicalPoints()
Dim Z As Long
For Z = 1 To PointCount
    BasePoint(Z).PhysicalX = BasePoint(Z).X
    BasePoint(Z).PhysicalY = BasePoint(Z).Y
    ToPhysical BasePoint(Z).PhysicalX, BasePoint(Z).PhysicalY
Next
End Sub

Public Sub RecalcPointWidths()
Dim Z As Long

For Z = 1 To PointCount
    With BasePoint(Z)
        .Width = .PhysicalWidth
        ToLogicalLength .Width
    End With
Next
End Sub
