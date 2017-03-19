Attribute VB_Name = "Geometry"
Option Explicit

'===============================================
'Data structures:
'===============================================

Public Type TwoNumbers
    n1 As Double
    n2 As Double
End Type

Public Type FourNumbers
    Pair1 As TwoNumbers
    Pair2 As TwoNumbers
End Type

Public Type OnePoint
    X As Double
    Y As Double
End Type

Public Type TwoPoints
    P1 As OnePoint
    P2 As OnePoint
End Type

Public Type LineGeneralEquation
    a As Double
    b As Double
    c As Double
End Type

Public Type LineCanonicEquation
    X0 As Double
    Y0 As Double
    A1 As Double
    A2 As Double
End Type

Public Type LineNormalEquation
    Ang As Double
    D As Double
End Type

Public Type CircleEquation
    a As Double
    b As Double
    c As Double
End Type

Public Type TransformType
    XOffset As Double
    YOffset As Double
    XScalar As Double
    YScalar As Double
    XUnit As Double
    YUnit As Double
    XCache As Double
    YCache As Double
    XCachedOffset As Double
    YCachedOffset As Double
    XDiv As Double
    YDiv As Double
    XScrOffset As Double
    YScrOffset As Double
    Epsilon As Double
End Type

Public Type BasePointType
    Enabled As Boolean
    FillColor As Long
    FillStyle As FillStyleConstants
    ForeColor As Long
    Hide As Boolean
    LabelOffsetX As Long
    LabelOffsetY As Long
    LabelWidth As Long
    LabelHeight As Long
    LabelLength As Long
    Locus As Long
    Name As String
    NameColor As Long
    ParentFigure As Long
    PhysicalX As Double
    PhysicalY As Double
    PhysicalWidth As Double
    Selected As Boolean
    Shape As Integer
    Shown As Boolean
    ShowName As Boolean
    ShowCoordinates As Boolean
    Tag As Variant
    Type As Integer
    Visible As Boolean
    InDemo As Boolean
    Description As String
    Width As Double
    X As Double
    Y As Double
    ZOrder As Long
End Type

Public Const AuxCount = 6

Public Type Figure
    FigureType As DrawState
    Name As String
    NumberOfPoints As Long
    NumberOfChildren As Long
    Hide As Boolean
    Points() As Long
    Parents() As Long
    Children() As Long
    DrawMode As Integer
    DrawStyle As Integer
    DrawWidth As Integer
    ForeColor As Long
    FillColor As Long
    FillStyle As Long
    Selected As Boolean
    Visible As Boolean
    InDemo As Boolean
    Description As String
    Tag As Long
    XTree As Tree
    YTree As Tree
    XS As String
    YS As String
    AlreadyHidden As Boolean
    AlreadyShown As Boolean
    AuxPoints(1 To AuxCount) As OnePoint
    AuxInfo(1 To AuxCount) As Double
    AuxArray() As Long
    ZOrder As Long
End Type

'Locus data type; keeps all locus-related properties
Public Type Locus
    DrawWidth As Integer
    Dynamic As Boolean
    Enabled As Boolean
    ForeColor As Long
    LocusPoints() As OnePoint
    LocusPixels() As POINTAPI
    LocusNumbers() As Long 'number of points in each piece
    LocusNumber As Long ' number of pieces
    LocusPointCount As Long 'total number of points
    ParentPoint As Long 'point that describes this locus
    ParentFigure As Long 'if DynamicLocus Then identifies dsDynamicLocus figure
    Type As Long
    Visible As Boolean
    Hide As Boolean
    InDemo As Boolean
    Description As String
    ShouldBreak As Boolean 'whether to begin a new locus piece
End Type

Public Type DynamicLabelUnit
    Type As DynamicLabelType
    DynamicTree As Tree
    StaticText As String
End Type

Public Type ObjectList
    PointCount As Long
    FigureCount As Long
    LabelCount As Long
    SGCount As Long
    LocusCount As Long
    WECount As Long
    ButtonCount As Long
    TotalCount As Long
    PointCountMax As Long
    FigureCountMax As Long
    LabelCountMax As Long
    SGCountMax As Long
    LocusCountMax As Long
    WECountMax As Long
    ButtonCountMax As Long
    TotalCountMax As Long
    Points() As Long
    Figures() As Long
    Labels() As Long
    SGs() As Long
    Loci() As Long
    WEs() As Long
    Buttons() As Long
    BoundingRect As TwoPoints
    Type As ObjectSelectionType
    SubType As ObjectSelectionCaller
End Type

Public Type LinearObjectListItem
    Type As GeometryObjectType
    index As Long
    Participate As Boolean
    Description As String
End Type

Public Type LinearObjectList
    Items() As LinearObjectListItem
    Count As Long
    Current As Long
    Step As Long
    StepCount As Long
End Type

Public Type Button
    Appearance As Long
    Caption As String
    Type As ButtonType
    Borders As TextBorderStyle
    Pushed As Boolean
    CurrentState As Long
    Position As RECT
    LogicalPosition As TwoPoints
    Charset As Long
    ForeColor As Long
    BackColor As Long
    FontName As String
    FontSize As Integer
    FontBold As Boolean
    FontItalic As Boolean
    FontUnderline As Boolean
    Hide As Boolean
    Visible As Boolean
    ObjectListAux As ObjectList
    InitiallyVisible As Boolean
    Message As String
    Path As String
    RemindToSaveFile As Boolean
    Fixed As Boolean
    InDemo As Boolean
    Description As String
End Type

Public Type TextLabel
    Borders As TextBorderStyle
    Caption As String
    CompiledCaptionParts() As DynamicLabelUnit
    CCCP As Long 'it's just CountofCompiledCaptionParts
    Dynamic As Boolean
    DisplayName As String
    LenDisplayName As Long
    Charset As Long
    Position As RECT
    LogicalPosition As TwoPoints
    ForeColor As Long
    BackColor As Long
    FontName As String
    FontSize As Integer
    FontBold As Boolean
    FontItalic As Boolean
    FontUnderline As Boolean
    Transparent As Boolean
    Shadow As Boolean
    Frame As Boolean
    Visible As Boolean
    Hide As Boolean
    Fixed As Boolean
    InDemo As Boolean
    Description As String
End Type

Public Type StaticGraphic
    ForeColor As Long
    FillColor As Long
    FillStyle As Long
    DrawMode As Long
    DrawStyle As Long
    DrawWidth As Long
    Hide As Boolean
    Visible As Boolean
    NumberOfPoints As Long
    Points() As Long
    ObjectPoints() As OnePoint
    ObjectPixels() As POINTAPI
    Type As StaticGraphicType
    Tag As Long
    InDemo As Boolean
    Description As String
End Type

Public Type MacroGiven
    Type As DrawState
    Description As String
    Order As Long
End Type

Public Type Macro
    Givens() As MacroGiven
    GivenCount As Long
    Results() As Figure
    ResultCount As Long
    FigurePoints() As BasePointType
    FigurePointCount As Long
    SG() As StaticGraphic
    SGCount As Long
    Name As String
    Description As String
End Type

Public Type WatchExpression
    Name As String
    Expression As String
    WatchTree As Tree
    Value As Double
    ParentPoints() As Long
    ParentWEs() As Long
End Type

'=================================================

Public Function AddBasePoint(ByVal X As Double, ByVal Y As Double, Optional ByVal PointName As String = "", Optional ByVal pType As Long = dsPoint, Optional ByVal Display As Boolean = True, Optional ByVal ParentFigure As Long = -1, Optional ByVal ShouldRecord As Boolean = True, Optional ByVal ShouldSetDescription As Boolean = True, Optional ByVal ShouldRefresh As Boolean = True) As Boolean
' Adds a point record to the Basepoint() array and makes the point usable
'=====================================================
Dim MaxZOrder As Long
Dim lpSize As Size
Dim pAction As Action

If PointCount >= MaxPointCount Then Exit Function

' Prepare a point name
If PointName = "" Then PointName = GenerateNewPointName(True)
'PointNames.Add PointName, PointName ' add a name into global name collection

MaxZOrder = GenerateNewPointZOrder

'===================================
' Now resizing the array and adding the structure
'===================================
PointCount = PointCount + 1
RedimPreserveBasePoint 1, PointCount

'=======================================
' Initialize a point record with the properties of the point
'=======================================
BasePoint(PointCount).Name = PointName
BasePoint(PointCount).Type = pType
BasePoint(PointCount).Enabled = True
BasePoint(PointCount).Visible = True
BasePoint(PointCount).ZOrder = MaxZOrder

If pType = dsPoint Then
    BasePoint(PointCount).ForeColor = setdefcolPoint
    BasePoint(PointCount).FillColor = setdefcolPointFill
    BasePoint(PointCount).Shape = setdefPointShape
ElseIf pType = dsPointOnFigure Then
    BasePoint(PointCount).ForeColor = setdefcolFigurePoint
    BasePoint(PointCount).FillColor = setdefcolFigurePointFill
    BasePoint(PointCount).Shape = setdefFigurePointShape
    If ParentFigure <> -1 Then BasePoint(PointCount).ParentFigure = ParentFigure
Else
    BasePoint(PointCount).ForeColor = setdefcolDependentPoint
    BasePoint(PointCount).FillColor = setdefcolDependentPointFill
    BasePoint(PointCount).Shape = setdefDependentPointShape
    If ParentFigure <> -1 Then BasePoint(PointCount).ParentFigure = ParentFigure
End If

BasePoint(PointCount).FillStyle = setdefPointFill
BasePoint(PointCount).ShowName = setAutoShowPointName
BasePoint(PointCount).NameColor = setdefcolPointName

BasePoint(PointCount).LabelOffsetX = 0
BasePoint(PointCount).LabelOffsetY = -setdefPointSize \ 2 + 1
BasePoint(PointCount).LabelWidth = Paper.TextWidth(PointName)
BasePoint(PointCount).LabelHeight = Paper.TextHeight(PointName)
BasePoint(PointCount).LabelLength = Len(PointName)

BasePoint(PointCount).PhysicalWidth = setdefPointSize
BasePoint(PointCount).Width = setdefPointSize
ToLogicalLength BasePoint(PointCount).Width

If ShouldSetDescription Then BasePoint(PointCount).Description = GetObjectDescription(gotPoint, PointCount)
BasePoint(PointCount).InDemo = True
BasePoint(PointCount).Locus = 0

'================================================
' Now display the point and finalize the procedure
'================================================

MovePoint PointCount, X, Y ' initialize coordinates

If Display Then ShowPoint Paper.hDC, PointCount, , ShouldRefresh ' DRAW!!!!

'================================================
' ...and finally notify the Undo engine of what has happened
' (a new point has been born!)
'================================================
pAction.Type = actAddPoint
pAction.pPoint = PointCount
If ShouldRecord Then RecordAction pAction

UpdateLabels

AddBasePoint = True
End Function

Public Sub MovePoint(ByVal Point1 As Long, ByVal X As Double, ByVal Y As Double)
Dim sStr As String, lpSize As Size

If Abs(X) > 1000 Or Abs(Y) > 1000 Then BasePoint(Point1).Visible = False: Exit Sub
'x=emptyvar or y=emptyvar

X = Round(X, NumDecimalDigits)
Y = Round(Y, NumDecimalDigits)

If BasePoint(Point1).Visible And BasePoint(Point1).Locus > 0 Then
    If Locuses(BasePoint(Point1).Locus).Enabled And Not Locuses(BasePoint(Point1).Locus).Dynamic Then AddPointToLocus BasePoint(Point1).Locus, X, Y
End If

BasePoint(Point1).X = X
BasePoint(Point1).Y = Y

BasePoint(Point1).PhysicalX = X
BasePoint(Point1).PhysicalY = Y
ToPhysical BasePoint(Point1).PhysicalX, BasePoint(Point1).PhysicalY

If BasePoint(Point1).ShowName And BasePoint(Point1).ShowCoordinates Then
    sStr = BasePoint(Point1).Name
    If BasePoint(Point1).ShowCoordinates Then sStr = sStr & " (" & Format(BasePoint(Point1).X, setFormatNumber) & "; " & Format(BasePoint(Point1).Y, setFormatNumber) & ")"
    BasePoint(Point1).LabelLength = Len(sStr)
    BasePoint(Point1).LabelWidth = Paper.TextWidth(sStr)
    BasePoint(Point1).LabelHeight = Paper.TextHeight(sStr)
End If
End Sub

Public Sub MovePointPure(ByVal Point1 As Long, ByVal X As Double, ByVal Y As Double)
If Abs(X) > 1000 Or Abs(Y) > 1000 Then BasePoint(Point1).Visible = False: Exit Sub

BasePoint(Point1).X = Round(X, NumDecimalDigits)
BasePoint(Point1).Y = Round(Y, NumDecimalDigits)
End Sub

Public Function AddFigureAux(ByVal FigureNum As Long, ByVal FigureType As DrawState, Optional ByVal DefaultName As String = "", Optional ByVal ShouldRecord As Boolean = True) As Boolean
Dim pAction As Action

If FigureNum > MaxFigureCount Then Exit Function

Figures(FigureNum).FigureType = FigureType
FormMain.mnuFigureList.Enabled = True
If DefaultName = "" Then DefaultName = GenerateNewFigureName(FigureType)
Figures(FigureNum).Name = DefaultName
'FigureNames.Add DefaultName

Figures(FigureCount).DrawMode = defFigureDrawMode
Figures(FigureCount).DrawStyle = defFigureDrawStyle
Figures(FigureCount).DrawWidth = setdefFigureDrawWidth
Figures(FigureCount).ForeColor = setdefcolFigure
Figures(FigureCount).FillColor = setdefcolFigureFill
Figures(FigureCount).FillStyle = 6
Figures(FigureCount).Hide = False
Figures(FigureCount).Visible = True
Figures(FigureCount).ZOrder = GenerateNewFigureZOrder
Figures(FigureCount).InDemo = True
FigureVisualOrder.Add Val(FigureCount)

pAction.Type = actAddFigure
pAction.pFigure = FigureCount
If ShouldRecord Then RecordAction pAction

AddFigureAux = True
End Function

Public Sub AddSegment(ByVal Point1 As Long, ByVal Point2 As Long, Optional ByVal SegmentName As String = "", Optional ByVal ShouldRecord As Boolean = True)
RedimPreserveFigures 0, FigureCount
If Not AddFigureAux(FigureCount, dsSegment, SegmentName, ShouldRecord) Then Exit Sub

Figures(FigureCount).NumberOfPoints = 2
ReDim Figures(FigureCount).Points(0 To 1)
Figures(FigureCount).Points(0) = Point1
Figures(FigureCount).Points(1) = Point2

Figures(FigureCount).Description = GetObjectDescription(gotFigure, FigureCount)

FigureCount = FigureCount + 1
RecalcAuxInfo FigureCount - 1
End Sub

Public Sub AddLine(ByVal Point1 As Long, ByVal Point2 As Long, Optional ByVal LineName As String = "")
RedimPreserveFigures 0, FigureCount
If Not AddFigureAux(FigureCount, dsLine_2Points, LineName) Then Exit Sub

Figures(FigureCount).NumberOfPoints = 2
ReDim Figures(FigureCount).Points(0 To 1)
Figures(FigureCount).NumberOfChildren = 0
Figures(FigureCount).Points(0) = Point1
Figures(FigureCount).Points(1) = Point2

Figures(FigureCount).Description = GetObjectDescription(gotFigure, FigureCount)

FigureCount = FigureCount + 1
RecalcAuxInfo FigureCount - 1
End Sub

Public Sub AddBisector(ByVal Point1 As Long, ByVal Point2 As Long, ByVal Point3 As Long, Optional ByVal LineName As String = "")
RedimPreserveFigures 0, FigureCount
If Not AddFigureAux(FigureCount, dsBisector, LineName) Then Exit Sub

Figures(FigureCount).NumberOfPoints = 3
ReDim Figures(FigureCount).Points(0 To 2)
Figures(FigureCount).NumberOfChildren = 0
Figures(FigureCount).Points(0) = Point1
Figures(FigureCount).Points(1) = Point2
Figures(FigureCount).Points(2) = Point3

Figures(FigureCount).Description = GetObjectDescription(gotFigure, FigureCount)

FigureCount = FigureCount + 1
RecalcAuxInfo FigureCount - 1
End Sub

Public Sub AddAnLineGeneral(ByVal a As Double, ByVal b As Double, ByVal c As Double)
Dim R As TwoPoints
RedimPreserveFigures 0, FigureCount
If Not AddFigureAux(FigureCount, dsAnLineGeneral) Then Exit Sub

Figures(FigureCount).NumberOfPoints = 0
ReDim Figures(FigureCount).Points(0 To 0)
Figures(FigureCount).NumberOfChildren = 0
Figures(FigureCount).AuxInfo(1) = a
Figures(FigureCount).AuxInfo(2) = b
Figures(FigureCount).AuxInfo(3) = c
Figures(FigureCount).AuxInfo(6) = 1 ' type of anline (0 to 4)
R = GetGeneralAnLineAbsolute(a, b, c)
Figures(FigureCount).AuxPoints(3) = R.P1
Figures(FigureCount).AuxPoints(4) = R.P2

Figures(FigureCount).Description = GetObjectDescription(gotFigure, FigureCount)

FigureCount = FigureCount + 1
RecalcAuxInfo FigureCount - 1
End Sub

Public Sub EditAnLineGeneral(ByVal a As Double, ByVal b As Double, ByVal c As Double)
Dim R As TwoPoints

Figures(ActiveFigure).FigureType = dsAnLineGeneral

Figures(ActiveFigure).AuxInfo(1) = a
Figures(ActiveFigure).AuxInfo(2) = b
Figures(ActiveFigure).AuxInfo(3) = c
Figures(ActiveFigure).AuxInfo(6) = 1 ' type of anline (0 to 4)
R = GetGeneralAnLineAbsolute(a, b, c)
Figures(ActiveFigure).AuxPoints(3) = R.P1
Figures(ActiveFigure).AuxPoints(4) = R.P2

RecalcAllAuxInfo
End Sub

Public Sub AddActiveAxes()
If nActiveAxesAdded Then Exit Sub
nActiveAxesAdded = True
FormMain.mnuActiveAxes.Enabled = False
AddAnLineGeneral 0, 1, 0
nActiveY = FigureCount - 1
AddAnLineGeneral 1, 0, 0
nActiveX = FigureCount - 1

Figures(FigureCount - 2).ForeColor = nAxesColor
Figures(FigureCount - 2).Name = GetString(ResXAxis)
Figures(FigureCount - 2).DrawMode = 13
Figures(FigureCount - 2).DrawWidth = 1
'FigureNames.Add GetString(ResXAxis)

Figures(FigureCount - 1).ForeColor = nAxesColor
Figures(FigureCount - 1).Name = GetString(ResYAxis)
Figures(FigureCount - 1).DrawMode = 13
Figures(FigureCount - 1).DrawWidth = 1
'FigureNames.Add GetString(ResYAxis)

If Not nShowAxes Then
    nShowAxes = True
    'setShowAxes = True
    'SaveSetting AppName, "Interface", "ShowAxes", Format(-CInt(setShowAxes))
    FormMain.mnuShowAxes.Checked = True
End If

ShowAll
End Sub

Public Sub AddAnLineCanonic(ByVal X0 As Double, ByVal Y0 As Double, ByVal A1 As Double, ByVal A2 As Double)
Dim R As TwoPoints
RedimPreserveFigures 0, FigureCount
If Not AddFigureAux(FigureCount, dsAnLineCanonic) Then Exit Sub

Figures(FigureCount).NumberOfPoints = 0
ReDim Figures(FigureCount).Points(0 To 0)
Figures(FigureCount).NumberOfChildren = 0
Figures(FigureCount).AuxInfo(1) = X0
Figures(FigureCount).AuxInfo(2) = Y0
Figures(FigureCount).AuxInfo(3) = A1
Figures(FigureCount).AuxInfo(4) = A2
Figures(FigureCount).AuxInfo(6) = 2
R = GetCanonicAnLineAbsolute(X0, Y0, A1, A2)
Figures(FigureCount).AuxPoints(3) = R.P1
Figures(FigureCount).AuxPoints(4) = R.P2

Figures(FigureCount).Description = GetObjectDescription(gotFigure, FigureCount)

FigureCount = FigureCount + 1
RecalcAuxInfo FigureCount - 1
End Sub

Public Sub EditAnLineCanonic(ByVal X0 As Double, ByVal Y0 As Double, ByVal A1 As Double, ByVal A2 As Double)
Dim R As TwoPoints

Figures(ActiveFigure).FigureType = dsAnLineCanonic

Figures(ActiveFigure).AuxInfo(1) = X0
Figures(ActiveFigure).AuxInfo(2) = Y0
Figures(ActiveFigure).AuxInfo(3) = A1
Figures(ActiveFigure).AuxInfo(4) = A2
Figures(ActiveFigure).AuxInfo(6) = 2
R = GetCanonicAnLineAbsolute(X0, Y0, A1, A2)
Figures(ActiveFigure).AuxPoints(3) = R.P1
Figures(ActiveFigure).AuxPoints(4) = R.P2

RecalcAllAuxInfo
End Sub

Public Sub AddAnLineNormal(ByVal a As Double, ByVal D As Double)
Dim R As TwoPoints
RedimPreserveFigures 0, FigureCount
If Not AddFigureAux(FigureCount, dsAnLineNormal) Then Exit Sub

Figures(FigureCount).NumberOfPoints = 0
ReDim Figures(FigureCount).Points(0 To 0)
Figures(FigureCount).NumberOfChildren = 0
Figures(FigureCount).AuxInfo(1) = a
Figures(FigureCount).AuxInfo(2) = D
Figures(FigureCount).AuxInfo(6) = 3
R = GetNormalAnLineAbsolute(a, D)
Figures(FigureCount).AuxPoints(3) = R.P1
Figures(FigureCount).AuxPoints(4) = R.P2

Figures(FigureCount).Description = GetObjectDescription(gotFigure, FigureCount)

FigureCount = FigureCount + 1
RecalcAuxInfo FigureCount - 1
End Sub

Public Sub EditAnLineNormal(ByVal a As Double, ByVal D As Double)
Dim R As TwoPoints
Figures(ActiveFigure).FigureType = dsAnLineNormal

Figures(ActiveFigure).AuxInfo(1) = a
Figures(ActiveFigure).AuxInfo(2) = D
Figures(ActiveFigure).AuxInfo(6) = 3
R = GetNormalAnLineAbsolute(a, D)
Figures(ActiveFigure).AuxPoints(3) = R.P1
Figures(ActiveFigure).AuxPoints(4) = R.P2

RecalcAllAuxInfo
End Sub

Public Sub AddAnLineNormalPoint(ByVal X0 As Double, ByVal Y0 As Double, ByVal A1 As Double, ByVal A2 As Double)
Dim R As TwoPoints
RedimPreserveFigures 0, FigureCount
If Not AddFigureAux(FigureCount, dsAnLineNormalPoint) Then Exit Sub

Figures(FigureCount).NumberOfPoints = 0
ReDim Figures(FigureCount).Points(0 To 0)
Figures(FigureCount).NumberOfChildren = 0
Figures(FigureCount).AuxInfo(1) = X0
Figures(FigureCount).AuxInfo(2) = Y0
Figures(FigureCount).AuxInfo(3) = A1
Figures(FigureCount).AuxInfo(4) = A2
Figures(FigureCount).AuxInfo(6) = 4
R = GetNormalPointAnLineAbsolute(X0, Y0, A1, A2)
Figures(FigureCount).AuxPoints(3) = R.P1
Figures(FigureCount).AuxPoints(4) = R.P2

Figures(FigureCount).Description = GetObjectDescription(gotFigure, FigureCount)

FigureCount = FigureCount + 1
RecalcAuxInfo FigureCount - 1
End Sub

Public Sub EditAnLineNormalPoint(ByVal X0 As Double, ByVal Y0 As Double, ByVal A1 As Double, ByVal A2 As Double)
Dim R As TwoPoints
Figures(ActiveFigure).FigureType = dsAnLineNormalPoint

Figures(ActiveFigure).AuxInfo(1) = X0
Figures(ActiveFigure).AuxInfo(2) = Y0
Figures(ActiveFigure).AuxInfo(3) = A1
Figures(ActiveFigure).AuxInfo(4) = A2
Figures(ActiveFigure).AuxInfo(6) = 4
R = GetNormalPointAnLineAbsolute(X0, Y0, A1, A2)
Figures(ActiveFigure).AuxPoints(3) = R.P1
Figures(ActiveFigure).AuxPoints(4) = R.P2

RecalcAllAuxInfo
End Sub

Public Sub AddAnCircle(ByVal a As Double, ByVal b As Double, ByVal c As Double)
RedimPreserveFigures 0, FigureCount
If Not AddFigureAux(FigureCount, dsAnCircle) Then Exit Sub

Figures(FigureCount).NumberOfPoints = 0
ReDim Figures(FigureCount).Points(0 To 0)
Figures(FigureCount).NumberOfChildren = 0
Figures(FigureCount).AuxInfo(1) = a
Figures(FigureCount).AuxInfo(2) = b
Figures(FigureCount).AuxInfo(3) = c

Figures(FigureCount).Description = GetObjectDescription(gotFigure, FigureCount)

FigureCount = FigureCount + 1
RecalcAuxInfo FigureCount - 1
End Sub

Public Sub EditAnCircle(ByVal a As Double, ByVal b As Double, ByVal c As Double)
Figures(ActiveFigure).FigureType = dsAnCircle

Figures(ActiveFigure).AuxInfo(1) = a
Figures(ActiveFigure).AuxInfo(2) = b
Figures(ActiveFigure).AuxInfo(3) = c

RecalcAllAuxInfo
End Sub

Public Sub AddRay(ByVal Point1 As Long, ByVal Point2 As Long, Optional ByVal RayName As String = "")
RedimPreserveFigures 0, FigureCount
If Not AddFigureAux(FigureCount, dsRay, RayName) Then Exit Sub

Figures(FigureCount).NumberOfPoints = 2
ReDim Figures(FigureCount).Points(0 To 1)
Figures(FigureCount).NumberOfChildren = 0
Figures(FigureCount).Points(0) = Point1
Figures(FigureCount).Points(1) = Point2

Figures(FigureCount).Description = GetObjectDescription(gotFigure, FigureCount)

FigureCount = FigureCount + 1
RecalcAuxInfo FigureCount - 1
End Sub

Public Sub AddParallelLine(ByVal Point1 As Long, ByVal Line1 As Long, Optional ByVal LineName As String = "")
RedimPreserveFigures 0, FigureCount
If Not AddFigureAux(FigureCount, dsLine_PointAndParallelLine, LineName) Then Exit Sub

Figures(FigureCount).NumberOfPoints = 1
ReDim Figures(FigureCount).Points(0 To 0)
ReDim Figures(FigureCount).Parents(0 To 0)
Figures(FigureCount).NumberOfChildren = 0
Figures(FigureCount).Points(0) = Point1
Figures(FigureCount).Parents(0) = Line1

ReDim Preserve Figures(Line1).Children(0 To Figures(Line1).NumberOfChildren)
Figures(Line1).Children(Figures(Line1).NumberOfChildren) = FigureCount
Figures(Line1).NumberOfChildren = Figures(Line1).NumberOfChildren + 1

Figures(FigureCount).Description = GetObjectDescription(gotFigure, FigureCount)

FigureCount = FigureCount + 1
RecalcAuxInfo FigureCount - 1
End Sub

Public Sub AddPerpendicularLine(ByVal Point1 As Long, ByVal Line1 As Long, Optional ByVal LineName As String = "")
RedimPreserveFigures 0, FigureCount
If Not AddFigureAux(FigureCount, dsLine_PointAndPerpendicularLine, LineName) Then Exit Sub

Figures(FigureCount).NumberOfPoints = 1
ReDim Figures(FigureCount).Points(0 To 0)
ReDim Figures(FigureCount).Parents(0 To 0)
Figures(FigureCount).Points(0) = Point1
Figures(FigureCount).Parents(0) = Line1

ReDim Preserve Figures(Line1).Children(0 To Figures(Line1).NumberOfChildren)
Figures(Line1).Children(Figures(Line1).NumberOfChildren) = FigureCount
Figures(Line1).NumberOfChildren = Figures(Line1).NumberOfChildren + 1

Figures(FigureCount).Description = GetObjectDescription(gotFigure, FigureCount)

FigureCount = FigureCount + 1
RecalcAuxInfo FigureCount - 1
End Sub

Public Sub AddCircle(ByVal Point1 As Long, ByVal Point2 As Long, Optional ByVal CircleName As String = "")
RedimPreserveFigures 0, FigureCount
If Not AddFigureAux(FigureCount, dsCircle_CenterAndCircumPoint, CircleName) Then Exit Sub

Figures(FigureCount).NumberOfPoints = 2
ReDim Figures(FigureCount).Points(0 To 1)
Figures(FigureCount).Points(0) = Point1
Figures(FigureCount).Points(1) = Point2

Figures(FigureCount).Description = GetObjectDescription(gotFigure, FigureCount)

FigureCount = FigureCount + 1
RecalcAuxInfo FigureCount - 1
End Sub

Public Sub AddCircleByRadius(ByVal Point1 As Long, ByVal Point2 As Long, ByVal Point3 As Long, Optional ByVal CircleName As String = "")
RedimPreserveFigures 0, FigureCount
If Not AddFigureAux(FigureCount, dsCircle_CenterAndTwoPoints, CircleName) Then Exit Sub

Figures(FigureCount).NumberOfPoints = 3
ReDim Figures(FigureCount).Points(0 To 2)
Figures(FigureCount).Points(0) = Point1
Figures(FigureCount).Points(1) = Point2
Figures(FigureCount).Points(2) = Point3

Figures(FigureCount).Description = GetObjectDescription(gotFigure, FigureCount)

FigureCount = FigureCount + 1
RecalcAuxInfo FigureCount - 1
End Sub

Public Sub AddArc(ByVal Point1 As Long, ByVal Point2 As Long, ByVal Point3 As Long, ByVal Point4 As Long, ByVal Point5 As Long, Optional ByVal ArcName As String = "")
Dim ArcCount As Long, Point6 As Integer, Point7 As Integer
Dim X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, X4 As Double, Y4 As Double, X5 As Double, Y5 As Double, X As Double, Y As Double
Dim A1 As Double, A2 As Double, Rad As Single
RedimPreserveFigures 0, FigureCount
If Not AddFigureAux(FigureCount, dsCircle_ArcCenterAndRadiusAndTwoPoints, ArcName) Then Exit Sub

X1 = BasePoint(Point1).X
Y1 = BasePoint(Point1).Y
X2 = BasePoint(Point2).X
Y2 = BasePoint(Point2).Y
X3 = BasePoint(Point3).X
Y3 = BasePoint(Point3).Y
X4 = BasePoint(Point4).X
Y4 = BasePoint(Point4).Y
X5 = BasePoint(Point5).X
Y5 = BasePoint(Point5).Y
Rad = Distance(X1, Y1, X2, Y2)
A1 = GetAngle(X3, Y3, X4, Y4)
A2 = GetAngle(X3, Y3, X5, Y5)
X = X3 + Rad * Cos(A1)
Y = Y3 - Rad * Sin(A1)
AddBasePoint X, Y, , dsCircle_ArcCenterAndRadiusAndTwoPoints, False, FigureCount, False, False
Point6 = PointCount
X = X3 + Rad * Cos(A2)
Y = Y3 - Rad * Sin(A2)
AddBasePoint X, Y, , dsCircle_ArcCenterAndRadiusAndTwoPoints, False, FigureCount, False, False
Point7 = PointCount

Figures(FigureCount).NumberOfPoints = 7
ReDim Figures(FigureCount).Points(0 To 6)
Figures(FigureCount).Points(0) = Point1
Figures(FigureCount).Points(1) = Point2
Figures(FigureCount).Points(2) = Point3
Figures(FigureCount).Points(3) = Point4
Figures(FigureCount).Points(4) = Point5
Figures(FigureCount).Points(5) = Point6
Figures(FigureCount).Points(6) = Point7

Figures(FigureCount).Description = GetObjectDescription(gotFigure, FigureCount)
BasePoint(PointCount - 1).Description = GetObjectDescription(gotPoint, PointCount - 1)
BasePoint(PointCount).Description = GetObjectDescription(gotPoint, PointCount)

FigureCount = FigureCount + 1
RecalcAuxInfo FigureCount - 1
End Sub

Public Sub AddMiddlePoint(ByVal Point1 As Long, ByVal Point2 As Long, Optional ByVal MiddlePointName As String = "")
RedimPreserveFigures 0, FigureCount

Figures(FigureCount).NumberOfPoints = 3
If Not AddFigureAux(FigureCount, dsMiddlePoint, MiddlePointName) Then Exit Sub
AddBasePoint (BasePoint(Point1).X + BasePoint(Point2).X) / 2, (BasePoint(Point1).Y + BasePoint(Point2).Y) / 2, , dsMiddlePoint, False, FigureCount, False, False

ReDim Figures(FigureCount).Points(0 To 2)
Figures(FigureCount).Points(0) = PointCount
Figures(FigureCount).Points(1) = Point1
Figures(FigureCount).Points(2) = Point2
'BasePoint(Figures(FigureCount).Points(0)).DrawWidth = defFigureDrawWidth
'BasePoint(Figures(FigureCount).Points(0)).ForeColor = setdefcolDependentPoint

Figures(FigureCount).Description = GetObjectDescription(gotFigure, FigureCount)
BasePoint(PointCount).Description = GetObjectDescription(gotPoint, PointCount)

FigureCount = FigureCount + 1
RecalcAuxInfo FigureCount - 1
End Sub

Public Sub AddIntersectionPoints(ByVal Figure1 As Long, ByVal Figure2 As Long, Optional ByVal IPointName As String = "")
Dim R As TwoPoints
RedimPreserveFigures 0, FigureCount
If Not AddFigureAux(FigureCount, dsIntersect, IPointName) Then Exit Sub

Figures(FigureCount).NumberOfPoints = 2
R = GetIntersectionPoints(Figure1, Figure2)
ReDim Figures(FigureCount).Points(0 To 1)
ReDim Figures(FigureCount).Parents(0 To 1)
Figures(FigureCount).Parents(0) = Figure1
Figures(FigureCount).Parents(1) = Figure2
Figures(FigureCount).NumberOfChildren = 0

ReDim Preserve Figures(Figure1).Children(0 To Figures(Figure1).NumberOfChildren)
Figures(Figure1).Children(Figures(Figure1).NumberOfChildren) = FigureCount
Figures(Figure1).NumberOfChildren = Figures(Figure1).NumberOfChildren + 1

ReDim Preserve Figures(Figure2).Children(0 To Figures(Figure2).NumberOfChildren)
Figures(Figure2).Children(Figures(Figure2).NumberOfChildren) = FigureCount
Figures(Figure2).NumberOfChildren = Figures(Figure2).NumberOfChildren + 1

If R.P1.X <> EmptyVar And R.P1.Y <> EmptyVar Then
    AddBasePoint R.P1.X, R.P1.Y, , dsIntersect, False, FigureCount, False, False
Else
    AddBasePoint 0, 0, , dsIntersect, False, FigureCount, False, False
    ShowPoint Paper.hDC, PointCount, True
    BasePoint(PointCount).Visible = False
End If
Figures(FigureCount).Points(0) = PointCount

If R.P2.X <> EmptyVar And R.P2.Y <> EmptyVar Then
    AddBasePoint R.P2.X, R.P2.Y, , dsIntersect, False, FigureCount, False, False
Else
    AddBasePoint 0, 0, , dsIntersect, False, FigureCount, False, False
    If IsLine(Figure1) And IsLine(Figure2) Then
        'PointNames.ReplaceItem PointNames.FindItem(BasePoint(PointCount).Name), "_" & BasePoint(PointCount).Name
        BasePoint(PointCount).Name = "_" & BasePoint(PointCount).Name
    End If
    ShowPoint Paper.hDC, PointCount, True
    BasePoint(PointCount).Visible = False
End If
Figures(FigureCount).Points(1) = PointCount

Figures(FigureCount).Description = GetObjectDescription(gotFigure, FigureCount)
BasePoint(PointCount - 1).Description = GetObjectDescription(gotPoint, PointCount - 1)
BasePoint(PointCount).Description = GetObjectDescription(gotPoint, PointCount)

'BasePoint(Figures(FigureCount).Points(0)).ForeColor = Figures(FigureCount).ForeColor
'BasePoint(Figures(FigureCount).Points(1)).ForeColor = Figures(FigureCount).ForeColor

FigureCount = FigureCount + 1
RecalcAuxInfo FigureCount - 1
End Sub

Public Sub AddPointOnFigure(ByVal Figure1 As Long, ByVal X As Double, ByVal Y As Double, Optional ByVal IPointName As String = "", Optional ByVal Display As Boolean = True)
' A point on figure is actually a figure and its dependent point.
' This procedure adds a figure and a corresponding point to the global arrays
' and makes the new figurepoint ready-to-use
'====================================================
Dim P As OnePoint, IPointCount As Long

'====================================================
' resize the figure array and add dsPointOnFigure
'====================================================
RedimPreserveFigures 0, FigureCount
If Not AddFigureAux(FigureCount, dsPointOnFigure, IPointName) Then Exit Sub

Figures(FigureCount).NumberOfPoints = 1

P = GetPointOnFigure(Figure1, X, Y)

ReDim Figures(FigureCount).Points(0 To 0)
ReDim Figures(FigureCount).Parents(0 To 0)
Figures(FigureCount).Parents(0) = Figure1
Figures(FigureCount).NumberOfChildren = 0

AddChildrenRecord Figures(Figure1), FigureCount
'ReDim Preserve Figures(Figure1).Children(0 To Figures(Figure1).NumberOfChildren)
'Figures(Figure1).Children(Figures(Figure1).NumberOfChildren) = FigureCount
'Figures(Figure1).NumberOfChildren = Figures(Figure1).NumberOfChildren + 1

'====================================================

If P.X <> EmptyVar And P.Y <> EmptyVar Then
    AddBasePoint P.X, P.Y, , dsPointOnFigure, False, FigureCount, False, False
Else
    AddBasePoint 0, 0, , dsPointOnFigure, False, FigureCount, False, False
    ShowPoint Paper.hDC, PointCount, True
    BasePoint(PointCount).Visible = False
End If

'====================================================

BasePoint(PointCount).ParentFigure = FigureCount
Figures(FigureCount).Points(0) = PointCount
Figures(FigureCount).ForeColor = setdefcolFigurePoint
'If setdefcolDependentPoint <> setdefcolFigurePoint Or setdefPointShape <> defFigurePointShape Then
'    BasePoint(Figures(FigureCount).Points(0)).Shape = defFigurePointShape
'    BasePoint(Figures(FigureCount).Points(0)).ForeColor = Figures(FigureCount).ForeColor
'End If

If Display Then ShowPoint Paper.hDC, PointCount, , True

Figures(FigureCount).Description = GetObjectDescription(gotFigure, FigureCount)
BasePoint(PointCount).Description = GetObjectDescription(gotPoint, PointCount)

FigureCount = FigureCount + 1
RecalcSemiDependentInfo FigureCount - 1, P.X, P.Y
'RecalcAuxInfo FigureCount - 1
End Sub

Public Sub AddSimmPoint(ByVal Point1 As Long, ByVal Point2 As Long, Optional ByVal SimmPointName As String = "")
RedimPreserveFigures 0, FigureCount

Figures(FigureCount).NumberOfPoints = 3
If Not AddFigureAux(FigureCount, dsSimmPoint, SimmPointName) Then Exit Sub

AddBasePoint 2 * BasePoint(Point2).X - BasePoint(Point1).X, 2 * BasePoint(Point2).Y - BasePoint(Point1).Y, , dsSimmPoint, False, FigureCount, False, False
ReDim Figures(FigureCount).Points(0 To 2)
Figures(FigureCount).Points(0) = PointCount
Figures(FigureCount).Points(1) = Point1
Figures(FigureCount).Points(2) = Point2
'BasePoint(Figures(FigureCount).Points(0)).DrawWidth = defFigureDrawWidth
'BasePoint(Figures(FigureCount).Points(0)).ForeColor = setdefcolDependentPoint

Figures(FigureCount).Description = GetObjectDescription(gotFigure, FigureCount)
BasePoint(PointCount).Description = GetObjectDescription(gotPoint, PointCount)

FigureCount = FigureCount + 1
RecalcAuxInfo FigureCount - 1
End Sub

Public Sub AddSimmPointByLine(ByVal Point1 As Long, ByVal Line1 As Long, Optional ByVal SimmPointName As String = "")
Dim P As OnePoint, R As TwoPoints
RedimPreserveFigures 0, FigureCount

Figures(FigureCount).NumberOfPoints = 2

R = GetLineCoordinates(Line1)
P = GetPerpPoint(BasePoint(Point1).X, BasePoint(Point1).Y, R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
If Not AddFigureAux(FigureCount, dsSimmPointByLine, SimmPointName) Then Exit Sub

AddBasePoint 2 * P.X - BasePoint(Point1).X, 2 * P.Y - BasePoint(Point1).Y, , dsSimmPointByLine, False, FigureCount, False, False

ReDim Figures(FigureCount).Points(0 To 1)
ReDim Figures(FigureCount).Parents(0 To 0)

Figures(FigureCount).NumberOfChildren = 0
Figures(FigureCount).Points(0) = PointCount
Figures(FigureCount).Points(1) = Point1
Figures(FigureCount).Parents(0) = Line1

ReDim Preserve Figures(Line1).Children(0 To Figures(Line1).NumberOfChildren)
Figures(Line1).Children(Figures(Line1).NumberOfChildren) = FigureCount
Figures(Line1).NumberOfChildren = Figures(Line1).NumberOfChildren + 1

'BasePoint(Figures(FigureCount).Points(0)).DrawWidth = defFigureDrawWidth
'BasePoint(Figures(FigureCount).Points(0)).ForeColor = setdefcolDependentPoint

Figures(FigureCount).Description = GetObjectDescription(gotFigure, FigureCount)
BasePoint(PointCount).Description = GetObjectDescription(gotPoint, PointCount)

FigureCount = FigureCount + 1
RecalcAuxInfo FigureCount - 1
End Sub

Public Sub AddInvertedPoint(ByVal Point1 As Long, ByVal Circle1 As Long, Optional ByVal InvPointName As String = "")
Dim P As OnePoint, P2 As OnePoint, R As TwoPoints, Rad As Double
RedimPreserveFigures 0, FigureCount

Figures(FigureCount).NumberOfPoints = 2

Rad = GetCircleRadius(Circle1)
P = GetCircleCenter(Circle1)
P2 = GetInvertedPoint(BasePoint(Point1).X, BasePoint(Point1).Y, P.X, P.Y, Rad)

If Not AddFigureAux(FigureCount, dsInvert, InvPointName) Then Exit Sub
AddBasePoint P2.X, P2.Y, , dsInvert, False, FigureCount, False, False

ReDim Figures(FigureCount).Points(0 To 1)
ReDim Figures(FigureCount).Parents(0 To 0)

Figures(FigureCount).NumberOfChildren = 0
Figures(FigureCount).Points(0) = PointCount
Figures(FigureCount).Points(1) = Point1
Figures(FigureCount).Parents(0) = Circle1
Figures(FigureCount).AuxInfo(2) = Rad
Figures(FigureCount).AuxPoints(2) = P

ReDim Preserve Figures(Circle1).Children(0 To Figures(Circle1).NumberOfChildren)
Figures(Circle1).Children(Figures(Circle1).NumberOfChildren) = FigureCount
Figures(Circle1).NumberOfChildren = Figures(Circle1).NumberOfChildren + 1

'BasePoint(Figures(FigureCount).Points(0)).DrawWidth = defFigureDrawWidth
'BasePoint(Figures(FigureCount).Points(0)).ForeColor = setdefcolDependentPoint

Figures(FigureCount).Description = GetObjectDescription(gotFigure, FigureCount)
BasePoint(PointCount).Description = GetObjectDescription(gotPoint, PointCount)

FigureCount = FigureCount + 1
RecalcAuxInfo FigureCount - 1
End Sub

Public Sub AddMeasureDistance(ByVal Point1 As Long, ByVal Point2 As Long)
RedimPreserveFigures 0, FigureCount
If Not AddFigureAux(FigureCount, dsMeasureDistance) Then Exit Sub

Figures(FigureCount).NumberOfPoints = 2
ReDim Figures(FigureCount).Points(0 To 1)
Figures(FigureCount).Points(0) = Point1
Figures(FigureCount).Points(1) = Point2
Figures(FigureCount).ForeColor = setdefcolPointName

Figures(FigureCount).Description = GetObjectDescription(gotFigure, FigureCount)

FigureCount = FigureCount + 1
RecalcAuxInfo FigureCount - 1
End Sub

Public Sub AddMeasureAngle(ByVal Point1 As Long, ByVal Point2 As Long, ByVal Point3 As Long)
RedimPreserveFigures 0, FigureCount
If Not AddFigureAux(FigureCount, dsMeasureAngle) Then Exit Sub

Figures(FigureCount).NumberOfPoints = 3
ReDim Figures(FigureCount).Points(0 To 2)
Figures(FigureCount).Points(0) = Point1
Figures(FigureCount).Points(1) = Point2
Figures(FigureCount).Points(2) = Point3
Figures(FigureCount).ForeColor = setdefcolPointName

'=================================================
' DrawStyle here indicates whether to show angle marks, and if yes
' then how many, and whether to hide measurement text
'
' DrawStyle = 0         =>          0 AngleMarks
' DrawStyle = 1         =>          1 AngleMarks
' DrawStyle = 2         =>          2 AngleMarks
' DrawStyle = 3         =>          3 AngleMarks
' DrawStyle = 4         =>          1 AngleMarks and hide text
' DrawStyle = 5         =>          2 AngleMarks and hide text
' DrawStyle = 6         =>          3 AngleMarks and hide text

'=================================================
Figures(FigureCount).DrawStyle = 1

Figures(FigureCount).Description = GetObjectDescription(gotFigure, FigureCount)

FigureCount = FigureCount + 1
RecalcAuxInfo FigureCount - 1
End Sub

Public Sub AddMeasureArea(P() As Long, Optional ByVal ShouldRecord As Boolean = True)
Dim LB As Long, UB As Long, Z As Long
Dim X As Double, Y As Double, S As String, T As String
Dim pAction As Action

LB = LBound(P)
UB = UBound(P)

If P(LB) = P(UB) Then UB = UB - 1

If UB = LB + 1 Then
    X = BasePoint(P(LB)).X
    Y = BasePoint(P(LB)).Y
    T = "(" & BasePoint(P(LB)).Name & "," & BasePoint(P(UB)).Name & ")"
    S = GetString(ResArea) & " = [PI*(Distance" & T & ")^2]"
    
Else
    
    For Z = LB To UB
        X = X + BasePoint(P(Z)).X
        Y = Y + BasePoint(P(Z)).Y
    Next
    X = X / (UB + 1 - LB)
    Y = Y / (UB + 1 - LB)
    
    T = "("
    For Z = LB To UB
        T = T & BasePoint(P(Z)).Name & IIf(Z = UB, "", ",")
    Next
    T = T & ")"
    
    S = GetString(ResArea) & T & " = [Area" & T & "]"
End If

pAction.Type = actAddLabel
pAction.pLabel = LabelCount + 1
RecordAction pAction

AddTextLabel S, X, Y
End Sub

Public Function IsInBasePoint(ByVal X As Double, ByVal Y As Double, ByVal index As Long) As Boolean
Dim W2 As Double, cx As Double, cy As Double

If BasePoint(index).Visible = False Or Not BasePoint(index).Enabled Then Exit Function

W2 = BasePoint(index).Width / 2 + Sensitivity
cx = BasePoint(index).X
cy = BasePoint(index).Y
If X >= cx - W2 And X <= cx + W2 And Y >= cy - W2 And Y <= cy + W2 Then IsInBasePoint = True

'On Local Error Resume Next
'Dim CX As Double, CY As Double
'Dim W2 As Double
'
'With BasePoint(Index)
'    If .Visible = False Then Exit Function
'    W2 = .Width / 2
'    If .Shape = vbShapeSquare Then
'        If X >= .X - W2 - Sensitivity And X <= .X + W2 + Sensitivity And Y >= .Y - W2 - Sensitivity And Y <= .Y + W2 + Sensitivity Then IsInBasePoint = True
'    Else
'        CX = .X
'        CY = .Y
'        If Sqr((X - CX) * (X - CX) + (Y - CY) * (Y - CY)) <= W2 + Sensitivity Then IsInBasePoint = True
'    End If
'End With
End Function

Public Function IsInPointLabel(ByVal X As Double, ByVal Y As Double, ByVal index As Long) As Boolean
Dim W2 As Double, cx As Double, cy As Double, WX As Double, WY As Double, tX As Double, tY As Double
If BasePoint(index).Visible = False Or BasePoint(index).Hide Or Not BasePoint(index).ShowName Then Exit Function

W2 = Sensitivity
ToPhysicalLength W2
tX = BasePoint(index).X
tY = BasePoint(index).Y
ToPhysical tX, tY
cx = tX + BasePoint(index).LabelOffsetX - BasePoint(index).LabelWidth / 2
cy = tY + BasePoint(index).LabelOffsetY - BasePoint(index).LabelHeight
WX = BasePoint(index).LabelWidth
WY = BasePoint(index).LabelHeight
If X >= cx - W2 And X <= cx + WX + W2 And Y >= cy - W2 And Y <= cy + WY + W2 Then IsInPointLabel = True
End Function

Public Function RecalcMeasure(ByVal hDC As Long, ByVal szS As String, Optional ByVal Ang As Single = 0) As OnePoint
Dim LF As LOGFONT, i As Long, NF As Long, lpPoint As POINTAPI, Q As Long, lpSize As Size
Dim Z As Long

LF.lfWidth = 0
LF.lfEscapement = CLng(Ang * 10)
LF.lfOrientation = LF.lfEscapement
LF.lfWeight = 400
LF.lfItalic = 0
LF.lfUnderline = 0
LF.lfStrikeOut = 0
LF.lfCharSet = 0
LF.lfOutPrecision = 0
LF.lfClipPrecision = 0
LF.lfQuality = 2
LF.lfPitchAndFamily = 0
For Z = 0 To 31
    LF.lfFaceName(Z) = ByteFontName(Z)
Next
LF.lfHeight = defSLabelFontSize * -20 / Screen.TwipsPerPixelY

i = CreateFontIndirect(LF)
NF = SelectObject(hDC, i)

SetTextAlign hDC, TA_BOTTOM Or TA_CENTER
GetTextExtentPoint32 hDC, szS, Len(szS), lpSize

SelectObject hDC, NF
DeleteObject i

RecalcMeasure.X = lpSize.cx
RecalcMeasure.Y = lpSize.cy
End Function

Public Sub HideFigureWithPoint(ByVal index As Long)
Dim Z As Long

For Z = 0 To FigureCount - 1
    If FigureHasPoint(Z, index) Then HideFigure Z
Next Z
End Sub

Public Sub RecalcFigureWithPoint(ByVal index As Long)
Dim Z As Long

For Z = 0 To FigureCount - 1
    If FigureHasPoint(Z, index) Then RecalcAuxInfo Z
Next Z
End Sub


Public Sub BringPointToFront(ByVal index As Long)
Dim MaxZOrder As Long, pAction As Action, Z As Long
If Not IsPoint(index) Then Exit Sub
'MaxZOrder = 1
'For Z = 1 To PointCount
'    If BasePoint(Z).ZOrder > BasePoint(MaxZOrder).ZOrder Then MaxZOrder = Z
'Next Z
'T = BasePoint(MaxZOrder).ZOrder

pAction.Type = actPointZOrder
ReDim pAction.AuxInfo(1 To 2) As Double
pAction.AuxInfo(1) = index
pAction.AuxInfo(2) = BasePoint(index).ZOrder

For Z = 1 To PointCount
    If BasePoint(Z).ZOrder > BasePoint(index).ZOrder Then BasePoint(Z).ZOrder = BasePoint(Z).ZOrder - 1
Next
BasePoint(index).ZOrder = PointCount

pAction.pPoint = index
RecordAction pAction
ShowAll
End Sub

Public Sub BringFigureToFront(ByVal index As Long)
Dim MaxZOrder As Long, pAction As Action, Z As Long
If Not IsFigure(index) Or FigureCount = 0 Then Exit Sub

MaxZOrder = 0
For Z = 0 To FigureCount - 1
    If Figures(Z).ZOrder > MaxZOrder Then MaxZOrder = Figures(Z).ZOrder
Next Z

'v1 = FigureVisualOrder(Figures(Index).ZOrder + 1)
'v2 = FigureVisualOrder(FigureVisualOrder.Count)
'T1 = Figures(Index).ZOrder + 1
'T2 = FigureVisualOrder.Count
'FigureVisualOrder.Remove T1
'If T1 < FigureVisualOrder.Count Then FigureVisualOrder.Add v2, , , T1 Else FigureVisualOrder.Add v2
'FigureVisualOrder.Remove T2
'FigureVisualOrder.Add v1

'Swap Figures(Index).ZOrder, Figures(MaxZOrder).ZOrder

pAction.Type = actFigureZOrder
ReDim pAction.AuxInfo(1 To 2) As Double
pAction.AuxInfo(1) = index
pAction.AuxInfo(2) = Figures(index).ZOrder

For Z = 0 To FigureCount - 1
    If Figures(Z).ZOrder > Figures(index).ZOrder Then Figures(Z).ZOrder = Figures(Z).ZOrder - 1
Next
Figures(index).ZOrder = MaxZOrder

pAction.pFigure = index
RecordAction pAction
ShowAll
End Sub

Public Function GetLineFromSegment(ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double) As TwoPoints
Dim X3 As Double, Y3 As Double, X4 As Double, Y4 As Double

If X1 = X2 And Y1 = Y2 Then
    X3 = X1
    Y3 = Y1
    X4 = X1
    Y4 = Y1
ElseIf X1 = X2 Then
    X3 = X1
    X4 = X1
    Y3 = CanvasBorders.P1.Y
    Y4 = CanvasBorders.P2.Y
    If Y1 < Y2 Then Swap Y3, Y4
ElseIf Y1 = Y2 Then
    Y3 = Y1
    Y4 = Y1
    X3 = CanvasBorders.P1.X
    X4 = CanvasBorders.P2.X
    If X1 < X2 Then Swap X3, X4
Else
    If Y1 > Y2 Then Y3 = CanvasBorders.P1.Y Else Y3 = CanvasBorders.P2.Y
    X3 = X1 - (Y1 - Y3) * (X2 - X1) / (Y2 - Y1)
    If X3 < CanvasBorders.P1.X Then
        X3 = CanvasBorders.P1.X
        Y3 = Y1 - (X1 - X3) * (Y2 - Y1) / (X2 - X1)
    ElseIf X3 > CanvasBorders.P2.X Then
        X3 = CanvasBorders.P2.X
        Y3 = Y1 - (X1 - X3) * (Y2 - Y1) / (X2 - X1)
    End If
    
    If X1 > X2 Then X4 = CanvasBorders.P2.X Else X4 = CanvasBorders.P1.X
    Y4 = Y2 + (X4 - X2) * (Y2 - Y1) / (X2 - X1)
    If Y4 < CanvasBorders.P1.Y Then
        Y4 = CanvasBorders.P1.Y
        X4 = X2 + (X2 - X1) * (Y4 - Y2) / (Y2 - Y1)
    ElseIf Y4 > CanvasBorders.P2.Y Then
        Y4 = CanvasBorders.P2.Y
        X4 = X2 + (X2 - X1) * (Y4 - Y2) / (Y2 - Y1)
    End If
End If

GetLineFromSegment.P1.X = X3
GetLineFromSegment.P1.Y = Y3
GetLineFromSegment.P2.X = X4
GetLineFromSegment.P2.Y = Y4
End Function

Public Function GetFigureByPoint(ByVal X As Double, ByVal Y As Double) As Long
Dim TempFig As Long, MaxFigureZOrder As Long, Z As Long, Q As Long

GetFigureByPoint = -1
TempFig = -1
If FigureCount = 0 Then Exit Function

For Z = 0 To FigureCount - 1
    If Figures(Z).NumberOfPoints > 0 Then
        For Q = 0 To Figures(Z).NumberOfPoints - 1
            If BasePoint(Figures(Z).Points(Q)).Visible = False And Figures(Z).FigureType <> dsIntersect Then GoTo NextZ
        Next Q
    End If
    If PointBelongsToFigure(X, Y, Z) And Not Figures(Z).Hide Then
        If TempFig = -1 Then
            TempFig = Z
            MaxFigureZOrder = Figures(Z).ZOrder
        Else
            If Figures(Z).ZOrder > Figures(TempFig).ZOrder Then
                TempFig = Z
                MaxFigureZOrder = Figures(Z).ZOrder
            End If
        End If
    End If
NextZ: Next Z
GetFigureByPoint = TempFig
End Function

Public Function GetFiguresByPoint(ByVal X As Double, ByVal Y As Double, Optional ByVal RestrictedPoint As Long = EmptyVar, Optional ByVal IncludeHidden As Boolean = True) As Long()
Dim FArray() As Long, FCount As Long, Z As Long

ReDim FArray(1 To 1)
FArray(1) = -1
If FigureCount > 0 Then
    For Z = 0 To FigureCount - 1
        If PointBelongsToFigure(X, Y, Z, IncludeHidden) Then
            If RestrictedPoint <> EmptyVar Then
                If FigureHasPoint(Z, RestrictedPoint) Then GoTo NextZ
            End If
            FCount = FCount + 1
            ReDim Preserve FArray(1 To FCount)
            FArray(FCount) = Z
        End If
NextZ:    Next Z
End If
GetFiguresByPoint = FArray
End Function

Public Function GetPointFromCursor(ByVal X As Double, ByVal Y As Double) As Long
Dim tZ As Long, Z As Long

For Z = 1 To PointCount
    If IsInBasePoint(X, Y, Z) Then
        If tZ > 0 Then
            If BasePoint(tZ).ZOrder < BasePoint(Z).ZOrder Then tZ = Z
        Else
            tZ = Z
        End If
    End If
Next Z

GetPointFromCursor = tZ
End Function

Public Function GetPointLabelFromCursor(ByVal X As Double, ByVal Y As Double) As Long
Dim tZ As Long, Z As Long
ToPhysical X, Y
For Z = 1 To PointCount
    If IsInPointLabel(X, Y, Z) Then
        If BasePoint(Z).Enabled Then
            If tZ > 0 Then
                If BasePoint(tZ).ZOrder < BasePoint(Z).ZOrder Then tZ = Z
            Else
                tZ = Z
            End If
        End If
    End If
Next Z
GetPointLabelFromCursor = tZ
End Function

Public Function GetPointsFromCursor(ByVal X As Double, ByVal Y As Double) As Long()
' X, Y - logical coordinates
'=================================================================
Dim PCount As Long, PArray() As Long, Z As Long

If PointCount = 0 Then Exit Function
ReDim PArray(1 To 1)
For Z = 1 To PointCount
    If IsInBasePoint(X, Y, Z) Then
        PCount = PCount + 1
        ReDim Preserve PArray(1 To PCount)
        PArray(PCount) = Z
    End If
Next Z
GetPointsFromCursor = PArray
End Function

Public Function GetLineCoordinates(ByVal Line1 As Long) As TwoPoints
Dim R As TwoPoints, P As OnePoint
Dim X1 As Double, Y1 As Double, X2 As Double, Y2 As Double

Select Case Figures(Line1).FigureType
    Case dsSegment
        X1 = BasePoint(Figures(Line1).Points(0)).X
        Y1 = BasePoint(Figures(Line1).Points(0)).Y
        X2 = BasePoint(Figures(Line1).Points(1)).X
        Y2 = BasePoint(Figures(Line1).Points(1)).Y
        R.P1.X = X1
        R.P1.Y = Y1
        R.P2.X = X2
        R.P2.Y = Y2
        
    Case dsRay
        X1 = BasePoint(Figures(Line1).Points(0)).X
        Y1 = BasePoint(Figures(Line1).Points(0)).Y
        X2 = BasePoint(Figures(Line1).Points(1)).X
        Y2 = BasePoint(Figures(Line1).Points(1)).Y
        R = GetLineFromSegment(X1, Y1, X2, Y2)
        R.P2.X = X1
        R.P2.Y = Y1
        
    Case dsLine_2Points
        X1 = BasePoint(Figures(Line1).Points(0)).X
        Y1 = BasePoint(Figures(Line1).Points(0)).Y
        X2 = BasePoint(Figures(Line1).Points(1)).X
        Y2 = BasePoint(Figures(Line1).Points(1)).Y
        R = GetLineFromSegment(X1, Y1, X2, Y2)
    
    Case dsBisector
        R.P1.X = BasePoint(Figures(Line1).Points(1)).X
        R.P1.Y = BasePoint(Figures(Line1).Points(1)).Y
        R.P2 = GetBisector(BasePoint(Figures(Line1).Points(0)).X, BasePoint(Figures(Line1).Points(0)).Y, R.P1.X, R.P1.Y, BasePoint(Figures(Line1).Points(2)).X, BasePoint(Figures(Line1).Points(2)).Y)
        If R.P2.X = EmptyVar Then R.P1 = EmptyOnePoint Else R = GetLineFromSegment(R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
    
    Case dsLine_PointAndParallelLine
        X1 = BasePoint(Figures(Line1).Points(0)).X
        Y1 = BasePoint(Figures(Line1).Points(0)).Y
        R = GetLineCoordinatesAbsolute(Figures(Line1).Parents(0)) '?????
        X2 = X1 + R.P1.X - R.P2.X
        Y2 = Y1 + R.P1.Y - R.P2.Y
        R = GetLineFromSegment(X1, Y1, X2, Y2)
    
    Case dsLine_PointAndPerpendicularLine
        X1 = BasePoint(Figures(Line1).Points(0)).X
        Y1 = BasePoint(Figures(Line1).Points(0)).Y
        R = GetLineCoordinatesAbsolute(Figures(Line1).Parents(0)) '?????
        R = GetPerpendicularLine(X1, Y1, R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
        
    Case dsAnLineGeneral
        R = GetGeneralAnLine(Figures(Line1).AuxInfo(1), Figures(Line1).AuxInfo(2), Figures(Line1).AuxInfo(3))

    Case dsAnLineCanonic
        R = GetCanonicAnLine(Figures(Line1).AuxInfo(1), Figures(Line1).AuxInfo(2), Figures(Line1).AuxInfo(3), Figures(Line1).AuxInfo(4))

    Case dsAnLineNormal
        R = GetNormalAnLine(Figures(Line1).AuxInfo(1), Figures(Line1).AuxInfo(2))
    
    Case dsAnLineNormalPoint
        R = GetNormalPointAnLine(Figures(Line1).AuxInfo(1), Figures(Line1).AuxInfo(2), Figures(Line1).AuxInfo(3), Figures(Line1).AuxInfo(4))

End Select
GetLineCoordinates = R
End Function

Public Function GetLineCoordinatesAbsolute(ByVal Line1 As Long) As TwoPoints
Dim R As TwoPoints
Dim P As OnePoint
Dim X1 As Double, Y1 As Double, X2 As Double, Y2 As Double
Dim a As Double, b As Double, c As Double, D As Double

Select Case Figures(Line1).FigureType
    Case dsSegment, dsRay, dsLine_2Points
        R.P1.X = BasePoint(Figures(Line1).Points(0)).X
        R.P1.Y = BasePoint(Figures(Line1).Points(0)).Y
        R.P2.X = BasePoint(Figures(Line1).Points(1)).X
        R.P2.Y = BasePoint(Figures(Line1).Points(1)).Y
    
    Case dsBisector
        R.P1.X = BasePoint(Figures(Line1).Points(1)).X
        R.P1.Y = BasePoint(Figures(Line1).Points(1)).Y
        R.P2 = GetBisector(BasePoint(Figures(Line1).Points(0)).X, BasePoint(Figures(Line1).Points(0)).Y, R.P1.X, R.P1.Y, BasePoint(Figures(Line1).Points(2)).X, BasePoint(Figures(Line1).Points(2)).Y)
        If R.P2.X = EmptyVar Then R.P1 = EmptyOnePoint
    
    Case dsLine_PointAndParallelLine
        X1 = BasePoint(Figures(Line1).Points(0)).X
        Y1 = BasePoint(Figures(Line1).Points(0)).Y
        R = GetLineCoordinatesAbsolute(Figures(Line1).Parents(0))
        X2 = X1 + R.P1.X - R.P2.X
        Y2 = Y1 + R.P1.Y - R.P2.Y
        R.P1.X = X1
        R.P1.Y = Y1
        R.P2.X = X2
        R.P2.Y = Y2
    
    Case dsLine_PointAndPerpendicularLine
        X1 = BasePoint(Figures(Line1).Points(0)).X
        Y1 = BasePoint(Figures(Line1).Points(0)).Y
        R = GetLineCoordinatesAbsolute(Figures(Line1).Parents(0))
        X2 = X1 + R.P2.Y - R.P1.Y
        Y2 = Y1 + R.P1.X - R.P2.X
        R.P1.X = X1
        R.P1.Y = Y1
        R.P2.X = X2
        R.P2.Y = Y2
        
    Case dsAnLineGeneral, dsAnLineCanonic, dsAnLineNormal, dsAnLineNormalPoint
        'R = GetGeneralAnLineAbsolute(Figures(Line1).AuxInfo(1), Figures(Line1).AuxInfo(2), Figures(Line1).AuxInfo(3))
        R.P1 = Figures(Line1).AuxPoints(3)
        R.P2 = Figures(Line1).AuxPoints(4)

'    Case dsAnLineCanonic
'        R = GetCanonicAnLineAbsolute(Figures(Line1).AuxInfo(1), Figures(Line1).AuxInfo(2), Figures(Line1).AuxInfo(3), Figures(Line1).AuxInfo(4))
'
'    Case dsAnLineNormal
'        R = GetNormalAnLineAbsolute(Figures(Line1).AuxInfo(1), Figures(Line1).AuxInfo(2))
'
'    Case dsAnLineNormalPoint
'        R = GetNormalPointAnLineAbsolute(Figures(Line1).AuxInfo(1), Figures(Line1).AuxInfo(2), Figures(Line1).AuxInfo(3), Figures(Line1).AuxInfo(4))
End Select
GetLineCoordinatesAbsolute = R
End Function

Public Sub ClearAll(Optional ByVal ClearUndo As Boolean = True, Optional ByVal ShouldRefresh As Boolean = True, Optional ByVal ShouldTryToAddAxes As Boolean = True)
modDrawing.InitDrawing ShouldRefresh, ClearUndo, ShouldTryToAddAxes

FormMain.SelectTool dsSelect
End Sub

Public Sub ClearPrivileges()
privNoAlter = False
End Sub

Public Function GetIntersectionPoints(ByVal Figure1 As Long, ByVal Figure2 As Long) As TwoPoints
Dim R As TwoPoints

If Figures(Figure1).Visible = False Or Figures(Figure2).Visible = False Then Exit Function
R.P1.X = EmptyVar
R.P1.Y = EmptyVar
R.P2.X = EmptyVar
R.P2.Y = EmptyVar

Select Case Figures(Figure1).FigureType
    Case dsCircle_CenterAndCircumPoint, dsCircle_CenterAndTwoPoints, dsCircle_ArcCenterAndRadiusAndTwoPoints, dsAnCircle
        If IsLine(Figure2) Then R = IntersectCircleAndLine(Figure1, Figure2)
        If IsCircle(Figure2) Then R = IntersectCircles(Figure1, Figure2)
    Case dsLine_2Points, dsLine_PointAndParallelLine, dsLine_PointAndPerpendicularLine, dsBisector, dsSegment, dsRay, dsAnLineGeneral, dsAnLineCanonic, dsAnLineNormal, dsAnLineNormalPoint
        If IsLine(Figure2) Then R.P1 = IntersectLines(Figure1, Figure2)
        If IsCircle(Figure2) Then R = IntersectCircleAndLine(Figure2, Figure1)
End Select

GetIntersectionPoints = R
End Function

Public Function IntersectLines(ByVal Line1 As Long, ByVal Line2 As Long) As OnePoint
Dim P As OnePoint
Dim R1 As TwoPoints
Dim R2 As TwoPoints
Dim T As Double

R1 = GetLineCoordinatesAbsolute(Line1)
R2 = GetLineCoordinatesAbsolute(Line2)

P = GetIntersectionOfLines(R1, R2)
If P.X = EmptyVar Or P.Y = EmptyVar Then IntersectLines = P: Exit Function

If Figures(Line1).FigureType = dsSegment Then
    If R1.P1.X > R1.P2.X Then
        T = R1.P1.X
        R1.P1.X = R1.P2.X
        R1.P2.X = T
    End If
    If R1.P1.Y > R1.P2.Y Then
        T = R1.P1.Y
        R1.P1.Y = R1.P2.Y
        R1.P2.Y = T
    End If
    
    If R1.P1.X = R1.P2.X Then
        If P.Y < R1.P1.Y Or P.Y > R1.P2.Y Then
            P.X = EmptyVar
            P.Y = EmptyVar
            GoTo Fin
        End If
    ElseIf R1.P1.Y = R1.P2.Y Then
        If P.X < R1.P1.X Or P.X > R1.P2.X Then
            P.X = EmptyVar
            P.Y = EmptyVar
            GoTo Fin
        End If
    Else
        If P.X < R1.P1.X Or P.X > R1.P2.X Or P.Y < R1.P1.Y Or P.Y > R1.P2.Y Then
            P.X = EmptyVar
            P.Y = EmptyVar
            GoTo Fin
        End If
    End If
End If

If Figures(Line2).FigureType = dsSegment Then
    If R2.P1.X > R2.P2.X Then
        T = R2.P1.X
        R2.P1.X = R2.P2.X
        R2.P2.X = T
    End If
    If R2.P1.Y > R2.P2.Y Then
        T = R2.P1.Y
        R2.P1.Y = R2.P2.Y
        R2.P2.Y = T
    End If
    
    If R2.P1.X = R2.P2.X Then
        If P.Y < R2.P1.Y Or P.Y > R2.P2.Y Then
            P.X = EmptyVar
            P.Y = EmptyVar
            GoTo Fin
        End If
    ElseIf R2.P1.Y = R2.P2.Y Then
        If P.X < R2.P1.X Or P.X > R2.P2.X Then
            P.X = EmptyVar
            P.Y = EmptyVar
            GoTo Fin
        End If
    Else
        If P.X < R2.P1.X Or P.X > R2.P2.X Or P.Y < R2.P1.Y Or P.Y > R2.P2.Y Then
            P.X = EmptyVar
            P.Y = EmptyVar
            GoTo Fin
        End If
    End If
End If

'If Figures(Line2).FigureType = dsSegment Then
'    If R2.P1.X = R2.P2.X Then
'        If P.Y < Minimum(R2.P1.Y, R2.P2.Y) Or P.Y > Maximum(R2.P1.Y, R2.P2.Y) Then
'            P.X = EmptyVar
'            P.Y = EmptyVar
'            GoTo Fin
'        End If
'    ElseIf R2.P1.Y = R2.P2.Y Then
'        If P.X < Minimum(R2.P1.X, R2.P2.X) Or P.X > Maximum(R2.P1.X, R2.P2.X) Then
'            P.X = EmptyVar
'            P.Y = EmptyVar
'            GoTo Fin
'        End If
'    Else
'        If (P.X < Minimum(R2.P1.X, R2.P2.X) Or P.X > Maximum(R2.P1.X, R2.P2.X)) Or (P.Y < Minimum(R2.P1.Y, R2.P2.Y) Or P.Y > Maximum(R2.P1.Y, R2.P2.Y)) Then
'            P.X = EmptyVar
'            P.Y = EmptyVar
'            GoTo Fin
'        End If
'    End If
'End If

If Figures(Line1).FigureType = dsRay Then
    If R1.P1.X = R1.P2.X Then
        If Sgn(R1.P1.Y - P.Y) <> Sgn(R1.P1.Y - R1.P2.Y) Then
            P.X = EmptyVar
            P.Y = EmptyVar
            GoTo Fin
        End If
    ElseIf R1.P1.Y = R1.P2.Y Then
        If Sgn(R1.P1.X - P.X) <> Sgn(R1.P1.X - R1.P2.X) Then
            P.X = EmptyVar
            P.Y = EmptyVar
            GoTo Fin
        End If
    Else
        If Sgn(R1.P1.X - P.X) <> Sgn(R1.P1.X - R1.P2.X) Or Sgn(R1.P1.Y - P.Y) <> Sgn(R1.P1.Y - R1.P2.Y) Then
            P.X = EmptyVar
            P.Y = EmptyVar
            GoTo Fin
        End If
    End If
End If

If Figures(Line2).FigureType = dsRay Then
    If R2.P1.X = R2.P2.X Then
        If Sgn(R2.P1.Y - P.Y) <> Sgn(R2.P1.Y - R2.P2.Y) Then
            P.X = EmptyVar
            P.Y = EmptyVar
        End If
    ElseIf R2.P1.Y = R2.P2.Y Then
        If Sgn(R2.P1.X - P.X) <> Sgn(R2.P1.X - R2.P2.X) Then
            P.X = EmptyVar
            P.Y = EmptyVar
        End If
    Else
        If Sgn(R2.P1.X - P.X) <> Sgn(R2.P1.X - R2.P2.X) Or Sgn(R2.P1.Y - P.Y) <> Sgn(R2.P1.Y - R2.P2.Y) Then
            P.X = EmptyVar
            P.Y = EmptyVar
        End If
    End If
End If
'If Figures(Line1).FigureType = dsRay Then
'    If R1.P1.X = R1.P2.X Then
'        If Sgn(R1.P2.Y - P.Y) <> Sgn(R1.P2.Y - R1.P1.Y) Then
'            P.X = EmptyVar
'            P.Y = EmptyVar
'            GoTo Fin
'        End If
'    ElseIf R1.P1.Y = R1.P2.Y Then
'        If Sgn(R1.P2.X - P.X) <> Sgn(R1.P2.X - R1.P1.X) Then
'            P.X = EmptyVar
'            P.Y = EmptyVar
'            GoTo Fin
'        End If
'    Else
'        If Sgn(R1.P2.X - P.X) <> Sgn(R1.P2.X - R1.P1.X) Or Sgn(R1.P2.Y - P.Y) <> Sgn(R1.P2.Y - R1.P1.Y) Then
'            P.X = EmptyVar
'            P.Y = EmptyVar
'            GoTo Fin
'        End If
'    End If
'End If
'If Figures(Line2).FigureType = dsRay Then
'    If R2.P1.X = R2.P2.X Then
'        If Sgn(R2.P2.Y - P.Y) <> Sgn(R2.P2.Y - R2.P1.Y) Then
'            P.X = EmptyVar
'            P.Y = EmptyVar
'        End If
'    ElseIf R2.P1.Y = R2.P2.Y Then
'        If Sgn(R2.P2.X - P.X) <> Sgn(R2.P2.X - R2.P1.X) Then
'            P.X = EmptyVar
'            P.Y = EmptyVar
'        End If
'    Else
'        If Sgn(R2.P2.X - P.X) <> Sgn(R2.P2.X - R2.P1.X) Or Sgn(R2.P2.Y - P.Y) <> Sgn(R2.P2.Y - R2.P1.Y) Then
'            P.X = EmptyVar
'            P.Y = EmptyVar
'        End If
'    End If
'End If

Fin:
IntersectLines = P
End Function

Public Function IntersectCircleAndLine(ByVal Circle1 As Long, ByVal Line1 As Long) As TwoPoints
Dim R As TwoPoints, TP As TwoPoints
Dim P As OnePoint
Dim Rad As Double
Dim X3 As Double, Y3 As Double, X4 As Double, Y4 As Double, X5 As Double, Y5 As Double
Dim A1 As Double, A2 As Double, A3 As Double

R = GetLineCoordinatesAbsolute(Line1)
P = GetCircleCenter(Circle1)
Rad = GetCircleRadius(Circle1)

TP = GetIntersectionOfCircleAndLine(P, Rad, R)

If Figures(Line1).FigureType = dsSegment Then
    If TP.P1.X <> EmptyVar And TP.P1.Y <> EmptyVar Then
        If Not PointInRectangle(TP.P1.X, TP.P1.Y, R.P1.X, R.P1.Y, R.P2.X, R.P2.Y) Then
        'If (tP.P1.X < Minimum(R.P1.X, R.P2.X) Or tP.P1.X > Maximum(R.P1.X, R.P2.X)) Or (tP.P1.Y < Minimum(R.P1.Y, R.P2.Y) Or tP.P1.Y > Maximum(R.P1.Y, R.P2.Y)) Then
            TP.P1.X = EmptyVar
            TP.P1.Y = EmptyVar
        End If
    End If
    If TP.P2.X <> EmptyVar And TP.P2.Y <> EmptyVar Then
        If Not PointInRectangle(TP.P2.X, TP.P2.Y, R.P1.X, R.P1.Y, R.P2.X, R.P2.Y) Then
        'If (tP.P2.X < Minimum(R.P1.X, R.P2.X) Or tP.P2.X > Maximum(R.P1.X, R.P2.X)) Or (tP.P2.Y < Minimum(R.P1.Y, R.P2.Y) Or tP.P2.Y > Maximum(R.P1.Y, R.P2.Y)) Then
            TP.P2.X = EmptyVar
            TP.P2.Y = EmptyVar
        End If
    End If
End If

If Figures(Line1).FigureType = dsRay Then
    If TP.P1.X <> EmptyVar And TP.P1.Y <> EmptyVar Then
        If (Sgn(TP.P1.X - R.P1.X) <> Sgn(R.P2.X - R.P1.X) Or Sgn(TP.P1.Y - R.P1.Y) <> Sgn(R.P2.Y - R.P1.Y)) Then
            TP.P1.X = EmptyVar
            TP.P1.Y = EmptyVar
        End If
    End If
    If TP.P2.X <> EmptyVar And TP.P2.Y <> EmptyVar Then
        If (Sgn(TP.P2.X - R.P1.X) <> Sgn(R.P2.X - R.P1.X) Or Sgn(TP.P2.Y - R.P1.Y) <> Sgn(R.P2.Y - R.P1.Y)) Then
            TP.P2.X = EmptyVar
            TP.P2.Y = EmptyVar
        End If
    End If
'    If TP.P1.X <> EmptyVar And TP.P1.Y <> EmptyVar Then
'        If (Sgn(R.P2.X - TP.P1.X) <> Sgn(R.P2.X - R.P1.X) Or Sgn(R.P2.Y - TP.P1.Y) <> Sgn(R.P2.Y - R.P1.Y)) Then
'            TP.P1.X = EmptyVar
'            TP.P1.Y = EmptyVar
'        End If
'    End If
'    If TP.P2.X <> EmptyVar And TP.P2.Y <> EmptyVar Then
'        If (Sgn(R.P2.X - TP.P2.X) <> Sgn(R.P2.X - R.P1.X) Or Sgn(R.P2.Y - TP.P2.Y) <> Sgn(R.P2.Y - R.P1.Y)) Then
'            TP.P2.X = EmptyVar
'            TP.P2.Y = EmptyVar
'        End If
'    End If
End If

If Figures(Circle1).FigureType = dsCircle_ArcCenterAndRadiusAndTwoPoints Then
    X3 = BasePoint(Figures(Circle1).Points(2)).X
    Y3 = BasePoint(Figures(Circle1).Points(2)).Y
    X4 = BasePoint(Figures(Circle1).Points(3)).X
    Y4 = BasePoint(Figures(Circle1).Points(3)).Y
    X5 = BasePoint(Figures(Circle1).Points(4)).X
    Y5 = BasePoint(Figures(Circle1).Points(4)).Y
    A1 = GetAngle(X3, Y3, X4, Y4)
    A2 = GetAngle(X3, Y3, X5, Y5)
    
    A3 = GetAngle(X3, Y3, TP.P1.X, TP.P1.Y)
    If A2 < A1 Then
        If Not (A3 < A1 And A3 > A2) Then
            TP.P1.X = EmptyVar
            TP.P1.Y = EmptyVar
        End If
    Else
        If Not (A1 > A3 Or A3 > A2) Then
            TP.P1.X = EmptyVar
            TP.P1.Y = EmptyVar
        End If
    End If

    A3 = GetAngle(X3, Y3, TP.P2.X, TP.P2.Y)
    If A2 < A1 Then
        If Not (A3 < A1 And A3 > A2) Then
            TP.P2.X = EmptyVar
            TP.P2.Y = EmptyVar
        End If
    Else
        If Not (A1 > A3 Or A3 > A2) Then
            TP.P2.X = EmptyVar
            TP.P2.Y = EmptyVar
        End If
    End If
    
'    A3 = GetAngle(X3, Y3, TP.P1.X, TP.P1.Y)
'    If A2 < A1 Then
'        If A3 < A1 And A3 > A2 Then
'            TP.P1.X = EmptyVar
'            TP.P1.Y = EmptyVar
'        End If
'    Else
'        If A1 > A3 Or A3 > A2 Then
'            TP.P1.X = EmptyVar
'            TP.P1.Y = EmptyVar
'        End If
'    End If
'
'    A3 = GetAngle(X3, Y3, TP.P2.X, TP.P2.Y)
'    If A2 < A1 Then
'        If A3 < A1 And A3 > A2 Then
'            TP.P2.X = EmptyVar
'            TP.P2.Y = EmptyVar
'        End If
'    Else
'        If A1 > A3 Or A3 > A2 Then
'            TP.P2.X = EmptyVar
'            TP.P2.Y = EmptyVar
'        End If
'    End If
End If
IntersectCircleAndLine = TP
End Function

Public Function IntersectCircles(ByVal Circle1 As Long, ByVal Circle2 As Long) As TwoPoints
Dim Rad1 As Double, Rad2 As Double
Dim C1 As OnePoint, C2 As OnePoint
Dim R1 As TwoPoints, R2 As TwoPoints, T As TwoPoints
Rad1 = GetCircleRadius(Circle1)
Rad2 = GetCircleRadius(Circle2)
C1 = GetCircleCenter(Circle1)
C2 = GetCircleCenter(Circle2)
T = GetIntersectionOfCircles(C1, Rad1, C2, Rad2)

If Figures(Circle1).FigureType = dsCircle_ArcCenterAndRadiusAndTwoPoints Then
    If T.P1.X <> EmptyVar And T.P1.Y <> EmptyVar Then
        If Not PointBelongsToFigure(T.P1.X, T.P1.Y, Circle1) Then
            T.P1.X = EmptyVar
            T.P1.Y = EmptyVar
        End If
    End If
    If T.P2.X <> EmptyVar And T.P2.Y <> EmptyVar Then
        If Not PointBelongsToFigure(T.P2.X, T.P2.Y, Circle1) Then
            T.P2.X = EmptyVar
            T.P2.Y = EmptyVar
        End If
    End If
End If

If Figures(Circle2).FigureType = dsCircle_ArcCenterAndRadiusAndTwoPoints Then
    If T.P1.X <> EmptyVar And T.P1.Y <> EmptyVar Then
        If Not PointBelongsToFigure(T.P1.X, T.P1.Y, Circle2) Then
            T.P1.X = EmptyVar
            T.P1.Y = EmptyVar
        End If
    End If
    If T.P2.X <> EmptyVar And T.P2.Y <> EmptyVar Then
        If Not PointBelongsToFigure(T.P2.X, T.P2.Y, Circle2) Then
            T.P2.X = EmptyVar
            T.P2.Y = EmptyVar
        End If
    End If
End If

IntersectCircles = T
End Function

Public Sub SwapTwoPoints(ByRef TP As TwoPoints)
Dim tempPoint As OnePoint
tempPoint = TP.P1
TP.P1 = TP.P2
TP.P2 = tempPoint
End Sub

Public Sub SwapTwoNumbers(ByRef TN As TwoNumbers)
Dim TempNum As Double
TempNum = TN.n1
TN.n1 = TN.n2
TN.n2 = TempNum
End Sub

Public Sub Swap(ByRef Arg1 As Variant, ByRef Arg2 As Variant)
Dim TVar As Variant
TVar = Arg1
Arg1 = Arg2
Arg2 = TVar
End Sub

Public Function GetCircleCenter(ByVal Circle1 As Long) As OnePoint
Select Case Figures(Circle1).FigureType
    Case dsCircle_CenterAndTwoPoints, dsCircle_ArcCenterAndRadiusAndTwoPoints
        GetCircleCenter.X = BasePoint(Figures(Circle1).Points(2)).X
        GetCircleCenter.Y = BasePoint(Figures(Circle1).Points(2)).Y
    Case dsCircle_CenterAndCircumPoint
        GetCircleCenter.X = BasePoint(Figures(Circle1).Points(0)).X
        GetCircleCenter.Y = BasePoint(Figures(Circle1).Points(0)).Y
    Case dsAnCircle
        GetCircleCenter = GetCircleCenterFromEquation(Figures(Circle1).AuxInfo(1), Figures(Circle1).AuxInfo(2), Figures(Circle1).AuxInfo(3))
End Select
End Function

Public Function GetCircleRadius(ByVal Circle1 As Long) As Double
If Figures(Circle1).FigureType = dsAnCircle Then
    GetCircleRadius = GetCircleRadiusFromEquation(Figures(Circle1).AuxInfo(1), Figures(Circle1).AuxInfo(2), Figures(Circle1).AuxInfo(3))
Else
    GetCircleRadius = Distance(BasePoint(Figures(Circle1).Points(0)).X, BasePoint(Figures(Circle1).Points(0)).Y, BasePoint(Figures(Circle1).Points(1)).X, BasePoint(Figures(Circle1).Points(1)).Y)
End If
End Function

Public Function GetCircleEquation(ByVal index As Long) As String
Dim CCent As OnePoint, Rad As Double
If IsCircle(index) Then
    CCent = GetCircleCenter(index)
    Rad = GetCircleRadius(index)
    GetCircleEquation = GetCircleEquationText(CCent, Rad)
End If
End Function

Public Function GetPointOnFigure(ByVal Figure1 As Long, ByVal X As Double, ByVal Y As Double) As OnePoint
On Local Error Resume Next
Dim P As OnePoint, R As TwoPoints, tB As Boolean, T As Double
Dim CCent As OnePoint, Rad As Double, Ang As Double, A1 As Double, A2 As Double, A3 As Double
P.X = EmptyVar
P.Y = EmptyVar

Select Case Figures(Figure1).FigureType
    Case dsSegment
        R = GetLineCoordinates(Figure1)
        P = GetPerpPoint(X, Y, R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
        If R.P1.X > R.P2.X Then
            T = R.P1.X
            R.P1.X = R.P2.X
            R.P2.X = T
        End If
        If R.P1.Y > R.P2.Y Then
            T = R.P1.Y
            R.P1.Y = R.P2.Y
            R.P2.Y = T
        End If
        If P.X < R.P1.X Or P.X > R.P2.X Or P.Y < R.P1.Y Or P.Y > R.P2.Y Then
            If P.X < R.P1.X Then P.X = R.P1.X
            If P.X > R.P2.X Then P.X = R.P2.X
            If P.Y < R.P1.Y Then P.Y = R.P1.Y
            If P.Y > R.P2.Y Then P.Y = R.P2.Y
        End If

    Case dsRay
        R = GetLineCoordinates(Figure1)
        P = GetPerpPoint(X, Y, R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
        If Not PointBelongsToFigure(P.X, P.Y, Figure1) Then
            R = GetLineCoordinatesAbsolute(Figure1)
            P.X = R.P1.X
            P.Y = R.P1.Y
        End If
        
    Case dsLine_2Points, dsLine_PointAndParallelLine, dsBisector, dsLine_PointAndPerpendicularLine, dsAnLineGeneral, dsAnLineCanonic, dsAnLineNormal, dsAnLineNormalPoint
        R = GetLineCoordinates(Figure1)
        P = GetPerpPoint(X, Y, R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
    
    Case dsCircle_CenterAndCircumPoint, dsCircle_CenterAndTwoPoints, dsAnCircle
        CCent = GetCircleCenter(Figure1)
        Ang = GetAngle(CCent.X, CCent.Y, X, Y)
        Rad = GetCircleRadius(Figure1)
        P.X = CCent.X + Rad * Cos(Ang)
        P.Y = CCent.Y - Rad * Sin(Ang)
        
    Case dsCircle_ArcCenterAndRadiusAndTwoPoints
        CCent = GetCircleCenter(Figure1)
        Ang = GetAngle(CCent.X, CCent.Y, X, Y)
        Rad = GetCircleRadius(Figure1)
        P.X = CCent.X + Rad * Cos(Ang)
        P.Y = CCent.Y - Rad * Sin(Ang)
        A1 = Figures(Figure1).AuxInfo(2)
        A2 = Figures(Figure1).AuxInfo(3)
        A3 = Ang
        If A2 < A1 Then
            If A3 > A1 Then
                P.X = BasePoint(Figures(Figure1).Points(5)).X
                P.Y = BasePoint(Figures(Figure1).Points(5)).Y
            End If
            If A3 < A2 Then
                P.X = BasePoint(Figures(Figure1).Points(6)).X
                P.Y = BasePoint(Figures(Figure1).Points(6)).Y
            End If
        Else
            If A3 > A1 And A3 < A2 Then
                P.X = BasePoint(Figures(Figure1).Points(5)).X
                P.Y = BasePoint(Figures(Figure1).Points(5)).Y
            End If
        End If
    
    Case dsDynamicLocus
        If BasePoint(Figures(Figure1).Points(0)).Locus > 0 Then
            If Locuses(BasePoint(Figures(Figure1).Points(0)).Locus).Dynamic Then
                P = GetPerpPointPolyline(X, Y, Locuses(BasePoint(Figures(Figure1).Points(0)).Locus).LocusPoints)
            End If
        End If
End Select

GetPointOnFigure = P
End Function

Public Sub RecalcAuxInfo(ByVal index As Long)
Dim R As TwoPoints, P As OnePoint, P1 As OnePoint, P2 As OnePoint, lpPoint As POINTAPI
Dim X As Double, Y As Double, Ang As Double, A1 As Double, A2 As Double, A3 As Double, tR As Double, Z As Long

Select Case Figures(index).FigureType
    Case dsSegment, dsRay, dsLine_2Points
        If BasePoint(Figures(index).Points(0)).Visible And BasePoint(Figures(index).Points(1)).Visible Then
            Figures(index).Visible = True
            R = GetLineCoordinates(index)
            Figures(index).AuxPoints(1) = R.P1
            Figures(index).AuxPoints(2) = R.P2
            ToPhysical Figures(index).AuxPoints(1).X, Figures(index).AuxPoints(1).Y
            ToPhysical Figures(index).AuxPoints(2).X, Figures(index).AuxPoints(2).Y
        Else
            Figures(index).Visible = False
        End If
        
    Case dsBisector
        If BasePoint(Figures(index).Points(0)).Visible And BasePoint(Figures(index).Points(1)).Visible And BasePoint(Figures(index).Points(2)).Visible Then
            Figures(index).Visible = True
            R = GetLineCoordinates(index)
            If R.P2.X = EmptyVar Then
                Figures(index).Visible = False
            Else
                Figures(index).AuxPoints(1) = R.P1
                Figures(index).AuxPoints(2) = R.P2
                ToPhysical Figures(index).AuxPoints(1).X, Figures(index).AuxPoints(1).Y
                ToPhysical Figures(index).AuxPoints(2).X, Figures(index).AuxPoints(2).Y
            End If
        Else
            Figures(index).Visible = False
        End If
    
    Case dsLine_PointAndParallelLine, dsLine_PointAndPerpendicularLine
        If BasePoint(Figures(index).Points(0)).Visible And Figures(Figures(index).Parents(0)).Visible Then
            Figures(index).Visible = True
            R = GetLineCoordinates(index)
            Figures(index).AuxPoints(1) = R.P1
            Figures(index).AuxPoints(2) = R.P2
            ToPhysical Figures(index).AuxPoints(1).X, Figures(index).AuxPoints(1).Y
            ToPhysical Figures(index).AuxPoints(2).X, Figures(index).AuxPoints(2).Y
        Else
            Figures(index).Visible = False
        End If
    
    Case dsCircle_CenterAndCircumPoint
        If BasePoint(Figures(index).Points(0)).Visible And BasePoint(Figures(index).Points(1)).Visible Then
            Figures(index).Visible = True
            Figures(index).AuxPoints(1).X = BasePoint(Figures(index).Points(0)).X
            Figures(index).AuxPoints(1).Y = BasePoint(Figures(index).Points(0)).Y
            Figures(index).AuxInfo(1) = Distance(Figures(index).AuxPoints(1).X, Figures(index).AuxPoints(1).Y, BasePoint(Figures(index).Points(1)).X, BasePoint(Figures(index).Points(1)).Y)
            ToPhysical Figures(index).AuxPoints(1).X, Figures(index).AuxPoints(1).Y
            ToPhysicalLength Figures(index).AuxInfo(1)
        Else
            Figures(index).Visible = False
        End If
    
    Case dsCircle_CenterAndTwoPoints
        If BasePoint(Figures(index).Points(0)).Visible And BasePoint(Figures(index).Points(1)).Visible And BasePoint(Figures(index).Points(2)).Visible Then
            Figures(index).Visible = True
            Figures(index).AuxPoints(1).X = BasePoint(Figures(index).Points(2)).X
            Figures(index).AuxPoints(1).Y = BasePoint(Figures(index).Points(2)).Y
            Figures(index).AuxInfo(1) = Distance(BasePoint(Figures(index).Points(0)).X, BasePoint(Figures(index).Points(0)).Y, BasePoint(Figures(index).Points(1)).X, BasePoint(Figures(index).Points(1)).Y)
            ToPhysical Figures(index).AuxPoints(1).X, Figures(index).AuxPoints(1).Y
            ToPhysicalLength Figures(index).AuxInfo(1)
        Else
            Figures(index).Visible = False
        End If
    
    Case dsCircle_ArcCenterAndRadiusAndTwoPoints
         If BasePoint(Figures(index).Points(1)).Visible And BasePoint(Figures(index).Points(2)).Visible And BasePoint(Figures(index).Points(3)).Visible And BasePoint(Figures(index).Points(4)).Visible And BasePoint(Figures(index).Points(0)).Visible Then
            Figures(index).Visible = True
            tR = Distance(BasePoint(Figures(index).Points(0)).X, BasePoint(Figures(index).Points(0)).Y, BasePoint(Figures(index).Points(1)).X, BasePoint(Figures(index).Points(1)).Y)
            A2 = GetAngle(BasePoint(Figures(index).Points(2)).X, BasePoint(Figures(index).Points(2)).Y, BasePoint(Figures(index).Points(3)).X, BasePoint(Figures(index).Points(3)).Y)
            A3 = GetAngle(BasePoint(Figures(index).Points(2)).X, BasePoint(Figures(index).Points(2)).Y, BasePoint(Figures(index).Points(4)).X, BasePoint(Figures(index).Points(4)).Y)
            X = BasePoint(Figures(index).Points(2)).X + tR * Cos(A2)
            Y = BasePoint(Figures(index).Points(2)).Y - tR * Sin(A2)
            MovePoint Figures(index).Points(5), X, Y
            X = BasePoint(Figures(index).Points(2)).X + tR * Cos(A3)
            Y = BasePoint(Figures(index).Points(2)).Y - tR * Sin(A3)
            MovePoint Figures(index).Points(6), X, Y
            Figures(index).AuxPoints(1).X = BasePoint(Figures(index).Points(2)).X
            Figures(index).AuxPoints(1).Y = BasePoint(Figures(index).Points(2)).Y
            Figures(index).AuxInfo(1) = tR
            Figures(index).AuxInfo(2) = A2
            Figures(index).AuxInfo(3) = A3
            ToPhysical Figures(index).AuxPoints(1).X, Figures(index).AuxPoints(1).Y
            ToPhysicalLength Figures(index).AuxInfo(1)
        Else
            Figures(index).Visible = False
        End If
        
    Case dsMiddlePoint
        If BasePoint(Figures(index).Points(1)).Visible And BasePoint(Figures(index).Points(2)).Visible Then
            BasePoint(Figures(index).Points(0)).Visible = True
            P1.X = BasePoint(Figures(index).Points(1)).X
            P1.Y = BasePoint(Figures(index).Points(1)).Y
            P2.X = BasePoint(Figures(index).Points(2)).X
            P2.Y = BasePoint(Figures(index).Points(2)).Y
            'P = GetMiddlePoint(P1, P2)
            P.X = (P1.X + P2.X) / 2
            P.Y = (P1.Y + P2.Y) / 2
            Figures(index).AuxPoints(1) = P
            MovePoint Figures(index).Points(0), P.X, P.Y
        Else
            BasePoint(Figures(index).Points(0)).Visible = False
        End If
    
    Case dsIntersect
        R = GetIntersectionPoints(Figures(index).Parents(0), Figures(index).Parents(1))
        If R.P1.X <> EmptyVar And R.P1.Y <> EmptyVar And Figures(Figures(index).Parents(0)).Visible And Figures(Figures(index).Parents(1)).Visible Then
            MovePoint Figures(index).Points(0), R.P1.X, R.P1.Y
            BasePoint(Figures(index).Points(0)).Visible = True
            Figures(index).AuxPoints(1) = R.P1
            ToPhysical Figures(index).AuxPoints(1).X, Figures(index).AuxPoints(1).Y
        Else
            BasePoint(Figures(index).Points(0)).Visible = False
        End If
        If R.P2.X <> EmptyVar And R.P2.Y <> EmptyVar And Figures(Figures(index).Parents(0)).Visible And Figures(Figures(index).Parents(1)).Visible Then
            MovePoint Figures(index).Points(1), R.P2.X, R.P2.Y
            BasePoint(Figures(index).Points(1)).Visible = True
            Figures(index).AuxPoints(2) = R.P2
            ToPhysical Figures(index).AuxPoints(2).X, Figures(index).AuxPoints(2).Y
        Else
            BasePoint(Figures(index).Points(1)).Visible = False
        End If
        
    Case dsSimmPoint
        If BasePoint(Figures(index).Points(1)).Visible And BasePoint(Figures(index).Points(2)).Visible Then
            BasePoint(Figures(index).Points(0)).Visible = True
            P1.X = BasePoint(Figures(index).Points(2)).X
            P1.Y = BasePoint(Figures(index).Points(2)).Y
            P2.X = BasePoint(Figures(index).Points(1)).X
            P2.Y = BasePoint(Figures(index).Points(1)).Y
            P = GetSimmPoint(P1, P2)
            Figures(index).AuxPoints(1) = P
            MovePoint Figures(index).Points(0), P.X, P.Y
        Else
            BasePoint(Figures(index).Points(0)).Visible = False
        End If
    
    Case dsSimmPointByLine
        If BasePoint(Figures(index).Points(1)).Visible And Figures(Figures(index).Parents(0)).Visible Then
            BasePoint(Figures(index).Points(0)).Visible = True
            R = GetLineCoordinates(Figures(index).Parents(0))
            P = GetPerpPoint(BasePoint(Figures(index).Points(1)).X, BasePoint(Figures(index).Points(1)).Y, R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
            P1.X = 2 * P.X - BasePoint(Figures(index).Points(1)).X
            P1.Y = 2 * P.Y - BasePoint(Figures(index).Points(1)).Y
            Figures(index).AuxPoints(1) = P1
            MovePoint Figures(index).Points(0), P1.X, P1.Y
        Else
            BasePoint(Figures(index).Points(0)).Visible = False
        End If

    Case dsPointOnFigure
        If ManualDragFlag = Figures(index).Points(0) Then
            Exit Sub
        End If
        P = RestoreSemiDependentPoint(index)
        If P.X <> EmptyVar And P.Y <> EmptyVar Then
            BasePoint(Figures(index).Points(0)).Visible = True
            MovePoint Figures(index).Points(0), P.X, P.Y
            Figures(index).AuxPoints(1) = P
            ToPhysical Figures(index).AuxPoints(1).X, Figures(index).AuxPoints(1).Y
        Else
            BasePoint(Figures(index).Points(0)).Visible = False
        End If

    Case dsMeasureDistance
        If BasePoint(Figures(index).Points(0)).Visible And BasePoint(Figures(index).Points(1)).Visible Then
            Figures(index).Visible = True
            Figures(index).AuxPoints(1).X = BasePoint(Figures(index).Points(0)).X
            Figures(index).AuxPoints(1).Y = BasePoint(Figures(index).Points(0)).Y
            Figures(index).AuxPoints(2).X = BasePoint(Figures(index).Points(1)).X
            Figures(index).AuxPoints(2).Y = BasePoint(Figures(index).Points(1)).Y
            Figures(index).AuxInfo(1) = Distance(Figures(index).AuxPoints(1).X, Figures(index).AuxPoints(1).Y, Figures(index).AuxPoints(2).X, Figures(index).AuxPoints(2).Y)
            Figures(index).AuxPoints(3).X = (Figures(index).AuxPoints(1).X + Figures(index).AuxPoints(2).X) / 2
            Figures(index).AuxPoints(3).Y = (Figures(index).AuxPoints(1).Y + Figures(index).AuxPoints(2).Y) / 2
            
            Figures(index).XS = Format(Figures(index).AuxInfo(1), setFormatDistance)
            
            Const AngleText As Boolean = True
            If AngleText Then
                If Figures(index).AuxPoints(1).X <> Figures(index).AuxPoints(2).X Then
                    Figures(index).AuxInfo(2) = Atn((Figures(index).AuxPoints(2).Y - Figures(index).AuxPoints(1).Y) / (Figures(index).AuxPoints(2).X - Figures(index).AuxPoints(1).X)) * ToDegrees
                Else
                    Figures(index).AuxInfo(2) = -90
                End If
            End If
            Figures(index).AuxPoints(5) = RecalcMeasure(Paper.hDC, Figures(index).XS, Figures(index).AuxInfo(2))
            ToPhysical Figures(index).AuxPoints(1).X, Figures(index).AuxPoints(1).Y
            ToPhysical Figures(index).AuxPoints(2).X, Figures(index).AuxPoints(2).Y
            ToPhysical Figures(index).AuxPoints(3).X, Figures(index).AuxPoints(3).Y
            A3 = AngleMarkDist
            ToPhysicalLength A3
            A3 = A3 * 2
            Figures(index).AuxPoints(6).X = Figures(index).AuxPoints(6).X + Figures(index).AuxPoints(3).X
            Figures(index).AuxPoints(6).Y = Figures(index).AuxPoints(6).Y + Figures(index).AuxPoints(3).Y
            LinkPointToSegment Figures(index).AuxPoints(6).X, Figures(index).AuxPoints(6).Y, Figures(index).AuxPoints(1).X, Figures(index).AuxPoints(1).Y, Figures(index).AuxPoints(2).X, Figures(index).AuxPoints(2).Y, A3
            Figures(index).AuxPoints(6).X = Figures(index).AuxPoints(6).X - Figures(index).AuxPoints(3).X
            Figures(index).AuxPoints(6).Y = Figures(index).AuxPoints(6).Y - Figures(index).AuxPoints(3).Y
        Else
            Figures(index).Visible = False
        End If
    
    Case dsMeasureAngle
        If BasePoint(Figures(index).Points(0)).Visible And BasePoint(Figures(index).Points(1)).Visible And BasePoint(Figures(index).Points(2)).Visible Then
            Figures(index).Visible = True
            Figures(index).AuxPoints(1).X = (BasePoint(Figures(index).Points(0)).X)
            Figures(index).AuxPoints(1).Y = (BasePoint(Figures(index).Points(0)).Y)
            Figures(index).AuxPoints(2).X = (BasePoint(Figures(index).Points(1)).X)
            Figures(index).AuxPoints(2).Y = (BasePoint(Figures(index).Points(1)).Y)
            Figures(index).AuxPoints(3).X = (BasePoint(Figures(index).Points(2)).X)
            Figures(index).AuxPoints(3).Y = (BasePoint(Figures(index).Points(2)).Y)
            Figures(index).AuxInfo(1) = Angle(Figures(index).AuxPoints(1).X, Figures(index).AuxPoints(1).Y, Figures(index).AuxPoints(2).X, Figures(index).AuxPoints(2).Y, Figures(index).AuxPoints(3).X, Figures(index).AuxPoints(3).Y) * ToDegrees
            
            Figures(index).XS = Format(Figures(index).AuxInfo(1), setFormatAngle) & "°"
            
            'Const AngRad = 0.5
            A1 = GetAngle(Figures(index).AuxPoints(2).X, Figures(index).AuxPoints(2).Y, Figures(index).AuxPoints(1).X, Figures(index).AuxPoints(1).Y)
            A2 = GetAngle(Figures(index).AuxPoints(2).X, Figures(index).AuxPoints(2).Y, Figures(index).AuxPoints(3).X, Figures(index).AuxPoints(3).Y)
            If A2 < A1 Then Swap A1, A2
            If A2 - A1 > PI Then Ang = Figures(index).AuxInfo(1) * ToRadians / 2 + A2 Else Ang = Figures(index).AuxInfo(1) * ToRadians / 2 + A1
            Figures(index).AuxPoints(4).X = Figures(index).AuxPoints(2).X + Cos(Ang) * AngleMarkDist
            Figures(index).AuxPoints(4).Y = Figures(index).AuxPoints(2).Y - Sin(Ang) * AngleMarkDist
            Figures(index).AuxPoints(5) = RecalcMeasure(Paper.hDC, Figures(index).XS)
            
            ToPhysical Figures(index).AuxPoints(1).X, Figures(index).AuxPoints(1).Y
            ToPhysical Figures(index).AuxPoints(2).X, Figures(index).AuxPoints(2).Y
            ToPhysical Figures(index).AuxPoints(3).X, Figures(index).AuxPoints(3).Y
    '
            Figures(index).AuxPoints(1).X = Round(Figures(index).AuxPoints(1).X)
            Figures(index).AuxPoints(1).Y = Round(Figures(index).AuxPoints(1).Y)
            Figures(index).AuxPoints(2).X = Round(Figures(index).AuxPoints(2).X)
            Figures(index).AuxPoints(2).Y = Round(Figures(index).AuxPoints(2).Y)
            Figures(index).AuxPoints(3).X = Round(Figures(index).AuxPoints(3).X)
            Figures(index).AuxPoints(3).Y = Round(Figures(index).AuxPoints(3).Y)
            
            ToPhysical Figures(index).AuxPoints(4).X, Figures(index).AuxPoints(4).Y
            A3 = AngleMarkDist
            ToPhysicalLength A3
            A3 = A3 * 2
            Figures(index).AuxPoints(6).X = Figures(index).AuxPoints(6).X + Figures(index).AuxPoints(4).X
            Figures(index).AuxPoints(6).Y = Figures(index).AuxPoints(6).Y + Figures(index).AuxPoints(4).Y
            LinkPointToPoint Figures(index).AuxPoints(6).X, Figures(index).AuxPoints(6).Y, Figures(index).AuxPoints(4).X, Figures(index).AuxPoints(4).Y, A3
            Figures(index).AuxPoints(6).X = Figures(index).AuxPoints(6).X - Figures(index).AuxPoints(4).X
            Figures(index).AuxPoints(6).Y = Figures(index).AuxPoints(6).Y - Figures(index).AuxPoints(4).Y
        Else
            Figures(index).Visible = False
        End If
        
    Case dsAnLineGeneral, dsAnLineCanonic, dsAnLineNormal, dsAnLineNormalPoint
        R = GetLineFromSegment(Figures(index).AuxPoints(3).X, Figures(index).AuxPoints(3).Y, Figures(index).AuxPoints(4).X, Figures(index).AuxPoints(4).Y)
        Figures(index).AuxPoints(1) = R.P1
        Figures(index).AuxPoints(2) = R.P2
        ToPhysical Figures(index).AuxPoints(1).X, Figures(index).AuxPoints(1).Y
        ToPhysical Figures(index).AuxPoints(2).X, Figures(index).AuxPoints(2).Y
        
'    Case dsAnLineCanonic
'        R = GetCanonicAnLine(Figures(Index).AuxInfo(1), Figures(Index).AuxInfo(2), Figures(Index).AuxInfo(3), Figures(Index).AuxInfo(4))
'        Figures(Index).AuxPoints(1) = R.P1
'        Figures(Index).AuxPoints(2) = R.P2
'        ToPhysical Figures(Index).AuxPoints(1).X, Figures(Index).AuxPoints(1).Y
'        ToPhysical Figures(Index).AuxPoints(2).X, Figures(Index).AuxPoints(2).Y
'
'    Case dsAnLineNormal
'        R = GetNormalAnLine(Figures(Index).AuxInfo(1), Figures(Index).AuxInfo(2))
'        Figures(Index).AuxPoints(1) = R.P1
'        Figures(Index).AuxPoints(2) = R.P2
'        ToPhysical Figures(Index).AuxPoints(1).X, Figures(Index).AuxPoints(1).Y
'        ToPhysical Figures(Index).AuxPoints(2).X, Figures(Index).AuxPoints(2).Y
'
'    Case dsAnLineNormalPoint
'        R = GetNormalPointAnLine(Figures(Index).AuxInfo(1), Figures(Index).AuxInfo(2), Figures(Index).AuxInfo(3), Figures(Index).AuxInfo(4))
'        Figures(Index).AuxPoints(1) = R.P1
'        Figures(Index).AuxPoints(2) = R.P2
'        ToPhysical Figures(Index).AuxPoints(1).X, Figures(Index).AuxPoints(1).Y
'        ToPhysical Figures(Index).AuxPoints(2).X, Figures(Index).AuxPoints(2).Y
    
    Case dsAnCircle
        P = GetCircleCenter(index)
        Figures(index).AuxPoints(1).X = P.X
        Figures(index).AuxPoints(1).Y = P.Y
        Figures(index).AuxInfo(4) = GetCircleRadius(index)
        ToPhysical Figures(index).AuxPoints(1).X, Figures(index).AuxPoints(1).Y
        ToPhysicalLength Figures(index).AuxInfo(4)
        
    Case dsInvert
        If Figures(Figures(index).Parents(0)).Visible And BasePoint(Figures(index).Points(1)).Visible Then
            P = GetCircleCenter(Figures(index).Parents(0))
            P = GetInvertedPoint(BasePoint(Figures(index).Points(1)).X, BasePoint(Figures(index).Points(1)).Y, P.X, P.Y, GetCircleRadius(Figures(index).Parents(0)))
            If P.X <> EmptyVar And P.Y <> EmptyVar Then
                BasePoint(Figures(index).Points(0)).Visible = True
                MovePoint Figures(index).Points(0), P.X, P.Y
                Figures(index).AuxPoints(1) = P
                ToPhysical Figures(index).AuxPoints(1).X, Figures(index).AuxPoints(1).Y
            Else
                BasePoint(Figures(index).Points(0)).Visible = False
            End If
        Else
            BasePoint(Figures(index).Points(0)).Visible = False
        End If
    
    Case dsAnPoint
        With Figures(index)
            For Z = 1 To .NumberOfPoints - 1
                If Not BasePoint(.Points(Z)).Visible Then
                    BasePoint(.Points(0)).Visible = False
                    Exit Sub
                End If
            Next
            RecalculateTree .XTree
            RecalculateTree .YTree
            If WasThereAnErrorEvaluatingLastExpression <> eetEverythingOK Then
                WasThereAnErrorEvaluatingLastExpression = eetEverythingOK
                BasePoint(.Points(0)).Visible = False
            Else
                BasePoint(.Points(0)).Visible = True
                P.X = .XTree.Branches(1).CurrentValue
                P.Y = .YTree.Branches(1).CurrentValue
                .AuxPoints(1) = P
                MovePoint .Points(0), P.X, P.Y
            End If
        End With
        
    Case dsDynamicLocus
        Locuses(BasePoint(Figures(index).Points(0)).Locus).Visible = True
        If BasePoint(Figures(index).Points(0)).Visible = False Then Locuses(BasePoint(Figures(index).Points(0)).Locus).Visible = False: Exit Sub
        If ManualDragFlag = Figures(index).Points(1) Then Exit Sub
        RecalcDynamicLocus index
End Select
End Sub

Public Sub RecalcPureAuxInfo(ByVal index As Long)
Dim R As TwoPoints, P As OnePoint, P1 As OnePoint, P2 As OnePoint
Dim X As Double, Y As Double, Ang As Double, A1 As Double, A2 As Double, A3 As Double, tR As Double, Z As Long

Select Case Figures(index).FigureType
'    Case dsSegment, dsRay, dsLine_2Points
'        R = GetLineCoordinates(Index)
'        Figures(Index).AuxPoints(1) = R.P1
'        Figures(Index).AuxPoints(2) = R.P2
    
'    Case dsLine_PointAndParallelLine, dsLine_PointAndPerpendicularLine
'        R = GetLineCoordinates(Index)
'        Figures(Index).AuxPoints(1) = R.P1
'        Figures(Index).AuxPoints(2) = R.P2
    
'    Case dsCircle_CenterAndCircumPoint
'        Figures(Index).AuxPoints(1).X = BasePoint(Figures(Index).Points(0)).X
'        Figures(Index).AuxPoints(1).Y = BasePoint(Figures(Index).Points(0)).Y
'        Figures(Index).AuxInfo(1) = Distance(Figures(Index).AuxPoints(1).X, Figures(Index).AuxPoints(1).Y, BasePoint(Figures(Index).Points(1)).X, BasePoint(Figures(Index).Points(1)).Y)
    
'    Case dsCircle_CenterAndTwoPoints
'        Figures(Index).AuxPoints(1).X = BasePoint(Figures(Index).Points(2)).X
'        Figures(Index).AuxPoints(1).Y = BasePoint(Figures(Index).Points(2)).Y
'        Figures(Index).AuxInfo(1) = Distance(BasePoint(Figures(Index).Points(0)).X, BasePoint(Figures(Index).Points(0)).Y, BasePoint(Figures(Index).Points(1)).X, BasePoint(Figures(Index).Points(1)).Y)
    
        
    Case dsMiddlePoint
        P1.X = BasePoint(Figures(index).Points(1)).X
        P1.Y = BasePoint(Figures(index).Points(1)).Y
        P2.X = BasePoint(Figures(index).Points(2)).X
        P2.Y = BasePoint(Figures(index).Points(2)).Y
        Figures(index).AuxPoints(1).X = (P1.X + P2.X) / 2
        Figures(index).AuxPoints(1).Y = (P1.Y + P2.Y) / 2
        MovePointPure Figures(index).Points(0), Figures(index).AuxPoints(1).X, Figures(index).AuxPoints(1).Y
        BasePoint(Figures(index).Points(0)).Visible = BasePoint(Figures(index).Points(1)).Visible And BasePoint(Figures(index).Points(2)).Visible
    
    Case dsIntersect
        R = GetIntersectionPoints(Figures(index).Parents(0), Figures(index).Parents(1))
        If R.P1.X <> EmptyVar And R.P1.Y <> EmptyVar And Figures(Figures(index).Parents(0)).Visible And Figures(Figures(index).Parents(1)).Visible Then
            MovePointPure Figures(index).Points(0), R.P1.X, R.P1.Y
            BasePoint(Figures(index).Points(0)).Visible = True
            Figures(index).AuxPoints(1) = R.P1
            'ToPhysical Figures(Index).AuxPoints(1).X, Figures(Index).AuxPoints(1).Y
        Else
            BasePoint(Figures(index).Points(0)).Visible = False
        End If
        If R.P2.X <> EmptyVar And R.P2.Y <> EmptyVar And Figures(Figures(index).Parents(0)).Visible And Figures(Figures(index).Parents(1)).Visible Then
            MovePointPure Figures(index).Points(1), R.P2.X, R.P2.Y
            BasePoint(Figures(index).Points(1)).Visible = True
            Figures(index).AuxPoints(2) = R.P2
            'ToPhysical Figures(Index).AuxPoints(2).X, Figures(Index).AuxPoints(2).Y
        Else
            BasePoint(Figures(index).Points(1)).Visible = False
        End If
        
    Case dsSimmPoint
        If BasePoint(Figures(index).Points(1)).Visible And BasePoint(Figures(index).Points(2)).Visible Then
            BasePoint(Figures(index).Points(0)).Visible = True
            P1.X = BasePoint(Figures(index).Points(2)).X
            P1.Y = BasePoint(Figures(index).Points(2)).Y
            P2.X = BasePoint(Figures(index).Points(1)).X
            P2.Y = BasePoint(Figures(index).Points(1)).Y
            Figures(index).AuxPoints(1).X = 2 * P1.X - P2.X
            Figures(index).AuxPoints(1).Y = 2 * P1.Y - P2.Y
            MovePointPure Figures(index).Points(0), Figures(index).AuxPoints(1).X, Figures(index).AuxPoints(1).Y
        Else
            BasePoint(Figures(index).Points(0)).Visible = False
        End If
    
    Case dsPointOnFigure
        P = RestoreSemiDependentPoint(index)
        If P.X <> EmptyVar And P.Y <> EmptyVar Then
            BasePoint(Figures(index).Points(0)).Visible = True
            MovePointPure Figures(index).Points(0), P.X, P.Y
            Figures(index).AuxPoints(1) = P
            'ToPhysical Figures(Index).AuxPoints(1).X, Figures(Index).AuxPoints(1).Y
        Else
            BasePoint(Figures(index).Points(0)).Visible = False
        End If
        
    Case dsInvert
        If Figures(Figures(index).Parents(0)).Visible And BasePoint(Figures(index).Points(1)).Visible Then
            P = GetCircleCenter(Figures(index).Parents(0))
            P = GetInvertedPoint(BasePoint(Figures(index).Points(1)).X, BasePoint(Figures(index).Points(1)).Y, P.X, P.Y, GetCircleRadius(Figures(index).Parents(0)))
            If P.X <> EmptyVar And P.Y <> EmptyVar Then
                BasePoint(Figures(index).Points(0)).Visible = True
                MovePointPure Figures(index).Points(0), P.X, P.Y
                Figures(index).AuxPoints(1) = P
                'ToPhysical Figures(Index).AuxPoints(1).X, Figures(Index).AuxPoints(1).Y
            Else
                BasePoint(Figures(index).Points(0)).Visible = False
            End If
        Else
            BasePoint(Figures(index).Points(0)).Visible = False
        End If
    
    Case dsAnPoint
        With Figures(index)
            For Z = 1 To .NumberOfPoints - 1
                If Not BasePoint(.Points(Z)).Visible Then
                    BasePoint(.Points(0)).Visible = False
                    Exit Sub
                End If
            Next
            RecalculateTree .XTree
            RecalculateTree .YTree
            If WasThereAnErrorEvaluatingLastExpression <> eetEverythingOK Then
                WasThereAnErrorEvaluatingLastExpression = eetEverythingOK
                BasePoint(.Points(0)).Visible = False
            Else
                BasePoint(.Points(0)).Visible = True
                P.X = .XTree.Branches(1).CurrentValue
                P.Y = .YTree.Branches(1).CurrentValue
                .AuxPoints(1) = P
                MovePointPure .Points(0), P.X, P.Y
            End If
        End With
        
    Case dsDynamicLocus
        Locuses(BasePoint(Figures(index).Points(0)).Locus).Visible = True
        'If BasePoint(Figures(Index).Points(0)).Visible = False Then Locuses(BasePoint(Figures(Index).Points(0)).Locus).Visible = False: Exit Sub
        RecalcDynamicLocus index
    
    Case dsCircle_ArcCenterAndRadiusAndTwoPoints
        tR = Distance(BasePoint(Figures(index).Points(0)).X, BasePoint(Figures(index).Points(0)).Y, BasePoint(Figures(index).Points(1)).X, BasePoint(Figures(index).Points(1)).Y)
        A2 = GetAngle(BasePoint(Figures(index).Points(2)).X, BasePoint(Figures(index).Points(2)).Y, BasePoint(Figures(index).Points(3)).X, BasePoint(Figures(index).Points(3)).Y)
        A3 = GetAngle(BasePoint(Figures(index).Points(2)).X, BasePoint(Figures(index).Points(2)).Y, BasePoint(Figures(index).Points(4)).X, BasePoint(Figures(index).Points(4)).Y)
        X = BasePoint(Figures(index).Points(2)).X + tR * Cos(A2)
        Y = BasePoint(Figures(index).Points(2)).Y - tR * Sin(A2)
        MovePointPure Figures(index).Points(5), X, Y
        X = BasePoint(Figures(index).Points(2)).X + tR * Cos(A3)
        Y = BasePoint(Figures(index).Points(2)).Y - tR * Sin(A3)
        MovePointPure Figures(index).Points(6), X, Y
        Figures(index).AuxPoints(1).X = BasePoint(Figures(index).Points(2)).X
        Figures(index).AuxPoints(1).Y = BasePoint(Figures(index).Points(2)).Y
        Figures(index).AuxInfo(1) = tR
        Figures(index).AuxInfo(2) = A2
        Figures(index).AuxInfo(3) = A3
    
    Case dsSimmPointByLine
        If BasePoint(Figures(index).Points(1)).Visible And Figures(Figures(index).Parents(0)).Visible Then
            BasePoint(Figures(index).Points(0)).Visible = True
            R = GetLineCoordinates(Figures(index).Parents(0))
            P = GetPerpPoint(BasePoint(Figures(index).Points(1)).X, BasePoint(Figures(index).Points(1)).Y, R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
            P1.X = 2 * P.X - BasePoint(Figures(index).Points(1)).X
            P1.Y = 2 * P.Y - BasePoint(Figures(index).Points(1)).Y
            Figures(index).AuxPoints(1) = P1
            MovePointPure Figures(index).Points(0), P1.X, P1.Y
        Else
            BasePoint(Figures(index).Points(0)).Visible = False
        End If

End Select
End Sub

Public Sub RecalcAllAuxInfo()
Dim Z As Long
For Z = 0 To FigureCount - 1
    RecalcAuxInfo Z
    If Figures(Z).FigureType = dsPointOnFigure And Figures(Z).AuxInfo(5) = 1 Then RecalcFigureWithPoint (Figures(Z).Points(0))
Next
RecalcStaticGraphics
End Sub

Public Sub RecalcSemiDependentInfo(ByVal Figure1 As Long, ByVal X As Single, ByVal Y As Single)
On Local Error Resume Next
Dim R As TwoPoints, CC As OnePoint, P As OnePoint, a As Double, b As Double

If IsLine(Figures(Figure1).Parents(0)) Then
    R = GetLineCoordinatesAbsolute(Figures(Figure1).Parents(0))
    P = GetPerpPoint(X, Y, R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
    If R.P1.X <> R.P2.X Then
        If R.P1.Y <> R.P2.Y Then
            Figures(Figure1).AuxInfo(1) = (P.X - R.P1.X) / (R.P2.X - R.P1.X)
        Else
            Figures(Figure1).AuxInfo(1) = (P.X - R.P1.X) / (R.P2.X - R.P1.X)
        End If
    Else
        Figures(Figure1).AuxInfo(1) = (P.Y - R.P1.Y) / (R.P2.Y - R.P1.Y)
    End If
    If Figures(Figures(Figure1).Parents(0)).FigureType = dsSegment Then
        If Figures(Figure1).AuxInfo(1) < 0 Then Figures(Figure1).AuxInfo(1) = 0
        If Figures(Figure1).AuxInfo(1) > 1 Then Figures(Figure1).AuxInfo(1) = 1
    End If
    If Figures(Figures(Figure1).Parents(0)).FigureType = dsRay Then
        If Figures(Figure1).AuxInfo(1) < 0 Then Figures(Figure1).AuxInfo(1) = 0
    End If
End If

If IsCircle(Figures(Figure1).Parents(0)) Then
    CC = GetCircleCenter(Figures(Figure1).Parents(0))
    Figures(Figure1).AuxInfo(1) = GetAngle(CC.X, CC.Y, X, Y)
    If Figures(Figures(Figure1).Parents(0)).FigureType = dsCircle_ArcCenterAndRadiusAndTwoPoints Then
        a = Figures(Figures(Figure1).Parents(0)).AuxInfo(2)
        b = Figures(Figures(Figure1).Parents(0)).AuxInfo(3)
        If b > a Then
            If Figures(Figure1).AuxInfo(1) < a Then Figures(Figure1).AuxInfo(1) = a
            If Figures(Figure1).AuxInfo(1) > b Then Figures(Figure1).AuxInfo(1) = b
        Else
            If Figures(Figure1).AuxInfo(1) > a Then Figures(Figure1).AuxInfo(1) = a
            If Figures(Figure1).AuxInfo(1) < b Then Figures(Figure1).AuxInfo(1) = b
        End If
    End If
End If

If Figures(Figures(Figure1).Parents(0)).FigureType = dsDynamicLocus Then
    a = Figures(Figures(Figure1).Parents(0)).AuxInfo(2)
    b = Figures(Figures(Figure1).Parents(0)).AuxInfo(3)
    Figures(Figure1).AuxInfo(1) = a + (b - a) * GetNearestSegmentNum(X, Y, Locuses(BasePoint(Figures(Figures(Figure1).Parents(0)).Points(0)).Locus).LocusPoints) / Locuses(BasePoint(Figures(Figures(Figure1).Parents(0)).Points(0)).Locus).LocusPointCount
End If
End Sub

Public Function RestoreSemiDependentPoint(ByVal Figure1 As Long) As OnePoint
Dim P As OnePoint, CC As OnePoint, R As TwoPoints, Rad As Double, Ang As Double
Dim FigurePoint As Long, Carrier As Long, CurLocus As Long, OldManualDragFlag As Long
Dim A1 As Double, A2 As Double, OldT As Double, Z As Long
P.X = EmptyVar
P.Y = EmptyVar

If IsLine(Figures(Figure1).Parents(0)) Then
    R = GetLineCoordinatesAbsolute(Figures(Figure1).Parents(0))
    P.X = R.P1.X + (R.P2.X - R.P1.X) * Figures(Figure1).AuxInfo(1)
    P.Y = R.P1.Y + (R.P2.Y - R.P1.Y) * Figures(Figure1).AuxInfo(1)
    With Figures(Figures(Figure1).Parents(0))
        If Not .Visible Then P = EmptyOnePoint
'        If Not PointBelongsToFigure(P.X, P.Y, Figures(Figure1).Parents(0)) Then
'            Figures(Figure1).AuxInfo(1) = 0.5
'            P.X = R.P1.X + (R.P2.X - R.P1.X) * Figures(Figure1).AuxInfo(1)
'            P.Y = R.P1.Y + (R.P2.Y - R.P1.Y) * Figures(Figure1).AuxInfo(1)
'        End If
'        If Figures(Figures(Figure1).Parents(0)).FigureType = dsSegment Then
'            If Figures(Figure1).AuxInfo(1) < 0 Then Figures(Figure1).AuxInfo(1) = 0
'            If Figures(Figure1).AuxInfo(1) > 1 Then Figures(Figure1).AuxInfo(1) = 1
'            P.X = R.P1.X + (R.P2.X - R.P1.X) * Figures(Figure1).AuxInfo(1)
'            P.Y = R.P1.Y + (R.P2.Y - R.P1.Y) * Figures(Figure1).AuxInfo(1)
'        End If
        For Z = 0 To .NumberOfPoints - 1
            If BasePoint(.Points(Z)).Visible = False And .FigureType <> dsIntersect And .FigureType <> dsPointOnFigure Then P = EmptyOnePoint
        Next Z
    End With
End If

If IsCircle(Figures(Figure1).Parents(0)) Then
    CC = GetCircleCenter(Figures(Figure1).Parents(0))
    Rad = GetCircleRadius(Figures(Figure1).Parents(0))
    Ang = Figures(Figure1).AuxInfo(1)
    P.X = CC.X + Rad * Cos(Ang)
    P.Y = CC.Y - Rad * Sin(Ang)
    If Figures(Figures(Figure1).Parents(0)).FigureType = dsCircle_ArcCenterAndRadiusAndTwoPoints Then
        A1 = Figures(Figures(Figure1).Parents(0)).AuxInfo(2)
        A2 = Figures(Figures(Figure1).Parents(0)).AuxInfo(3)
        If A2 < A1 Then
            If Not (Ang < A1 And Ang > A2) Then
                Figures(Figure1).AuxInfo(1) = (A1 + A2) / 2
                Ang = Figures(Figure1).AuxInfo(1)
                P.X = CC.X + Rad * Cos(Ang)
                P.Y = CC.Y - Rad * Sin(Ang)
            End If
        Else
            If Not (Ang < A1 Or Ang > A2) Then
                Figures(Figure1).AuxInfo(1) = (A1 + A2) / 2 + PI
                Ang = Figures(Figure1).AuxInfo(1)
                P.X = CC.X + Rad * Cos(Ang)
                P.Y = CC.Y - Rad * Sin(Ang)
            End If
        End If
    End If
End If

If Figures(Figures(Figure1).Parents(0)).FigureType = dsDynamicLocus Then
    FigurePoint = BasePoint(Figures(Figures(Figure1).Parents(0)).Points(1)).ParentFigure
    Carrier = Figures(FigurePoint).Parents(0)
    CurLocus = BasePoint(Figures(Figures(Figure1).Parents(0)).Points(0)).Locus
    
    If Figures(Carrier).Visible = False Or Locuses(CurLocus).Visible = False Then RestoreSemiDependentPoint = EmptyOnePoint: Exit Function
    
    OldManualDragFlag = ManualDragFlag
    ManualDragFlag = 0
    OldT = Figures(FigurePoint).AuxInfo(1)
    
    Figures(FigurePoint).AuxInfo(1) = Figures(Figure1).AuxInfo(1)
    For Z = 1 To Figures(Figures(Figure1).Parents(0)).AuxInfo(1)
        RecalcPureAuxInfo Figures(Figures(Figure1).Parents(0)).AuxArray(Z)
    Next
    
    P.X = BasePoint(Figures(Figures(Figure1).Parents(0)).Points(0)).X
    P.Y = BasePoint(Figures(Figures(Figure1).Parents(0)).Points(0)).Y
    If Abs(P.X) > 1000 Or Abs(P.Y) > 1000 Or P.X = EmptyVar Or P.Y = EmptyVar Then P = EmptyOnePoint
    
    Figures(FigurePoint).AuxInfo(1) = OldT
    'RecalcAuxInfo FigurePoint
    For Z = 1 To Figures(Figures(Figure1).Parents(0)).AuxInfo(1)
        RecalcPureAuxInfo Figures(Figures(Figure1).Parents(0)).AuxArray(Z)
    Next
    'Figures(FigurePoint).AuxInfo(1) = OldT
    'RecalcAuxInfo FigurePoint
    
    ManualDragFlag = OldManualDragFlag
End If

RestoreSemiDependentPoint = P
End Function



Public Function FormatLineGeneralEquation(Eq As LineGeneralEquation) As String
Const Prec1 = 10, Prec2 = 100
Dim S As String, a As Double, b As Double, c As Double, Min As Double, F1 As Long
a = Eq.a
b = Eq.b
c = Eq.c
If (a <= 0 And b <= 0 And c <= 0) Or (a <= 0 And b <= 0 And c >= 0) Or (a <= 0 And b >= 0 And c <= 0) Or (a <= 0 And b >= 0 And c >= 0) Then a = -a: b = -b: c = -c
If a = 0 And c = 0 And b <> 0 Then b = 1
If a <> 0 And b = 0 And c = 0 Then a = 1
If a = 0 And b = 0 And c = 0 Then
    FormatLineGeneralEquation = "x ^ 2 + y ^ 2 = -1"
    Exit Function
End If
Eq.a = a
Eq.b = b
Eq.c = c

If Abs(a) > Abs(b) Then Min = b Else Min = a
If Min = 0 Then If a = 0 Then Min = b Else Min = a
If c <> 0 Then If Abs(Min) > Abs(c) Then Min = c
Min = Abs(Min)
a = Round(a / Min, Prec1)
b = Round(b / Min, Prec1)
c = Round(c / Min, Prec1)
F1 = 1
Do While (F1 * a <> Round(F1 * a) Or F1 * b <> Round(F1 * b) Or F1 * c <> Round(F1 * c)) And F1 < Prec2
    F1 = F1 + 1
Loop
If F1 < Prec2 - 2 Then
    Eq.a = a * F1
    Eq.b = b * F1
    Eq.c = c * F1
End If

If Eq.a <> 0 Then If Abs(Eq.a) <> 1 Then S = Round(Eq.a, setNumberPrecision) & "x" Else S = IIf(Eq.a = 1, "", "-") & "x"
If Eq.b <> 0 Then
    If Eq.a <> 0 Then
        If Abs(Eq.b) <> 1 Then
            S = S & " " & IIf(Eq.b > 0, "+ ", "- ") & Round(Abs(Eq.b), setNumberPrecision) & "y"
        Else
            S = S & " " & IIf(Eq.b = 1, "+ y", "- y")
        End If
    Else
        If Abs(Eq.b) <> 1 Then S = S & " " & Round(Eq.b, setNumberPrecision) & "y" Else S = S & " " & IIf(Eq.b = 1, "", "-") & "y"
    End If
End If
If Eq.c <> 0 Then
    S = S & " " & IIf(Eq.c > 0, "+ ", "- ") & Round(Abs(Eq.c), setNumberPrecision)
End If
S = S & " = 0"
FormatLineGeneralEquation = Trim(S)
End Function

Public Function GenerateNewPointName(Optional ByVal ShouldAllocate As Boolean = True) As String
Dim tName As String

tName = IncrementPointName(LastAllocatedPointName)
Do While GetPointByName(tName) > 0
    tName = IncrementPointName(tName)
Loop
GenerateNewPointName = tName
If ShouldAllocate Then LastAllocatedPointName = tName
'Dim TNum As Long, tName As String
'Do
'    If TNum < 26 Then tName = Chr(65 + TNum) Else tName = Chr(65 + (TNum Mod 26)) & (TNum \ 26)
'    TNum = TNum + 1
'Loop Until PointNames.FindItem(tName) = 0
'GenerateNewPointName = tName
End Function

Public Function IncrementPointName(ByVal S As String) As String
Dim FL As String, LS As String, Z As Long

If S = "" Or S = "_" Then IncrementPointName = "A": Exit Function
FL = Left(S, 1)
LS = Right(S, Len(S) - 1)

If S = FL And IsAlpha(FL) Then
    If FL < "Z" Then FL = Chr(Asc(FL) + 1) Else FL = "A1"
    IncrementPointName = FL
    Exit Function
End If

If IsAlpha(FL) And IsNumeric(LS) Then
    If FL < "Z" Then
        IncrementPointName = Chr(Asc(FL) + 1) & LS
    Else
        IncrementPointName = "A" & Int(Val(LS) + 1)
    End If
    Exit Function
End If

If Len(S) > 1 Then
    For Z = 2 To Len(S)
        LS = Right(S, Len(S) - Z + 1)
        If IsNumeric(LS) Then
            IncrementPointName = Left(S, Z - 1) & Int(Val(LS) + 1)
            Exit Function
        End If
    Next
End If

IncrementPointName = S & "1"
End Function

Public Function DecrementPointName(ByVal S As String, Optional ByVal ReturnValidNameAnyway As Boolean = False) As String
Dim FL As String, LS As String, Z As Long

' This is done to further increment it to A again in IncrementPointName
If S = "A" Then S = "_": Exit Function

If S = "" And ReturnValidNameAnyway Then DecrementPointName = "A": Exit Function
FL = Left(S, 1)
LS = Right(S, Len(S) - 1)

If S = FL And IsAlpha(FL) Then
    If FL > "A" Then FL = Chr(Asc(FL) - 1) Else FL = "A"
    DecrementPointName = FL
    Exit Function
End If

If IsAlpha(FL) And IsNumeric(LS) Then
    If FL > "A" Then
        DecrementPointName = Chr(Asc(FL) - 1) & LS
    Else
        If Val(LS) > 1 Then
            DecrementPointName = "Z" & Int(Val(LS) - 1)
        Else
            DecrementPointName = "Z"
        End If
    End If
    Exit Function
End If

If Len(S) > 1 Then
    For Z = 2 To Len(S)
        LS = Right(S, Len(S) - Z + 1)
        If IsNumeric(LS) Then
            If LS > 1 Then
                DecrementPointName = Left(S, Z - 1) & Int(Val(LS) - 1)
                Exit Function
            Else
                DecrementPointName = FL
                Exit Function
            End If
        End If
    Next
End If

If ReturnValidNameAnyway Then DecrementPointName = S & "1"
End Function

Public Function IncrementFigureName(ByVal S As String) As String
Dim LS As String, Z As Long

If S = "" Then IncrementFigureName = "A": Exit Function

If Len(S) > 1 Then
    For Z = 2 To Len(S)
        LS = Right(S, Len(S) - Z + 1)
        If IsNumeric(LS) Then
            IncrementFigureName = Left(S, Z - 1) & Int(Val(LS) + 1)
            Exit Function
        End If
    Next
End If

IncrementFigureName = S & "1"
End Function

Public Function GenerateNewFigureName(ByVal FigureType As DrawState) As String
Dim tName As String

If FigureType < dsPoint Or FigureType > dsBisector Then FigureType = dsPoint
If LastAllocatedFigureName(FigureType) = "" Then LastAllocatedFigureName(FigureType) = GetString(ResFigureBase + 2 * FigureType) & "0"

tName = IncrementFigureName(LastAllocatedFigureName(FigureType))
Do While GetFigureByName(tName) >= 0
    tName = IncrementPointName(tName)
Loop
GenerateNewFigureName = tName
LastAllocatedFigureName(FigureType) = tName

'Dim Z As Long
'Dim TN As String, S As String
'S = GetString(ResFigureBase + FigureType * 2)
'
'Do
'    Z = Z + 1
'    TN = S & Z
'Loop Until FigureNames.FindItem(TN) = 0
'GenerateNewFigureName = TN
End Function

Public Function GenerateNewPointZOrder() As Long
Dim Z As Long
Dim tZOrders() As Long
Dim MaxZOrder As Long

If PointCount > 0 Then
    ReDim tZOrders(1 To PointCount)
    
    For Z = 1 To PointCount
        If BasePoint(Z).ZOrder > 0 And BasePoint(Z).ZOrder <= PointCount Then
            tZOrders(BasePoint(Z).ZOrder) = Z
        End If
    Next
    
    For Z = 1 To PointCount
        If tZOrders(Z) = 0 Then
            GenerateNewPointZOrder = Z
            Exit Function
        End If
    Next
    
    MaxZOrder = -2 ^ 30
    For Z = 1 To PointCount
        If BasePoint(Z).ZOrder > MaxZOrder Then MaxZOrder = BasePoint(Z).ZOrder
    Next
    MaxZOrder = MaxZOrder + 1
    GenerateNewPointZOrder = MaxZOrder
Else
    GenerateNewPointZOrder = 0
End If
'Dim Z As Long
'Dim MaxZOrder As Long
'
'If PointCount > 0 Then
'    MaxZOrder = -2 ^ 30
'    For Z = 1 To PointCount
'        If MaxZOrder < BasePoint(Z).ZOrder Then MaxZOrder = BasePoint(Z).ZOrder
'    Next
'    MaxZOrder = MaxZOrder + 1
'Else
'    MaxZOrder = 1
'End If
'GenerateNewPointZOrder = MaxZOrder
End Function

Public Function GenerateNewFigureZOrder() As Long
Dim Z As Long
Dim tZOrders() As Long
Dim MaxZOrder As Long

If FigureCount > 0 Then
    ReDim tZOrders(0 To FigureCount - 1)
    
    For Z = 0 To FigureCount - 1
        If Figures(Z).ZOrder >= 0 And Figures(Z).ZOrder < FigureCount Then
            tZOrders(Figures(Z).ZOrder) = Z + 1
        End If
    Next
    
    For Z = 0 To FigureCount - 1
        If tZOrders(Z) = 0 Then
            GenerateNewFigureZOrder = Z
            Exit Function
        End If
    Next
    
    MaxZOrder = -2 ^ 30
    For Z = 0 To FigureCount - 1
        If Figures(Z).ZOrder > MaxZOrder Then MaxZOrder = Figures(Z).ZOrder
    Next
    MaxZOrder = MaxZOrder + 1
    GenerateNewFigureZOrder = MaxZOrder
Else
    GenerateNewFigureZOrder = 0
End If
End Function

Public Function NumOfChildPoints(FigureType As DrawState) As Long
Select Case FigureType
    Case dsCircle_ArcCenterAndRadiusAndTwoPoints, dsIntersect
        NumOfChildPoints = 2
    Case dsMiddlePoint, dsSimmPoint, dsSimmPointByLine, dsInvert, dsPointOnFigure, dsAnPoint
        NumOfChildPoints = 1
    Case Else
        NumOfChildPoints = 0
End Select
End Function

Public Sub AddChildrenRecord(Figure1 As Figure, ChildFigure As Long)
Dim Z As Long
For Z = 0 To Figure1.NumberOfChildren - 1
    If Figure1.Children(Z) = ChildFigure Then Exit Sub
Next
Figure1.NumberOfChildren = Figure1.NumberOfChildren + 1
ReDim Preserve Figure1.Children(0 To Figure1.NumberOfChildren - 1)
Figure1.Children(Figure1.NumberOfChildren - 1) = ChildFigure
End Sub
'
'Public Sub UndeletePoint(pActionIndex As Long)
'On Local Error GoTo EH:
'With Activity(pActionIndex)
'    If .sPoint(1).ZOrder > 0 Then
'        For Z = 1 To PointCount
'            If BasePoint(Z).ZOrder >= .sPoint(1).ZOrder And BasePoint(Z).Tag = 0 Then BasePoint(Z).ZOrder = BasePoint(Z).ZOrder + 1
'        Next
'    End If
'
'    If FigureCount > 0 Then
'        For Z = 0 To FigureCount - 1
'            For Q = 0 To Figures(Z).NumberOfPoints - 1
'                'If .sPoint(1).Type <> dsPoint Then
'                'If .sPoint(1).ParentFigure = Z Then
'                'If IsChildPointPos(Figures(Z), Q) And Figures(Z).AuxInfo(6) <> 0 Then GoTo NextQ
'                'End If
'                'End If
'                'If PointCount >= Figures(Z).Points(Q) Then
'                If Figures(Z).Points(Q) >= .pPoint And Figures(Z).Tag = 0 Then Figures(Z).Points(Q) = Figures(Z).Points(Q) + 1
'                'End If
''NextQ:
'            Next Q
'        Next Z
'    End If
'
'    PointCount = PointCount + 1
'    ReDimPreserveBasePoint 1, PointCount
'    If .pPoint < PointCount Then
'        For Z = PointCount To .pPoint + 1 Step -1
'            BasePoint(Z) = BasePoint(Z - 1)
'
''            For Q = 0 To FigureCount - 1
''                For I = 0 To Figures(Q).NumberOfPoints - 1
''                    If Figures(Q).Points(I) = Z - 1 Then Figures(Q).Points(I) = Z
''                Next
''            Next
'
'            If BasePoint(Z).Locus > 0 Then
'                Locuses(BasePoint(Z).Locus).ParentPoint = Locuses(BasePoint(Z).Locus).ParentPoint + 1
'            End If
'        Next Z
'    End If
'    BasePoint(.pPoint) = .sPoint(1)
'
'    BasePoint(.pPoint).Tag = 1
'
'    'For Z = 0 To Figures(BasePoint(.pPoint).ParentFigure).NumberOfPoints - 1
'    '    If Figures(BasePoint(.pPoint).ParentFigure).Points(Z) = -.pPoint Then Figures(BasePoint(.pPoint).ParentFigure).Points(Z) = .pPoint
'    'Next
'
'    If .pLocus <> 0 Then
'        For Z = 1 To PointCount
'            If Z <> .pPoint Then If BasePoint(Z).Locus >= .pLocus Then BasePoint(Z).Locus = BasePoint(Z).Locus + 1
'        Next
'        LocusCount = LocusCount + 1
'        ReDim Preserve Locuses(1 To LocusCount)
'        If .pLocus < LocusCount Then
'            For Z = LocusCount To .pLocus + 1 Step -1
'                Locuses(Z) = Locuses(Z - 1)
'            Next
'        End If
'        Locuses(.pLocus) = .sLocus(1)
'    End If
'End With
'
'EH:
'End Sub
'
'Public Sub UndeleteFigure(pActionIndex As Long)
'Dim FigNum As Long
'
'On Local Error GoTo EH:
'
'With Activity(pActionIndex)
'    If PointCount > 0 Then
'        For Z = 1 To PointCount
'            If Not IsChildPoint(.sFigure(1), Z) Then
'                If BasePoint(Z).Type <> dsPoint Then
'                    If BasePoint(Z).ParentFigure >= .pFigure And BasePoint(Z).Tag = 0 Then BasePoint(Z).ParentFigure = BasePoint(Z).ParentFigure + 1
'                End If
'            End If
'        Next
'    End If
'
'    For Z = 0 To FigureCount - 1
'        N = GetProperParentNumber(Figures(Z).FigureType)
'        If N > 0 Then
'            For Q = 0 To N - 1
'                If Figures(Z).Parents(Q) >= .pFigure And Figures(Z).Tag = 0 Then Figures(Z).Parents(Q) = Figures(Z).Parents(Q) + 1
'                'Figures(Figures(Z).Parents(Q)).NumberOfChildren = Figures(Figures(Z).Parents(Q)).NumberOfChildren + 1
'                'ReDim Preserve Figures(Figures(Z).Parents(Q)).Children(0 To Figures(Figures(Z).Parents(Q)).NumberOfChildren - 1)
'                'Figures(Figures(Z).Parents(Q)).Children(Figures(Figures(Z).Parents(Q)).NumberOfChildren - 1) = .pFigure
'            Next
'        End If
'        If Figures(Z).NumberOfChildren > 0 Then
'            For Q = 0 To Figures(Z).NumberOfChildren - 1
'                If Figures(Z).Children(Q) >= .pFigure And Figures(Z).Tag = 0 Then Figures(Z).Children(Q) = Figures(Z).Children(Q) + 1
'            Next
'        End If
'        If Figures(Z).ZOrder >= .sFigure(1).ZOrder And Figures(Z).Tag = 0 Then Figures(Z).ZOrder = Figures(Z).ZOrder + 1
'    Next
'
'    FigureCount = FigureCount + 1
'    ReDim Preserve Figures(0 To FigureCount - 1)
'    If .pFigure < FigureCount - 1 Then
'        For Z = FigureCount - 1 To .pFigure + 1 Step -1
'            Figures(Z) = Figures(Z - 1)
'        Next Z
'    End If
'    Figures(.pFigure) = .sFigure(1)
'End With
'
'FigNum = Activity(pActionIndex).pFigure
'
'N = GetProperParentNumber(FigNum)
'If N > 0 Then
'    For Z = 0 To N - 1
'        AddChildrenRecord Figures(Figures(FigNum).Parents(Z)), FigNum
''        With Figures(Figures(FigNum).Parents(Z))
''            .NumberOfChildren = .NumberOfChildren + 1
''            ReDim Preserve .Children(0 To .NumberOfChildren - 1)
''            .Children(.NumberOfChildren - 1) = FigNum
''        End With
'    Next
'End If
'
'Figures(FigNum).AuxInfo(6) = 1 'NumOfChildPoints(Figures(FigNum).FigureType)
'Figures(FigNum).AuxInfo(7) = 1 'Figures(FigNum).NumberOfChildren
'Figures(FigNum).Tag = 1
'
''k = 0
''If Figures(FigNum).NumberOfPoints > 0 Then
''    For Z = 0 To Figures(FigNum).NumberOfPoints - 1
''        If IsChildPointPos(Figures(FigNum), Z) Then
''            k = k + 1
''            Figures(FigNum).Points(Z) = -Figures(FigNum).Points(Z)
''        End If
''    Next
''End If
'
'EH:
'End Sub

Public Function FigureMemoryConsumption(pFigure As Figure) As Long
Dim MemSum As Long
MemSum = MemSum + Len(pFigure)
With pFigure
    MemSum = MemSum + 4 * .NumberOfPoints
    MemSum = MemSum + 4 * .NumberOfChildren
    MemSum = MemSum + 4 * GetProperParentNumber(.FigureType)
End With
FigureMemoryConsumption = MemSum
End Function

Public Function DynamicTreeMemoryConsumption(pTree As Tree) As Long
DynamicTreeMemoryConsumption = Len(pTree) + Len(pTree.Branches(1)) * pTree.BranchCount
End Function

Public Sub RenamePoint(ByVal Point1 As Long, ByVal NewName As String)
If IsPoint(GetPointByName(NewName)) Then
    i_AskAboutPointRename Point1, NewName
Else
    RecordGenericAction ResUndoRenamePoint
    RenamePointActions Point1, NewName
End If

'RecordAction pRenameAction  '?????
'
End Sub

Public Sub RenamePointActions(ByVal Point1 As Long, ByVal NewName As String)
Dim tName As String

If IsPoint(GetPointByName(NewName)) Then Exit Sub

BasePoint(Point1).Name = NewName
ReRestoreExpressionsFromTrees
End Sub

'========================================================================
' This procedure restores all expression strings from all trees in the DS (data structures)
' Called from RenamePoint
'========================================================================

Public Sub ReRestoreExpressionsFromTrees()
Dim BigTextSum As String
Dim Z As Long, Q As Long

For Z = 1 To LabelCount
    With TextLabels(Z)
        If .Dynamic Then
            BigTextSum = ""
            For Q = 1 To .CCCP
                If .CompiledCaptionParts(Q).Type = DynamicString Then
                   .CompiledCaptionParts(Q).StaticText = RestoreExpressionFromTree(.CompiledCaptionParts(Q).DynamicTree)
                    BigTextSum = BigTextSum & "[" & .CompiledCaptionParts(Q).StaticText & "]"
                Else
                    BigTextSum = BigTextSum & .CompiledCaptionParts(Q).StaticText
                End If
            Next
            .Caption = BigTextSum
        End If
    End With
Next
UpdateLabels

For Z = 0 To FigureCount - 1
    If Figures(Z).FigureType = dsAnPoint Then
        Figures(Z).XS = RestoreExpressionFromTree(Figures(Z).XTree)
        Figures(Z).YS = RestoreExpressionFromTree(Figures(Z).YTree)
    End If
Next
End Sub

'
'Public Sub RenamePoint(ByVal Point1 As Long, ByVal NewName As String)
'Dim tName As String, BigTextSum As String, Z As Long, Q As Long
'If IsPoint(GetPointByName(NewName)) Then Exit Sub
'
'PointNames.ReplaceItem PointNames.FindItem(BasePoint(Point1).Name), NewName
'BasePoint(Point1).Name = NewName
'
'For Z = 1 To LabelCount
'    With TextLabels(Z)
'        If .Dynamic Then
'            BigTextSum = ""
'            For Q = 1 To .CCCP
'                If .CompiledCaptionParts(Q).Type = DynamicString Then
'                   .CompiledCaptionParts(Q).StaticText = RestoreExpressionFromTree(.CompiledCaptionParts(Q).DynamicTree)
'                    BigTextSum = BigTextSum & "[" & .CompiledCaptionParts(Q).StaticText & "]"
'                Else
'                    BigTextSum = BigTextSum & .CompiledCaptionParts(Q).StaticText
'                End If
'            Next
'            .Caption = BigTextSum
'        End If
'    End With
'Next
'UpdateLabels
'
'For Z = 1 To WECount
'    With WatchExpressions(Z)
'        If TreeDependsOnPoint(Point1, .WatchTree) Then
'            If .Name = .Expression Then .Name = ""
'            .Expression = RestoreExpressionFromTree(.WatchTree)
'            If .Name = "" Then .Name = .Expression
'        End If
'    End With
'Next
''FormMain.ValueTable1.UpdateExpressions
'
'For Z = 0 To FigureCount - 1
'    If Figures(Z).FigureType = dsAnPoint Then
'        Figures(Z).XS = RestoreExpressionFromTree(Figures(Z).XTree)
'        Figures(Z).YS = RestoreExpressionFromTree(Figures(Z).YTree)
'    End If
'Next
'End Sub

Public Function GetPolygonArea(Points() As Long) As Double
Dim Pts() As OnePoint, Z As Long
If UBound(Points) < 2 Then Exit Function
ReDim Pts(1 To UBound(Points))
For Z = 1 To UBound(Pts)
    Pts(Z).X = BasePoint(Points(Z)).X
    Pts(Z).Y = BasePoint(Points(Z)).Y
Next
GetPolygonArea = Area(Pts)
End Function

Public Function WereActiveAxesAdded() As Boolean
Dim WasX As Boolean, WasY As Boolean, Z As Long
For Z = 0 To FigureCount - 1
    If Figures(Z).FigureType = dsAnLineGeneral Then
        If Figures(Z).AuxInfo(1) = 0 And Figures(Z).AuxInfo(2) = 1 And Figures(Z).AuxInfo(3) = 0 Then WasX = True
    End If
    If Figures(Z).FigureType = dsAnLineGeneral Then
        If Figures(Z).AuxInfo(1) = 1 And Figures(Z).AuxInfo(2) = 0 And Figures(Z).AuxInfo(3) = 0 Then WasY = True
    End If
Next
nActiveAxesAdded = WasX And WasY
If Not nActiveAxesAdded Then nActiveX = -1: nActiveY = -1
FormMain.mnuActiveAxes.Enabled = Not nActiveAxesAdded
WereActiveAxesAdded = nActiveAxesAdded
End Function

Public Sub HidePoint(ByVal index As Long)
Dim pAction As Action

If index < 1 Or index > PointCount Then Exit Sub

pAction.Type = actHidePoint
pAction.pPoint = ActivePoint
RecordAction pAction

If BasePoint(index).Locus > 0 Then Locuses(BasePoint(index).Locus).ShouldBreak = True

BasePoint(index).Hide = True
BasePoint(index).Enabled = False

PaperCls
ShowAll
End Sub

Public Sub ToggleAxes(Optional ByVal TurnOn As Boolean = True, Optional ByVal ShouldRedraw As Boolean = True)
' Turns axes TurnOn ? on : off.
'=================================================
If nShowAxes = TurnOn Then Exit Sub

nShowAxes = TurnOn
FormMain.mnuShowAxes.Checked = nShowAxes
'setShowAxes = nShowAxes
'SaveSetting AppName, "Interface", "ShowAxes", Format(-CInt(setShowAxes))

If Not nShowAxes Then
    If IsFigure(nActiveX) And IsFigure(nActiveY) Then
        Figures(nActiveX).Hide = True
        Figures(nActiveY).Hide = True
    End If
Else
    If IsFigure(nActiveX) And IsFigure(nActiveY) Then
        Figures(nActiveX).Hide = False
        Figures(nActiveY).Hide = False
    Else
        AddActiveAxes
    End If
End If

If ShouldRedraw Then
    PaperCls
    ShowAll
End If
End Sub

Public Sub SnapAllPointNamesToPoints()
Dim Z As Long

For Z = 1 To PointCount
    SnapPointLabel Z
Next
End Sub

Public Sub SnapPointLabel(ByVal PointIndex As Long)
Dim tX As Double, tY As Double, TW As Double

If PointIndex < 1 Or PointIndex > PointCount Then Exit Sub

With BasePoint(PointIndex)
    tX = .LabelOffsetX
    tY = .LabelOffsetY
    TW = Maximum(.LabelWidth, .LabelHeight) + .PhysicalWidth + setCursorSensitivity
    If Sqr(tX ^ 2 + tY ^ 2) > TW Then
        NormalizeVector tX, tY
        tX = tX * TW
        tY = tY * TW
        tY = tY + .PhysicalWidth
    End If
    .LabelOffsetX = tX
    .LabelOffsetY = tY
    
End With
End Sub
