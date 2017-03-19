Attribute VB_Name = "modSG"
Option Explicit

Public Sub AddStaticGraphic(ByVal sgType As StaticGraphicType, Points() As Long, Optional ByVal ShouldRecord As Boolean = True, Optional ByVal ShouldRefresh As Boolean = True, Optional ByVal ShouldAddSegments As Boolean = True)
Dim Z As Long

If ShouldRecord Then RecordGenericAction Switch(sgType = sgPolygon, ResUndoCreatePolygon, sgType = sgBezier, ResUndoCreateBezier, sgType = sgVector, ResUndoCreateVector)

StaticGraphicCount = StaticGraphicCount + 1
ReDim Preserve StaticGraphics(1 To StaticGraphicCount)
With StaticGraphics(StaticGraphicCount)
    .Type = sgType
    If .Type = sgPolygon Then .DrawMode = 9 Else .DrawMode = 13
    .DrawStyle = vbSolid
    .DrawWidth = 1
    .ForeColor = setdefcolFigure
    .FillColor = setdefcolFigureFill
    .FillStyle = -1
    .Visible = True
    .NumberOfPoints = UBound(Points)
    If Points(.NumberOfPoints) = Points(1) Then
        .NumberOfPoints = .NumberOfPoints - 1
        'ReDim Preserve Points(1 To .NumberOfPoints)
    End If
    
    ReDim .Points(1 To .NumberOfPoints)
    For Z = 1 To .NumberOfPoints
        .Points(Z) = Points(Z)
    Next
    
    .Description = GetObjectDescription(gotSG, StaticGraphicCount)
    .InDemo = True
    ReDim .ObjectPixels(0 To .NumberOfPoints)
    ReDim .ObjectPoints(1 To .NumberOfPoints)
End With

If sgType = sgPolygon And ShouldAddSegments Then
    With StaticGraphics(StaticGraphicCount)
        For Z = 1 To .NumberOfPoints
            If FindFigureWithPoints(dsSegment, Val(Points(Z)), Val(Points(IIf(Z < .NumberOfPoints, Z + 1, 1)))) = -1 Then
                AddSegment Points(Z), Points(IIf(Z < .NumberOfPoints, Z + 1, 1)), , False
            End If
        Next
    End With
End If

If sgType = sgVector And ShouldAddSegments Then
    If FindFigureWithPoints(dsSegment, Points(1), Points(2)) = -1 Then
        AddSegment Points(1), Points(2), , False
    End If
    StaticGraphics(StaticGraphicCount).FillColor = Figures(FindFigureWithPoints(dsSegment, Points(1), Points(2))).ForeColor
    StaticGraphics(StaticGraphicCount).ForeColor = StaticGraphics(StaticGraphicCount).FillColor
End If

RecalcStaticGraphic StaticGraphicCount
If ShouldRefresh Then PaperCls: ShowAll
End Sub

Public Sub DeleteStaticGraphic(ByVal Index As Long, Optional ByVal ShouldRecord As Boolean = True)
Dim pAction As Action, Z As Long

If ShouldRecord Then
    pAction.Type = actRemoveSG
    MakeStructureSnapshot pAction
    ReDim pAction.AuxInfo(1 To 1)
    pAction.AuxInfo(1) = Index
    RecordAction pAction
End If

DeleteFromDependentButtons Index, gotSG, False

If StaticGraphics(Index).Type = sgVector Then
    BasePoint(StaticGraphics(Index).Points(2)).Hide = False
End If

If Index < StaticGraphicCount Then
    For Z = Index To StaticGraphicCount - 1
        StaticGraphics(Z) = StaticGraphics(Z + 1)
        OffsetButtonObjectDependencies Z + 1, Z, gotSG
    Next
End If
StaticGraphicCount = StaticGraphicCount - 1
If StaticGraphicCount > 0 Then ReDim Preserve StaticGraphics(1 To StaticGraphicCount) Else ReDim StaticGraphics(1 To 1)
End Sub

Public Function StaticGraphicDependsOnPoint(ByVal SG1 As Long, ByVal Point1 As Long) As Boolean
Dim Z As Long
For Z = 1 To StaticGraphics(SG1).NumberOfPoints
    If StaticGraphics(SG1).Points(Z) = Point1 Then StaticGraphicDependsOnPoint = True: Exit Function
Next
StaticGraphicDependsOnPoint = False
End Function

Public Function IsSG(ByVal Index As Long) As Boolean
If Index >= 1 And Index <= StaticGraphicCount Then IsSG = True
End Function

Public Function RecalcStaticGraphic(ByVal Index As Long)
Dim X As Double, Y As Double, Z As Long
If Index = 0 Then Exit Function
With StaticGraphics(Index)
    .Visible = True
    For Z = 1 To .NumberOfPoints
        X = BasePoint(.Points(Z)).X
        Y = BasePoint(.Points(Z)).Y
        .ObjectPoints(Z).X = X
        .ObjectPoints(Z).Y = Y
        ToPhysical X, Y
        .ObjectPixels(Z).X = X
        .ObjectPixels(Z).Y = Y
        If Not BasePoint(.Points(Z)).Visible Then .Visible = False: Exit Function
    Next
    If .Type = sgVector Then
        RecalcVectorArrow Index
    End If
End With
End Function

Public Function RecalcStaticGraphics()
Dim Z As Long
For Z = 1 To StaticGraphicCount
    RecalcStaticGraphic Z
Next
End Function

Public Function GetSGFromPoint(ByVal X As Double, ByVal Y As Double) As Long
Dim OX As Double, OY As Double, hRgn As Long, Z As Long, Segm As Long
OX = X
OY = Y
ToPhysical OX, OY

For Z = 1 To StaticGraphicCount
    With StaticGraphics(Z)
        If .Visible Then
            Select Case .Type
                Case sgPolygon
                    hRgn = CreatePolygonRgn(.ObjectPixels(1), .NumberOfPoints, ALTERNATE)
                    If PtInRegion(hRgn, OX, OY) <> 0 Then GetSGFromPoint = Z
                    DeleteObject hRgn
                Case sgBezier
                    If PtInBezier(Z, X, Y) Then GetSGFromPoint = Z
                Case sgVector
                    Segm = GetFigureByPoint(X, Y)
                    If Segm <> -1 Then
                        If Figures(Segm).FigureType = dsSegment Then
                            If FigureHasPoint(Segm, .Points(1)) And FigureHasPoint(Segm, .Points(2)) Then
                                GetSGFromPoint = Z
                            End If
                        End If
                    End If
            End Select
        End If
    End With
Next
End Function

Public Function GetSGsFromPoint(ByVal X As Double, ByVal Y As Double) As Long()
Dim OX As Double, OY As Double, hRgn As Long, Z As Long, Segm As Long
Dim Arr() As Long, Count As Long

OX = X
OY = Y
ToPhysical OX, OY

ReDim Arr(1 To 1)
Count = 0

For Z = 1 To StaticGraphicCount
    With StaticGraphics(Z)
        If .Visible Then
            Select Case .Type
                Case sgPolygon
                    hRgn = CreatePolygonRgn(.ObjectPixels(1), .NumberOfPoints, ALTERNATE)
                    If PtInRegion(hRgn, OX, OY) <> 0 Then
                        Count = Count + 1
                        ReDim Preserve Arr(1 To Count)
                        Arr(Count) = Z
                    End If
                    DeleteObject hRgn
                Case sgBezier
                    If PtInBezier(Z, X, Y) Then
                        Count = Count + 1
                        ReDim Preserve Arr(1 To Count)
                        Arr(Count) = Z
                    End If
                Case sgVector
                    Segm = GetFigureByPoint(X, Y)
                    If Segm <> -1 Then
                        If Figures(Segm).FigureType = dsSegment Then
                            If FigureHasPoint(Segm, .Points(1)) And FigureHasPoint(Segm, .Points(2)) Then
                                Count = Count + 1
                                ReDim Preserve Arr(1 To Count)
                                Arr(Count) = Z
                            End If
                        End If
                    End If
            End Select
        End If
    End With
Next
End Function

Public Function RecalcVectorArrow(ByVal SGNum As Long) As Boolean
Const K = 1 / 3
Dim pPoints(1 To 3) As POINTAPI
Dim pPointsExact(0 To 3) As OnePoint
Dim dDist As Double
Dim ArrLength As Double, hRgn As Long
Dim X1 As Double, Y1 As Double, X2 As Double, Y2 As Double

X1 = BasePoint(StaticGraphics(SGNum).Points(1)).X
Y1 = BasePoint(StaticGraphics(SGNum).Points(1)).Y
X2 = BasePoint(StaticGraphics(SGNum).Points(2)).X
Y2 = BasePoint(StaticGraphics(SGNum).Points(2)).Y

pPointsExact(1).X = X2
pPointsExact(1).Y = Y2
ArrLength = ArrowLength
ToLogicalLength ArrLength
dDist = Distance(X1, Y1, X2, Y2)
If dDist = 0 Then Exit Function
dDist = dDist / ArrLength
pPointsExact(0).X = X2 + (X1 - X2) / dDist
pPointsExact(0).Y = Y2 + (Y1 - Y2) / dDist
pPointsExact(2).X = pPointsExact(0).X + (pPointsExact(1).Y - pPointsExact(0).Y) * K
pPointsExact(2).Y = pPointsExact(0).Y + (pPointsExact(0).X - pPointsExact(1).X) * K
pPointsExact(3).X = pPointsExact(0).X + (pPointsExact(0).Y - pPointsExact(1).Y) * K
pPointsExact(3).Y = pPointsExact(0).Y + (pPointsExact(1).X - pPointsExact(0).X) * K

ToPhysical pPointsExact(1).X, pPointsExact(1).Y
ToPhysical pPointsExact(2).X, pPointsExact(2).Y
ToPhysical pPointsExact(3).X, pPointsExact(3).Y

pPoints(1).X = pPointsExact(2).X
pPoints(1).Y = pPointsExact(2).Y
pPoints(2).X = pPointsExact(1).X
pPoints(2).Y = pPointsExact(1).Y
pPoints(3).X = pPointsExact(3).X
pPoints(3).Y = pPointsExact(3).Y

StaticGraphics(SGNum).ObjectPixels = pPoints
End Function

Public Function PtInBezier(ByVal hSG As Long, ByVal X As Double, ByVal Y As Double) As Boolean
Dim ValidDistance As Double

With StaticGraphics(hSG)
    ValidDistance = .DrawWidth
    ToLogicalLength ValidDistance
    ValidDistance = ValidDistance + Sensitivity * 2
    If DistanceToBezier(X, Y, .ObjectPoints) <= ValidDistance Then PtInBezier = True Else PtInBezier = False
End With
End Function

Public Sub OffsetSGPointDependencies(ByVal OldValue As Long, ByVal NewValue As Long)
Dim Z As Long, Q As Long

For Z = 1 To StaticGraphicCount
    For Q = 1 To StaticGraphics(Z).NumberOfPoints
        If StaticGraphics(Z).Points(Q) = OldValue Then StaticGraphics(Z).Points(Q) = NewValue
    Next
Next
End Sub

