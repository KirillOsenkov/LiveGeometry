Attribute VB_Name = "modDataStructure"
Option Explicit

'===================================================
'===================================================
'                                       INDEX OFFSETS
'===================================================
'===================================================

Public Sub OffsetLocusParentFigureDependencies(ByVal OldValue As Long, ByVal NewValue As Long)
Dim Z As Long
For Z = 1 To LocusCount
    If Locuses(Z).Dynamic Then If Locuses(Z).ParentFigure = OldValue Then Locuses(Z).ParentFigure = NewValue
Next Z
End Sub

Public Sub OffsetObjectPixels(ObjPixels() As POINTAPI, ByVal DX As Integer, ByVal DY As Integer)
Dim Z As Long
For Z = LBound(ObjPixels) To UBound(ObjPixels)
    ObjPixels(Z).X = ObjPixels(Z).X + DX
    ObjPixels(Z).Y = ObjPixels(Z).Y + DY
Next
End Sub

'====================================================
'                               ARRAY MANAGEMENT
'====================================================

Public Function DeleteFigure(ByVal Index As Integer, Optional ByVal NeedToShowAll As Boolean = True, Optional ByVal ShouldRecord As Boolean = True, Optional ByVal DontDeleteChildPoints As Boolean = False) As Long
Dim N As Long, OldZOrder As Long, Q As Long, Z As Long
Dim pAction As Action

If ShouldRecord Then
    RecordGenericAction ResUndoDeleteFigure
    'pAction.Type = actRemoveFigure
    'MakeStructureSnapshot pAction
    'RecordAction pAction
End If

'FigureNames.Remove FigureNames.FindItem(Figures(Index).Name)

DeleteFromDependentButtons Index, gotFigure, False

OldZOrder = Figures(Index).ZOrder
HideFigure Index, False

'Do While Figures(Index).NumberOfChildren > 0
'    DeleteFigure Figures(Index).Children(Figures(Index).NumberOfChildren - 1), False
'Loop

'?????

If Not DontDeleteChildPoints Then
    Select Case Figures(Index).FigureType
        Case dsCircle_ArcCenterAndRadiusAndTwoPoints
            DeletePoint Figures(Index).Points(5), False, False ', ShouldRecord
            DeletePoint Figures(Index).Points(6), False, False 'ShouldRecord
            'pAction.sFigure(1).Points(5) = -pAction.sFigure(1).Points(5) '- 1
            'pAction.sFigure(1).Points(6) = -pAction.sFigure(1).Points(6) '- 1
            'pAction.Group = pAction.Group + 2
        Case dsIntersect
            DeletePoint Figures(Index).Points(0), False, False ' ShouldRecord
            DeletePoint Figures(Index).Points(1), False, False 'ShouldRecord
            'pAction.sFigure(1).Points(0) = -pAction.sFigure(1).Points(0) '- 1
            'pAction.sFigure(1).Points(1) = -pAction.sFigure(1).Points(1) '- 1
            'pAction.Group = pAction.Group + 2
        Case dsMiddlePoint, dsSimmPoint, dsSimmPointByLine, dsInvert, dsAnPoint
            DeletePoint Figures(Index).Points(0), False, False 'ShouldRecord
            'pAction.sFigure(1).Points(0) = -pAction.sFigure(1).Points(0) '- 1
            'pAction.Group = pAction.Group + 1
        Case dsPointOnFigure
            If Figures(Index).AuxInfo(5) = 1 Then
                ShowPoint Paper.hDC, Figures(Index).Points(0), True
                BasePoint(Figures(Index).Points(0)).Type = dsPoint
                BasePoint(Figures(Index).Points(0)).ParentFigure = 0
                BasePoint(Figures(Index).Points(0)).ForeColor = setdefcolPoint
                BasePoint(Figures(Index).Points(0)).Shape = defPointShape
            Else
                DeletePoint Figures(Index).Points(0), False, False
            End If
    End Select
End If

If Figures(Index).FigureType = dsDynamicLocus Then
    Locuses(BasePoint(Figures(Index).Points(0)).Locus).Dynamic = False
    Locuses(BasePoint(Figures(Index).Points(0)).Locus).ParentFigure = -1
    If DontDeleteChildPoints = False Then
        DeleteLocus Figures(Index).Points(0), False
    End If
End If

Z = 0
Do While Z <= FigureCount - 1
    N = GetProperParentNumber(Figures(Z).FigureType)
    If N > 0 Then
        For Q = 0 To N - 1
            If Figures(Z).Parents(Q) = Index Then
                DeleteFigure Z, False, False 'ShouldRecord
                pAction.Group = pAction.Group + 1
                GoTo NextZ
            End If
        Next Q
    End If
Z = Z + 1
NextZ:
Loop

N = GetProperParentNumber(Figures(Index).FigureType)
If N > 0 Then
    For Z = 0 To N - 1
        RemoveFromChildren Figures(Figures(Index).Parents(Z)), Index
        'With Figures(Figures(Index).Parents(Z))
'            For Q = 0 To .NumberOfChildren - 1
'                If .Children(Q) = Index Then
'                    If Q < .NumberOfChildren - 1 Then
'                        For I = Q To .NumberOfChildren - 2
'                            .Children(I) = .Children(I + 1)
'                        Next I
'                    End If
'                    Exit For
'                End If
'            Next Q
'            .NumberOfChildren = .NumberOfChildren - 1
'            If .NumberOfChildren > 0 Then ReDim Preserve .Children(0 To .NumberOfChildren - 1) Else ReDim .Children(0 To 0)
'        End With
    Next Z
End If

'_____________________________________________________

FigureCount = FigureCount - 1
If FigureCount > 0 Then
    For Z = Index To FigureCount - 1
        Figures(Z) = Figures(Z + 1)
        For Q = 0 To FigureCount - 1
            OffsetAnPointFigureDependencies Q, Z + 1, Z
        Next
        OffsetButtonObjectDependencies Z + 1, Z, gotFigure
        OffsetLocusParentFigureDependencies Z + 1, Z
    Next Z
    ReDim Preserve Figures(0 To FigureCount - 1)
End If

For Z = 0 To FigureCount - 1
    N = GetProperParentNumber(Figures(Z).FigureType)
    If N > 0 Then
        For Q = 0 To N - 1
            If Figures(Z).Parents(Q) > Index Then Figures(Z).Parents(Q) = Figures(Z).Parents(Q) - 1
        Next
    End If
    If Figures(Z).NumberOfChildren > 0 Then
        For Q = 0 To Figures(Z).NumberOfChildren - 1
            If Figures(Z).Children(Q) > Index Then Figures(Z).Children(Q) = Figures(Z).Children(Q) - 1
        Next
    End If
    If Figures(Z).ZOrder > OldZOrder Then Figures(Z).ZOrder = Figures(Z).ZOrder - 1
    
'    Q = 0
'    Do While Q <= Figures(Z).NumberOfChildren - 1
'        If Figures(Z).Children(Q) = Index Then
'            If Q < Figures(Z).NumberOfChildren - 1 Then
'                For T = Q To Figures(Z).NumberOfChildren - 2
'                    Figures(Z).Children(T) = Figures(Z).Children(T + 1)
'                Next
'            End If
'            Figures(Z).NumberOfChildren = Figures(Z).NumberOfChildren - 1
'            If Figures(Z).NumberOfChildren > 0 Then ReDim Preserve Figures(Z).Children(0 To Figures(Z).NumberOfChildren - 1) '?????
'        End If
'        Q = Q + 1
'    Loop
    
'    If Figures(Z).NumberOfChildren > 0 Then
'        For Q = 0 To Figures(Z).NumberOfChildren - 1
'            If Figures(Z).Children(Q) = Index Then
'                If Q < Figures(Z).NumberOfChildren - 1 Then
'                    For T = Q To Figures(Z).NumberOfChildren - 2
'                        Figures(Z).Children(T) = Figures(Z).Children(T + 1)
'                    Next
'                End If
'                Figures(Z).NumberOfChildren = Figures(Z).NumberOfChildren - 1
'                If Figures(Z).NumberOfChildren > 0 Then ReDim Preserve Figures(Z).Children(0 To Figures(Z).NumberOfChildren - 1) '?????
'            End If
'        Next
'    End If
Next

For Z = 1 To PointCount
    If BasePoint(Z).Type <> dsPoint Then
        If BasePoint(Z).ParentFigure > Index Then BasePoint(Z).ParentFigure = BasePoint(Z).ParentFigure - 1
    End If
Next

RebuildDynamicLocusDependencies

WereActiveAxesAdded

If NeedToShowAll Then PaperCls: ShowAll
End Function

Public Sub DeletePoint(ByVal Index As Long, Optional ByVal NeedToShowAll As Boolean = True, Optional ByVal ShouldRecord As Boolean = True)
Dim pAction As Action, OldZOrder As Long, Z As Long, Q As Long
Dim S As String

If ShouldRecord Then
    RecordGenericAction ResUndoDeletePoint
    'pAction.Type = actRemovePoint
    'MakeStructureSnapshot pAction
    'RecordAction pAction
End If

S = DecrementPointName(BasePoint(Index).Name)
If S <> "" Then LastAllocatedPointName = S

'PointNames.Remove PointNames.FindItem(BasePoint(Index).Name)

DeleteFromDependentButtons Index, gotPoint, False

OldZOrder = BasePoint(Index).ZOrder

ShowPoint Paper.hDC, Index, True

'Z = 1
'Do While Z <= WECount
'    If TreeDependsOnPoint(Index, WatchExpressions(Z).WatchTree) Then
'        RemoveWatchExpression Z, False
'    Else
'        Z = Z + 1
'    End If
'Loop

Z = 1
Do While Z <= StaticGraphicCount
    If StaticGraphicDependsOnPoint(Z, Index) Then DeleteStaticGraphic Z, False Else Z = Z + 1
Loop

Z = 0
Do While Z < FigureCount
    If IsParentPoint(Figures(Z), Index) Then
        DeleteFigure Z, False, False ', ShouldRecord: pAction.Group = pAction.Group + 1
    Else
        Z = Z + 1
    End If
Loop

If BasePoint(Index).Locus <> 0 Then DeleteLocus Index, False

If PointCount > 0 Then
    If Index < PointCount Then
        For Z = Index To PointCount - 1
            BasePoint(Z) = BasePoint(Z + 1)
            'OffsetWEPointDependencies Z + 1, Z
            OffsetSGPointDependencies Z + 1, Z
            OffsetAnPointDependencies Z + 1, Z
            OffsetButtonObjectDependencies Z + 1, Z, gotPoint
            If BasePoint(Z).Locus > 0 Then
                Locuses(BasePoint(Z).Locus).ParentPoint = Locuses(BasePoint(Z).Locus).ParentPoint - 1
            End If
        Next Z
    End If
    If PointCount > 1 Then ReDim Preserve BasePoint(1 To PointCount - 1)
End If
PointCount = PointCount - 1

If FigureCount > 0 Then
    For Z = 0 To FigureCount - 1
        For Q = 0 To Figures(Z).NumberOfPoints - 1
            If Figures(Z).Points(Q) > Index Then Figures(Z).Points(Q) = Figures(Z).Points(Q) - 1
        Next
    Next
End If

If OldZOrder > 0 Then
    For Z = 1 To PointCount
        If BasePoint(Z).ZOrder > OldZOrder Then BasePoint(Z).ZOrder = BasePoint(Z).ZOrder - 1
    Next Z
End If

'FormMain.ValueTable1.UpdateExpressions
If NeedToShowAll Then PaperCls: ShowAll
End Sub

Public Function VisualFigureCount() As Long
Dim Z As Long, C As Long

C = 0
For Z = 0 To FigureCount - 1
    If IsVisual(Z) Then C = C + 1
Next

VisualFigureCount = C
End Function

Public Function IndependentFigureCount() As Long
Dim Z As Long, C As Long

C = 0
For Z = 0 To FigureCount - 1
    If IsIndependentFigure(Z) Then C = C + 1
Next

IndependentFigureCount = C
End Function

'====================================================
'                                       FIGURE TYPE
'====================================================

Public Function IsVisual(ByVal Figure1 As Long) As Boolean
If Figure1 < FigureCount And Figure1 >= 0 Then _
    If (Figures(Figure1).FigureType >= dsSegment _
    And Figures(Figure1).FigureType <= dsLine_PointAndPerpendicularLine) _
    Or (Figures(Figure1).FigureType >= dsAnLineGeneral _
    And Figures(Figure1).FigureType <= dsAnLineNormalPoint) _
    Or (Figures(Figure1).FigureType = dsCircle_CenterAndCircumPoint _
    Or Figures(Figure1).FigureType = dsCircle_CenterAndTwoPoints _
    Or Figures(Figure1).FigureType = dsCircle_ArcCenterAndRadiusAndTwoPoints _
    Or Figures(Figure1).FigureType = dsAnCircle) _
    Or (Figures(Figure1).FigureType = dsBisector) _
    Then _
        IsVisual = True
End Function

Public Function IsIndependentFigure(ByVal Figure1 As Long) As Boolean
If Figure1 < FigureCount And Figure1 >= 0 Then _
    If (Figures(Figure1).FigureType >= dsAnLineGeneral _
    And Figures(Figure1).FigureType <= dsAnLineNormalPoint) _
    Or (Figures(Figure1).FigureType = dsAnCircle) _
    Then IsIndependentFigure = True
End Function

Public Function IsLine(ByVal Figure1 As Long) As Boolean
If Figure1 < FigureCount And Figure1 >= 0 Then If (Figures(Figure1).FigureType >= dsSegment And Figures(Figure1).FigureType <= dsLine_PointAndPerpendicularLine) Or (Figures(Figure1).FigureType >= dsAnLineGeneral And Figures(Figure1).FigureType <= dsAnLineNormalPoint) Or (Figures(Figure1).FigureType = dsBisector) Then IsLine = True
End Function

Public Function IsLineType(ByVal fType As DrawState) As Boolean
If (fType >= dsSegment And fType <= dsLine_PointAndPerpendicularLine) Or (fType >= dsAnLineGeneral And fType <= dsAnLineNormalPoint) Or (fType = dsBisector) Then IsLineType = True
End Function

Public Function IsCircle(ByVal Figure1 As Long) As Boolean
If Figure1 < FigureCount And Figure1 >= 0 Then If Figures(Figure1).FigureType = dsCircle_CenterAndCircumPoint Or Figures(Figure1).FigureType = dsCircle_CenterAndTwoPoints Or Figures(Figure1).FigureType = dsCircle_ArcCenterAndRadiusAndTwoPoints Or Figures(Figure1).FigureType = dsAnCircle Then IsCircle = True
End Function

Public Function IsCircleType(ByVal fType As DrawState) As Boolean
If fType = dsCircle_CenterAndCircumPoint Or fType = dsCircle_CenterAndTwoPoints Or fType = dsCircle_ArcCenterAndRadiusAndTwoPoints Or fType = dsAnCircle Then IsCircleType = True
End Function

Public Function IsThereATypeMatch(ByVal FType1 As DrawState, ByVal FType2 As DrawState) As Boolean
If (IsCircleType(FType1) And IsCircleType(FType2)) Or (IsLineType(FType1) And IsLineType(FType2)) Then IsThereATypeMatch = True
End Function

Public Function IsFigure(ByVal Figure1 As Long) As Boolean
If Figure1 < FigureCount And Figure1 >= 0 Then IsFigure = True
End Function

Public Function IsPoint(ByVal Point1 As Long) As Boolean
If Point1 > 0 And Point1 <= PointCount Then IsPoint = True
End Function

'======================================================

Public Sub FillFigureWithDefaults(Figure1 As Figure)
With Figure1
    .DrawMode = defFigureDrawMode
    .DrawStyle = defFigureDrawStyle
    .DrawWidth = setdefFigureDrawWidth
    .ForeColor = setdefcolFigure
    .FillColor = colPolygonFillColor
    .FillStyle = 6
    .Hide = False
    .InDemo = True
    .Visible = True
    .XS = ""
    .YS = ""
    .ZOrder = 0
    '.ZOrder = GenerateNewFigureZOrder
End With
End Sub

Public Sub FillPointWithDefaults(Point1 As BasePointType)
With Point1
    .Name = ""
    .Type = dsPoint
    .LabelLength = Len(.Name)
    .LabelOffsetX = 0
    .LabelOffsetY = -setdefPointSize \ 2 + 1
    '.LabelWidth = Paper.TextWidth(.Name)
    '.LabelHeight = Paper.TextHeight(.Name)
    
    .PhysicalWidth = defPointSize
    .Width = .PhysicalWidth
    ToLogicalLength .Width
    
    .Locus = 0
    .ParentFigure = 0
    .ZOrder = 0 'GenerateNewPointZOrder
    .Tag = 0
    
    .FillStyle = setdefPointFill
    .FillColor = setdefcolPointFill
    .ForeColor = setdefcolPoint
    .Shape = setdefPointShape
    .ShowName = setAutoShowPointName
    .ShowCoordinates = False
    .NameColor = setdefcolPointName
    
    .Visible = True
    .Enabled = True
    .Hide = False
    .InDemo = True
    
    .X = 0
    .Y = 0
End With
End Sub

'====================================================
'                               DEPENDENCIES
'====================================================

Public Function FindFigureWithPoints(ByVal FigType As Long, ParamArray Points() As Variant) As Long
Dim Z As Long, Q As Long, i As Long, Counter As Long

FindFigureWithPoints = -1

For Z = 0 To FigureCount - 1
    Counter = 0
    
    For Q = LBound(Points) To UBound(Points)
        For i = 0 To Figures(Z).NumberOfPoints - 1
            If Not IsChildPointPos(Figures(Z), i) Then
                If Points(Q) = Figures(Z).Points(i) Then Counter = Counter + 1
            End If
        Next
    Next
    
    If Counter = UBound(Points) - LBound(Points) + 1 Then
        If FigType <> -1 Then
            If Figures(Z).FigureType = FigType Then
                FindFigureWithPoints = Z
                Exit Function
            End If
        Else
            FindFigureWithPoints = Z
            Exit Function
        End If
    End If

Next
End Function

Public Function GetPointByName(ByVal PointName As String) As Long
Dim Z As Long
If PointCount <> 0 Then
    For Z = 1 To PointCount
        If StrComp(PointName, BasePoint(Z).Name, vbTextCompare) = 0 Then GetPointByName = Z: Exit Function
    Next
End If
End Function

Public Function GetFigureByName(ByVal FigureName As String) As Long
Dim Z As Long
GetFigureByName = -1
If FigureCount <> 0 Then
    For Z = 0 To FigureCount - 1
        If StrComp(FigureName, Figures(Z).Name, vbTextCompare) = 0 Then GetFigureByName = Z: Exit Function
    Next
End If
End Function

Public Function GetProperParentNumber(ByVal FigType As DrawState) As Long
Select Case FigType
    Case dsIntersect
        GetProperParentNumber = 2
    Case dsLine_PointAndParallelLine, dsLine_PointAndPerpendicularLine, dsSimmPointByLine, dsPointOnFigure, dsInvert
        GetProperParentNumber = 1
    Case Else
        GetProperParentNumber = 0
End Select
End Function

Public Function IsParentFigure(Figure1 As Figure, ByVal Figure2 As Long) As Boolean
Dim Z As Long
With Figure1
    If GetProperParentNumber(.FigureType) > 0 Then
        For Z = 0 To GetProperParentNumber(.FigureType) - 1
            If .Parents(Z) = Figure2 Then IsParentFigure = True: Exit Function
        Next
    End If
End With
End Function

Public Function IsChildFigure(Figure1 As Figure, ByVal Figure2 As Long) As Boolean
Dim Z As Long
With Figure1
    If .NumberOfChildren > 0 Then
        For Z = 0 To .NumberOfChildren - 1
            If .Children(Z) = Figure2 Then IsChildFigure = True: Exit Function
        Next
    End If
End With
End Function

Public Function IsParentPoint(Figure1 As Figure, ByVal Point1 As Long) As Boolean
Dim Z As Long
With Figure1
    Select Case .FigureType
        Case dsCircle_ArcCenterAndRadiusAndTwoPoints
            For Z = 0 To 4
                If .Points(Z) = Point1 Then IsParentPoint = True: Exit Function
            Next Z
        Case dsMiddlePoint, dsSimmPoint
            If .Points(1) = Point1 Or .Points(2) = Point1 Then IsParentPoint = True: Exit Function
        Case dsSimmPointByLine, dsInvert
            If .Points(1) = Point1 Then IsParentPoint = True: Exit Function
        Case dsIntersect, dsPointOnFigure
            'do nothing
        Case dsAnPoint
            For Z = 1 To .NumberOfPoints - 1
                If .Points(Z) = Point1 Then IsParentPoint = True: Exit Function
            Next
        Case Else
            If .NumberOfPoints > 0 Then
                For Z = 0 To .NumberOfPoints - 1
                    If .Points(Z) = Point1 Then IsParentPoint = True: Exit Function
                Next
            End If
    End Select
End With
End Function

Public Function IsChildPointPos(Figure1 As Figure, ByVal Point1 As Long) As Boolean
With Figure1
    Select Case .FigureType
        Case dsMiddlePoint, dsSimmPoint, dsSimmPointByLine, dsInvert, dsPointOnFigure, dsAnPoint
            If Point1 = 0 Then IsChildPointPos = True: Exit Function
        Case dsIntersect
            If Point1 = 0 Or Point1 = 1 Then IsChildPointPos = True: Exit Function
        Case dsCircle_ArcCenterAndRadiusAndTwoPoints
            If Point1 = 5 Or Point1 = 6 Then IsChildPointPos = True: Exit Function
    End Select
End With
End Function

Public Function IsChildPoint(Figure1 As Figure, ByVal Point1 As Long) As Boolean
Dim Z As Long
With Figure1
    Select Case .FigureType
        Case dsCircle_ArcCenterAndRadiusAndTwoPoints
            For Z = 5 To 6
                If .Points(Z) = Point1 Then IsChildPoint = True: Exit Function
            Next Z
        Case dsIntersect
            For Z = 0 To 1
                If .Points(Z) = Point1 Then IsChildPoint = True: Exit Function
            Next Z
        Case dsMiddlePoint, dsSimmPoint, dsSimmPointByLine, dsInvert, dsPointOnFigure, dsAnPoint
            If .Points(0) = Point1 Then IsChildPoint = True: Exit Function
    End Select
End With
IsChildPoint = False
End Function

Public Function FigureHasPoint(ByVal Figure1 As Long, ByVal Point1 As Long) As Boolean
Dim Q As Long, Z As Long

For Q = 0 To Figures(Figure1).NumberOfPoints - 1
    If Figures(Figure1).Points(Q) = Point1 Then FigureHasPoint = True: Exit Function
Next Q

Select Case Figures(Figure1).FigureType
    Case dsSegment, dsRay, dsLine_2Points, dsBisector, dsCircle_CenterAndCircumPoint, dsCircle_CenterAndTwoPoints, dsMeasureDistance, dsMeasureAngle
        For Z = 0 To Figures(Figure1).NumberOfPoints - 1
            If BasePoint(Figures(Figure1).Points(Z)).Type <> dsPoint And Not IsChildPointPos(Figures(Figure1), Z) Then
               If FigureHasPoint(BasePoint(Figures(Figure1).Points(Z)).ParentFigure, Point1) Then FigureHasPoint = True: Exit Function
            End If
        Next
        
    Case dsLine_PointAndParallelLine, dsLine_PointAndPerpendicularLine
        If BasePoint(Figures(Figure1).Points(0)).Type <> dsPoint Then
           If FigureHasPoint(BasePoint(Figures(Figure1).Points(0)).ParentFigure, Point1) Then FigureHasPoint = True: Exit Function
        End If
        For Q = 0 To UBound(Figures(Figure1).Parents)
            If FigureHasPoint(Figures(Figure1).Parents(Q), Point1) Then FigureHasPoint = True: Exit Function
        Next Q
    
    Case dsSimmPointByLine, dsInvert
'        For Z = 0 To Figures(Figure1).NumberOfPoints - 1
'            If BasePoint(Figures(Figure1).Points(Z)).Type <> dsPoint Then
'               If FigureHasPoint(BasePoint(Figures(Figure1).Points(Z)).ParentFigure, Point1) Then FigureHasPoint = True: Exit Function
'            End If
'        Next
        If BasePoint(Figures(Figure1).Points(1)).Type <> dsPoint Then
           If FigureHasPoint(BasePoint(Figures(Figure1).Points(1)).ParentFigure, Point1) Then FigureHasPoint = True: Exit Function
        End If
        For Q = 0 To UBound(Figures(Figure1).Parents)
            If FigureHasPoint(Figures(Figure1).Parents(Q), Point1) Then FigureHasPoint = True: Exit Function
        Next Q
        
    Case dsIntersect, dsPointOnFigure
        For Q = 0 To UBound(Figures(Figure1).Parents)
            If FigureHasPoint(Figures(Figure1).Parents(Q), Point1) Then FigureHasPoint = True: Exit Function
        Next Q
        
    Case dsCircle_ArcCenterAndRadiusAndTwoPoints
        For Z = 0 To 4
            If BasePoint(Figures(Figure1).Points(Z)).Type <> dsPoint Then
               If FigureHasPoint(BasePoint(Figures(Figure1).Points(Z)).ParentFigure, Point1) Then FigureHasPoint = True: Exit Function
            End If
        Next
    
    Case dsMiddlePoint, dsSimmPoint
        For Z = 1 To 2
            If BasePoint(Figures(Figure1).Points(Z)).Type <> dsPoint Then
               If FigureHasPoint(BasePoint(Figures(Figure1).Points(Z)).ParentFigure, Point1) Then FigureHasPoint = True: Exit Function
            End If
        Next
        
    Case dsAnPoint
        For Z = 1 To Figures(Figure1).NumberOfPoints - 1
            If BasePoint(Figures(Figure1).Points(Z)).Type <> dsPoint Then
               If FigureHasPoint(BasePoint(Figures(Figure1).Points(Z)).ParentFigure, Point1) Then FigureHasPoint = True: Exit Function
            End If
        Next
    
    Case dsDynamicLocus
        If FigureHasPoint(BasePoint(Figures(Figure1).Points(1)).ParentFigure, Point1) Then FigureHasPoint = True: Exit Function
        
End Select
End Function

Public Function FigureDependsOnFigure(ByVal Figure1 As Long, ByVal Figure2 As Long) As Boolean
Dim Z As Long

If Figure1 = Figure2 Then FigureDependsOnFigure = True: Exit Function
For Z = 0 To GetProperParentNumber(Figures(Figure1).FigureType) - 1
    If FigureDependsOnFigure(Figures(Figure1).Parents(Z), Figure2) Then FigureDependsOnFigure = True: Exit Function
Next

For Z = 0 To Figures(Figure1).NumberOfPoints - 1
    If Not IsChildPointPos(Figures(Figure1), Z) Then
        If BasePoint(Figures(Figure1).Points(Z)).Type <> dsPoint Then
            If BasePoint(Figures(Figure1).Points(Z)).ParentFigure <> Figure1 Then
                If FigureDependsOnFigure(BasePoint(Figures(Figure1).Points(Z)).ParentFigure, Figure2) Then FigureDependsOnFigure = True: Exit Function
            End If
        End If
    End If
Next

FigureDependsOnFigure = False
End Function

Public Sub RemoveFromChildren(Figure1 As Figure, ByVal Index As Long)
Dim Z As Long, Q As Long

Do While Z <= Figure1.NumberOfChildren - 1
    If Figure1.Children(Z) = Index Then
        If Z < Figure1.NumberOfChildren - 1 Then
            For Q = Z To Figure1.NumberOfChildren - 2
                Figure1.Children(Q) = Figure1.Children(Q + 1)
            Next
        End If
        Figure1.NumberOfChildren = Figure1.NumberOfChildren - 1
        If Figure1.NumberOfChildren > 0 Then
            ReDim Preserve Figure1.Children(0 To Figure1.NumberOfChildren - 1)
        Else
            ReDim Figure1.Children(0 To 0)
        End If
    Else
        Z = Z + 1
    End If
Loop
End Sub

'=========================================================

Public Sub PointOnFigure2Basepoint(ByVal Point1 As Long, Optional ByVal ShouldRecord As Boolean = True)
Dim Z As Long
Dim tempPoint As BasePointType, pAction As Action

ShowPoint Paper.hDC, Point1, True
pAction.Type = actReleasePoint
ReDim pAction.sFigure(1 To 1)
pAction.sFigure(1) = Figures(BasePoint(Point1).ParentFigure)
ReDim pAction.sPoint(1 To 1)
pAction.sPoint(1) = BasePoint(Point1)
pAction.pPoint = Point1
pAction.pFigure = Figures(BasePoint(Point1).ParentFigure).Parents(0)

Z = 0
Do While Z < FigureCount
    If Figures(Z).FigureType = dsDynamicLocus Then
        If Figures(Z).Points(1) = Point1 Or Figures(Z).Points(0) = Point1 Then DeleteFigure Z Else Z = Z + 1
    Else
        Z = Z + 1
    End If
Loop

If ShouldRecord Then RecordGenericAction ResUndoReleasePoint

DeleteFigure BasePoint(Point1).ParentFigure, False, False, True

BasePoint(Point1).Type = dsPoint
BasePoint(Point1).ParentFigure = 0
'BasePoint(Point1).ForeColor = setdefcolPoint
'BasePoint(Point1).Shape = defPointShape
PaperCls
ShowAll
End Sub

Public Sub Basepoint2PointOnFigure(ByVal Point1 As Long, ByVal Figure1 As Long, Optional ByVal ShouldRecord As Boolean = True)
Dim P As OnePoint, IPointCount As Long, NewFigure As Long, pAction As Action

ReDim Preserve Figures(0 To FigureCount)
'If Figure1 < FigureCount - 1 Then
'    NewFigure = Figure1 + 1
'    For Z = FigureCount To Figure1 + 1 Step -1
'        Figures(Z) = Figures(Z - 1)
'    Next
'    For Z = 1 To PointCount
'        If BasePoint(Z).ParentFigure >= NewFigure Then BasePoint(Z).ParentFigure = BasePoint(Z).ParentFigure + 1
'    Next
'    For Z = 0 To FigureCount
'        If Figures(Z).NumberOfChildren > 0 Then
'            For Q = 0 To Figures(Z).NumberOfChildren - 1
'                If Figures(Z).Children(Q) >= NewFigure Then Figures(Z).Children(Q) = Figures(Z).Children(Q) + 1
'            Next Q
'        End If
'        If GetProperParentNumber(Figures(Z).FigureType) > 0 Then
'            For Q = 0 To GetProperParentNumber(Figures(Z).FigureType) - 1
'                If Figures(Z).Parents(Q) >= NewFigure Then Figures(Z).Parents(Q) = Figures(Z).Parents(Q) + 1
'            Next Q
'        End If
'    Next Z
'Else

NewFigure = FigureCount

'End If
'pAction.Type = actSnapPoint
'pAction.pPoint = Point1
'If ShouldRecord Then RecordAction pAction

RecordGenericAction ResUndoSnapPoint

IPointCount = AddFigureAux(NewFigure, dsPointOnFigure, , False)

Figures(NewFigure).NumberOfPoints = 1

P = GetPointOnFigure(Figure1, BasePoint(Point1).X, BasePoint(Point1).Y)
ReDim Figures(NewFigure).Points(0 To 0)
ReDim Figures(NewFigure).Parents(0 To 0)
Figures(NewFigure).Parents(0) = Figure1
Figures(NewFigure).NumberOfChildren = 0
Figures(NewFigure).AuxInfo(5) = 1

ReDim Preserve Figures(Figure1).Children(0 To Figures(Figure1).NumberOfChildren)
Figures(Figure1).Children(Figures(Figure1).NumberOfChildren) = NewFigure
Figures(Figure1).NumberOfChildren = Figures(Figure1).NumberOfChildren + 1

ShowPoint Paper.hDC, Point1, True

BasePoint(Point1).Type = dsPointOnFigure
BasePoint(Point1).ParentFigure = NewFigure

Figures(NewFigure).Points(0) = Point1
Figures(NewFigure).ForeColor = setdefcolFigurePoint
'If setdefcolDependentPoint <> setdefcolFigurePoint Or setdefPointShape <> defSemiDependentPointShape Then
    'ShowPoint Paper.hDC, Point1, True
    'BasePoint(Point1).Shape = defSemiDependentPointShape
    'BasePoint(Point1).ForeColor = Figures(NewFigure).ForeColor
    'ShowPoint Paper.hDC, Point1
'End If

If P.X <> EmptyVar And P.Y <> EmptyVar Then
    'AddBasePoint P.X, P.Y, , dsPointOnFigure, , FigureCount
    MovePoint Point1, P.X, P.Y
    ShowPoint Paper.hDC, Point1
Else
    'AddBasePoint 0, 0, , dsPointOnFigure, , FigureCount
    ShowPoint Paper.hDC, Point1, True
    BasePoint(Point1).Visible = False
End If

FigureCount = FigureCount + 1
RecalcSemiDependentInfo NewFigure, P.X, P.Y

RecalcAllAuxInfo
PaperCls
ShowAll
'Dim P As OnePoint, IPointCount As Long, NewFigure As Long
'ReDim Preserve Figures(0 To FigureCount)
'
'IPointCount = AddFigureAux(FigureCount, dsPointOnFigure)
'
'Figures(FigureCount).NumberOfPoints = 1
'
'P = GetPointOnFigure(Figure1, BasePoint(Point1).X, BasePoint(Point1).Y)
'ReDim Figures(FigureCount).Points(0 To 0)
'ReDim Figures(FigureCount).Parents(0 To 0)
'Figures(FigureCount).Parents(0) = Figure1
'Figures(FigureCount).NumberOfChildren = 0
'
'ReDim Preserve Figures(Figure1).Children(0 To Figures(Figure1).NumberOfChildren)
'Figures(Figure1).Children(Figures(Figure1).NumberOfChildren) = FigureCount
'Figures(Figure1).NumberOfChildren = Figures(Figure1).NumberOfChildren + 1
'
'ShowPoint Paper.hDC, Point1, True
'
'BasePoint(Point1).Type = dsPointOnFigure
'BasePoint(Point1).ParentFigure = FigureCount
'
'Figures(FigureCount).Points(0) = Point1
'Figures(FigureCount).Letters(0) = BasePoint(Point1).Name
'Figures(FigureCount).ForeColor = setdefcolFigurePoint
'If setdefcolDependentPoint <> setdefcolFigurePoint Or setdefPointShape <> defSemiDependentPointShape Then
'    'ShowPoint Paper.hDC, Point1, True
'    BasePoint(Point1).Shape = defSemiDependentPointShape
'    BasePoint(Point1).ForeColor = Figures(FigureCount).ForeColor
'    'ShowPoint Paper.hDC, Point1
'End If
'
'If P.X <> EmptyVar And P.Y <> EmptyVar Then
'    'AddBasePoint P.X, P.Y, , dsPointOnFigure, , FigureCount
'    MovePoint Point1, P.X, P.Y
'    ShowPoint Paper.hDC, Point1
'Else
'    'AddBasePoint 0, 0, , dsPointOnFigure, , FigureCount
'    ShowPoint Paper.hDC, Point1, True
'    BasePoint(Point1).Visible = False
'End If
'
'FigureCount = FigureCount + 1
'RecalcSemiDependentInfo FigureCount - 1, P.X, P.Y
'
'RecalcAllAuxInfo
'PaperCls
'ShowAll
End Sub

'====================================================
'====================================================


'====================================================
'====================================================

