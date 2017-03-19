Attribute VB_Name = "modLocus"
'Implementation of loci functionality
'Static and dynamic loci
'Creation, operation, destruction

Option Explicit

'#####################################################
'Adds an element to the global loci array; initializes it
'#####################################################

Public Sub AddLocus(ByVal Point1 As Long)
Dim tX As Double, tY As Double

If Not IsPoint(Point1) Then Exit Sub

LocusCount = LocusCount + 1
ReDim Preserve Locuses(1 To LocusCount)
Locuses(LocusCount).Enabled = True
Locuses(LocusCount).ForeColor = setdefcolLocus
Locuses(LocusCount).ParentPoint = Point1
Locuses(LocusCount).Visible = True
Locuses(LocusCount).DrawWidth = setdefLocusDrawWidth
Locuses(LocusCount).Type = setdefLocusType
Locuses(LocusCount).Description = GetObjectDescription(gotLocus, LocusCount)

ReDim Locuses(LocusCount).LocusPoints(1 To 1)
ReDim Locuses(LocusCount).LocusPixels(1 To 1)
ReDim Locuses(LocusCount).LocusNumbers(1 To 1)
Locuses(LocusCount).LocusPointCount = 1
Locuses(LocusCount).LocusNumber = 1
Locuses(LocusCount).LocusNumbers(1) = 1

tX = BasePoint(Point1).X
tY = BasePoint(Point1).Y
Locuses(LocusCount).LocusPoints(1).X = tX
Locuses(LocusCount).LocusPoints(1).Y = tY
ToPhysical tX, tY
Locuses(LocusCount).LocusPixels(1).X = tX
Locuses(LocusCount).LocusPixels(1).Y = tY

BasePoint(Point1).Locus = LocusCount
End Sub

'#####################################################
'Check sanity of dynamic locus candidate
'#####################################################

Public Function CanAddDynamicLocus(ByVal Point1 As Long, ByVal FigurePoint1 As Long) As Boolean
If Not IsPoint(Point1) Or Not IsPoint(FigurePoint1) Then Exit Function
If Point1 = FigurePoint1 Then Exit Function
If BasePoint(FigurePoint1).Type <> dsPointOnFigure Or BasePoint(Point1).Type = dsPoint Then Exit Function
If Not FigureDependsOnFigure(BasePoint(Point1).ParentFigure, BasePoint(FigurePoint1).ParentFigure) Then Exit Function

CanAddDynamicLocus = True
End Function

'#####################################################
'Adds a dynamic locus figure, also adds a corresponding locus object
'#####################################################

Public Sub AddDynamicLocus(ByVal Point1 As Long, ByVal FigurePoint1 As Long)
On Local Error GoTo EH
Dim Z As Long

If Point1 = FigurePoint1 Then Exit Sub

For Z = 0 To FigureCount - 1
    If Figures(Z).FigureType = dsDynamicLocus Then
        If Figures(Z).Points(0) = Point1 Then DeleteFigure Z: Exit For
    End If
Next

RedimPreserveFigures 0, FigureCount
AddFigureAux FigureCount, dsDynamicLocus

If BasePoint(FigurePoint1).Type <> dsPointOnFigure Then Exit Sub
If BasePoint(Point1).Type = dsPoint Then Exit Sub

Figures(FigureCount).NumberOfPoints = 2
ReDim Figures(FigureCount).Points(0 To 1)
Figures(FigureCount).Points(0) = Point1
Figures(FigureCount).Points(1) = FigurePoint1

If BasePoint(Point1).Locus = 0 Then
    AddLocus Point1
Else
    EraseLocus BasePoint(Point1).Locus
End If

With Locuses(BasePoint(Point1).Locus)
    .Visible = True
    .Enabled = True
    .Dynamic = True
    .LocusPointCount = setLocusDetails
    ReDim .LocusPixels(1 To .LocusPointCount)
    ReDim .LocusPoints(1 To .LocusPointCount)
    ReDim .LocusNumbers(1 To 1)
    .LocusNumber = 1
    .ParentFigure = FigureCount
End With

BuildDynamicLocusDependency Figures(FigureCount), Point1, FigurePoint1
Figures(FigureCount).Description = GetObjectDescription(gotFigure, FigureCount)

FigureCount = FigureCount + 1
RecalcAuxInfo FigureCount - 1

If Figures(FigureCount - 1).AuxInfo(1) = 0 Then DeleteFigure FigureCount - 1
Exit Sub

EH:
End Sub

'#####################################################
'Adds a point to the existing locus object
'#####################################################

Public Sub AddPointToLocus(ByVal Locus1 As Long, ByVal X As Double, ByVal Y As Double, Optional ByVal ShouldBreak As Boolean = False)
On Local Error GoTo EH
Dim LPC As Long

If Locus1 < 1 Or Locus1 > LocusCount Then Exit Sub

With Locuses(Locus1)
    If Abs(X) > 1000 Or Abs(Y) > 1000 Then
        .ShouldBreak = True
        Exit Sub
    End If
    
    If ShouldBreak Then .ShouldBreak = True
    
    If .LocusNumbers(.LocusNumber) > 1 And Not .ShouldBreak Then
        If .LocusPixels(.LocusPointCount).X = .LocusPixels(.LocusPointCount - 1).X And .LocusPixels(.LocusPointCount).Y = .LocusPixels(.LocusPointCount - 1).Y Then
            .LocusPointCount = .LocusPointCount - 1
            .LocusNumbers(.LocusNumber) = .LocusNumbers(.LocusNumber) - 1
        End If
    End If
    
    LPC = .LocusPointCount + 1
    .LocusPointCount = LPC
    ReDim Preserve .LocusPoints(1 To LPC)
    ReDim Preserve .LocusPixels(1 To LPC)
    .LocusPoints(LPC).X = X
    .LocusPoints(LPC).Y = Y
    ToPhysical X, Y
    .LocusPixels(LPC).X = X
    .LocusPixels(LPC).Y = Y
    
    If .ShouldBreak Then
        If .LocusNumbers(.LocusNumber) < 2 Then
            If .LocusNumbers(.LocusNumber) = 1 Then
                LPC = .LocusPointCount + 1
                .LocusPointCount = LPC
                ReDim Preserve .LocusPoints(1 To LPC)
                ReDim Preserve .LocusPixels(1 To LPC)
                .LocusPoints(LPC).X = X
                .LocusPoints(LPC).Y = Y
                ToPhysical X, Y
                .LocusPixels(LPC).X = X
                .LocusPixels(LPC).Y = Y
            Else
                .LocusNumber = .LocusNumber - 1
            End If
        End If
        
        .LocusNumber = .LocusNumber + 1
        ReDim Preserve .LocusNumbers(1 To .LocusNumber)
        .LocusNumbers(.LocusNumber) = 2
    
        LPC = .LocusPointCount + 1
        .LocusPointCount = LPC
        ReDim Preserve .LocusPoints(1 To LPC)
        ReDim Preserve .LocusPixels(1 To LPC)
        .LocusPoints(LPC).X = .LocusPoints(LPC - 1).X
        .LocusPoints(LPC).Y = .LocusPoints(LPC - 1).Y
        .LocusPixels(LPC).X = .LocusPixels(LPC - 1).X
        .LocusPixels(LPC).Y = .LocusPixels(LPC - 1).Y
        .ShouldBreak = False
    Else
        .LocusNumbers(.LocusNumber) = .LocusNumbers(.LocusNumber) + 1
    End If
End With
Exit Sub

EH:
If LPC > 0 Then
    Locuses(Locus1).LocusPointCount = LPC - 1
    If Locuses(Locus1).LocusNumbers(Locuses(Locus1).LocusNumber) > 0 Then Locuses(Locus1).LocusNumbers(Locuses(Locus1).LocusNumber) = Locuses(Locus1).LocusNumbers(Locuses(Locus1).LocusNumber) - 1
End If
End Sub

'#####################################################
'Begins a new locus piece
'#####################################################

Public Sub AddLocusDiscontinuity(ByVal Locus1 As Long)
If Locus1 < 1 Or Locus1 > LocusCount Then Exit Sub

With Locuses(Locus1)
    If .LocusNumbers(.LocusNumber) < 2 Then
        
    End If
    .LocusNumber = .LocusNumber + 1
    ReDim Preserve .LocusNumbers(1 To .LocusNumber)
    .LocusNumbers(.LocusNumber) = 0
End With
End Sub

'#####################################################
'Self-explanatory
'#####################################################

Public Sub BuildDynamicLocusDependency(Figure1 As Figure, ByVal Point1 As Long, ByVal FigurePoint1 As Long)
Dim ArSize As Long, Z As Long, Q As Long

ArSize = 0
For Z = 0 To FigureCount - 1
    If FigureDependsOnFigure(BasePoint(Point1).ParentFigure, Z) Then
         If FigureDependsOnFigure(Z, BasePoint(FigurePoint1).ParentFigure) Then
            If Figures(Z).Name <> Figure1.Name Then
                ArSize = ArSize + 1
                ReDim Preserve Figure1.AuxArray(1 To ArSize)
                Figure1.AuxArray(ArSize) = Z
            End If
         End If
     End If
Next
If ArSize > 1 Then
    For Z = 1 To ArSize - 1
        For Q = Z + 1 To ArSize
            If FigureDependsOnFigure(Figure1.AuxArray(Z), Figure1.AuxArray(Q)) Then
                Swap Figure1.AuxArray(Z), Figure1.AuxArray(Q)
            End If
        Next
    Next
End If

Figure1.AuxInfo(1) = ArSize
End Sub

'#####################################################
'Self-explanatory
'#####################################################

Public Sub RebuildDynamicLocusDependencies()
Dim Z As Long

For Z = 0 To FigureCount - 1
    If Figures(Z).FigureType = dsDynamicLocus Then
        BuildDynamicLocusDependency Figures(Z), Figures(Z).Points(0), Figures(Z).Points(1)
    End If
Next
End Sub

'###############################################################
'Deletes current locus element from Locuses()
'###############################################################

Public Sub DeleteLocus(ByVal Index As Long, Optional ByVal NeedToShowAll As Boolean = True)
Dim Z As Long

DeleteFromDependentButtons Index, gotLocus, False

ShowLocus Paper.hDC, BasePoint(Index).Locus, False, True
If BasePoint(Index).Locus < LocusCount Then
    For Z = BasePoint(Index).Locus To LocusCount - 1
        Locuses(Z) = Locuses(Z + 1)
        OffsetButtonObjectDependencies Z + 1, Z, gotLocus
        BasePoint(Locuses(Z).ParentPoint).Locus = BasePoint(Locuses(Z).ParentPoint).Locus - 1
    Next Z
End If
BasePoint(Index).Locus = 0
LocusCount = LocusCount - 1
If LocusCount > 0 Then
    ReDim Preserve Locuses(1 To LocusCount)
Else
    ReDim Locuses(1 To 1)
End If
If NeedToShowAll Then ShowAll
End Sub

'###############################################################
'Clears the locus
'Prepares to build dynamic locus, if needed
'###############################################################

Public Sub EraseLocus(ByVal Index As Long, Optional ByVal ShouldShowAll As Boolean = True)
Dim pActionLocus As Action

If Index < 1 Or Index > LocusCount Then Exit Sub

With Locuses(Index)
    ReDim pActionLocus.sLocus(1 To 1)
    pActionLocus.sLocus(1) = Locuses(Index)
    
    ShowLocus Paper.hDC, Index, False, True
    
    ReDim .LocusPoints(1 To 1)
    ReDim .LocusPixels(1 To 1)
    ReDim .LocusNumbers(1 To 1)
    .LocusNumber = 1
    .LocusPointCount = 0
    .LocusNumbers(1) = 0
    
    pActionLocus.pPoint = ActivePoint
    pActionLocus.Type = actRemoveLocus
    pActionLocus.pLocus = Index
    RecordAction pActionLocus
    
    If ShouldShowAll Then PaperCls: ShowAll
End With
End Sub

'###############################################################
'Retrieves index of locus under point (X,Y). Returns 0 if finds none.
'###############################################################

Public Function GetLocusFromPoint(ByVal X As Double, ByVal Y As Double) As Long
Dim Z As Long, P As OnePoint

For Z = 1 To LocusCount
    P = GetPerpPointPolyline(X, Y, Locuses(Z).LocusPoints)
    If Distance(X, Y, P.X, P.Y) < Sensitivity And Not Locuses(Z).Hide Then GetLocusFromPoint = Z: Exit Function
Next
End Function

'###############################################################
'###############################################################
Public Function IsLocus(ByVal Index As Long) As Boolean
IsLocus = Index >= 1 And Index <= LocusCount
End Function

'###############################################################
'Calculates the amount of memory taken by the locus data structure
'###############################################################
Public Function LocusMemoryConsumption(pLocus As Locus) As Long
Dim MemSum As Long
MemSum = MemSum + Len(pLocus)
MemSum = MemSum + 24 * pLocus.LocusPointCount
MemSum = MemSum + 4 * pLocus.LocusNumber
LocusMemoryConsumption = MemSum
End Function

'###############################################################
'Rebuilds dynamic locus trajectory according to changes in the drawing
'###############################################################
Public Sub RecalcDynamicLocus(ByVal Figure1 As Long, Optional ByVal RenderHighQuality As Boolean = False)
Dim T As Double, A As Double, B As Double, CurPoint As Long, tX As Double, tY As Double, OldX As Double, OldY As Double
Dim TP As TwoPoints, OldT As Double, Z As Long
Dim Carrier As Long, FigurePoint As Long, St As Double
Dim OldManualDragFlag As Long
Dim LPC As Long, CurLocus As Long
Dim Gap As Double, MaxGap As Double

FigurePoint = BasePoint(Figures(Figure1).Points(1)).ParentFigure
Carrier = Figures(FigurePoint).Parents(0)
CurLocus = BasePoint(Figures(Figure1).Points(0)).Locus
Locuses(CurLocus).Visible = True

OldT = Figures(FigurePoint).AuxInfo(1)
OldManualDragFlag = ManualDragFlag
ManualDragFlag = 0

Select Case Figures(Carrier).FigureType
    Case dsSegment
        A = 0
        B = 1
    Case dsRay
        A = 0
        TP = GetLineCoordinatesAbsolute(Carrier)
        If (TP.P1.X = TP.P2.X And TP.P1.Y = TP.P2.Y) Then B = 0 Else B = Abs(CanvasBorders.P2.X - CanvasBorders.P1.X) / Distance(TP.P1.X, TP.P1.Y, TP.P2.X, TP.P2.Y)
    Case dsLine_2Points, dsLine_PointAndParallelLine, dsBisector, dsLine_PointAndPerpendicularLine, dsAnLineCanonic, dsAnLineGeneral, dsAnLineNormal, dsAnLineNormalPoint
        TP = GetLineCoordinatesAbsolute(Carrier)
        If (TP.P1.X = TP.P2.X And TP.P1.Y = TP.P2.Y) Then
            A = 0
            B = 0
        Else
            B = (Abs(CanvasBorders.P2.X - CanvasBorders.P1.X)) / Distance(TP.P1.X, TP.P1.Y, TP.P2.X, TP.P2.Y)
            A = -B
        End If
    Case dsCircle_CenterAndCircumPoint, dsCircle_CenterAndTwoPoints, dsAnCircle
        A = 0
        B = 2 * PI
    Case dsCircle_ArcCenterAndRadiusAndTwoPoints
        A = Figures(Carrier).AuxInfo(2)
        B = Figures(Carrier).AuxInfo(3)
        If B <= A Then B = B + 2 * PI
    Case dsDynamicLocus
        A = Figures(Carrier).AuxInfo(2)
        B = Figures(Carrier).AuxInfo(3)
End Select

If B = A Then
    Locuses(BasePoint(Figures(Figure1).Points(0)).Locus).Visible = False
    Exit Sub
End If

Figures(Figure1).AuxInfo(2) = A
Figures(Figure1).AuxInfo(3) = B

If RenderHighQuality Then
    With Locuses(CurLocus)
        .LocusPointCount = setLocusDetailsHigh
        ReDim .LocusPoints(1 To .LocusPointCount)
        ReDim .LocusPixels(1 To .LocusPointCount)
        LPC = .LocusPointCount
    End With
    St = (B - A) / (setLocusDetailsHigh - 1)
Else
    With Locuses(CurLocus)
        If .LocusPointCount <> setLocusDetails Then
            .LocusPointCount = setLocusDetails
            ReDim .LocusPoints(1 To .LocusPointCount)
            ReDim .LocusPixels(1 To .LocusPointCount)
        End If
        LPC = .LocusPointCount
    End With
    St = (B - A) / (setLocusDetails - 1)
End If

CurPoint = 0
With Locuses(CurLocus)
    .LocusNumber = 1
    ReDim .LocusNumbers(1 To 1)
    .ShouldBreak = False
    
    MaxGap = Maximum(Abs(CanvasBorders.P1.X - CanvasBorders.P2.X), Abs(CanvasBorders.P1.Y - CanvasBorders.P2.Y)) / 2
    
    For T = A To B + St Step St 'To B + St //why was that needed???
        Figures(FigurePoint).AuxInfo(1) = T
        'RecalcDynamicLocusChildren Figure1, CurPoint, CurLocus
        
        For Z = 1 To Figures(Figure1).AuxInfo(1)
            RecalcPureAuxInfo Figures(Figure1).AuxArray(Z)
        Next
        
        tX = BasePoint(Figures(Figure1).Points(0)).X
        tY = BasePoint(Figures(Figure1).Points(0)).Y
        
        If CurPoint < Locuses(CurLocus).LocusPointCount And Abs(tX) < 1000 And Abs(tY) < 1000 And BasePoint(Figures(Figure1).Points(0)).Visible Then
            If CurPoint > 0 And Not .ShouldBreak Then
                Gap = Distance(tX, tY, .LocusPoints(CurPoint).X, .LocusPoints(CurPoint).Y)
                If Gap >= MaxGap Then .ShouldBreak = True
            End If
            
            If .ShouldBreak Then
                If .LocusNumbers(.LocusNumber) < 2 Then
                    If .LocusNumbers(.LocusNumber) = 1 Then
                        CurPoint = CurPoint + 1
                        .LocusPoints(CurPoint).X = .LocusPoints(CurPoint - 1).X
                        .LocusPoints(CurPoint).Y = .LocusPoints(CurPoint - 1).Y
                        .LocusPixels(CurPoint).X = .LocusPixels(CurPoint - 1).X
                        .LocusPixels(CurPoint).Y = .LocusPixels(CurPoint - 1).Y
                        .LocusNumbers(.LocusNumber) = 2
                    Else
                        .LocusNumber = .LocusNumber - 1
                    End If
                End If
                .LocusNumber = .LocusNumber + 1
                ReDim Preserve .LocusNumbers(1 To .LocusNumber)
                .LocusNumbers(.LocusNumber) = 0
                .ShouldBreak = False
            End If
            
            If CurPoint < Locuses(CurLocus).LocusPointCount Then
                CurPoint = CurPoint + 1
                .LocusPoints(CurPoint).X = tX
                .LocusPoints(CurPoint).Y = tY
                ToPhysical tX, tY
                .LocusPixels(CurPoint).X = tX
                .LocusPixels(CurPoint).Y = tY
                .LocusNumbers(.LocusNumber) = .LocusNumbers(.LocusNumber) + 1
            End If
        Else
            If CurPoint < Locuses(CurLocus).LocusPointCount Then .ShouldBreak = True
        End If
        
    Next T
    
    If .LocusNumbers(.LocusNumber) = 1 Then
        CurPoint = CurPoint - 1
        If .LocusNumber > 1 Then
            .LocusNumber = .LocusNumber - 1
            ReDim Preserve .LocusNumbers(1 To .LocusNumber)
        End If
    End If
    .LocusPointCount = CurPoint
    ReDim Preserve .LocusPoints(1 To .LocusPointCount)
    ReDim Preserve .LocusPixels(1 To .LocusPointCount)
    
End With


Figures(FigurePoint).AuxInfo(1) = OldT
RecalcAuxInfo FigurePoint

For Z = 1 To Figures(Figure1).AuxInfo(1)
    RecalcAuxInfo Figures(Figure1).AuxArray(Z)
Next

'Figures(FigurePoint).AuxInfo(1) = OldT
'RecalcAuxInfo FigurePoint
ManualDragFlag = OldManualDragFlag
End Sub
'
'Public Sub RecalcDynamicLocusChildren(Figure1 As Long, CurPoint As Long, CurLocus As Long)
'Dim tX As Double, tY As Double
'For Z = 1 To Figures(Figure1).AuxInfo(1)
'    RecalcPureAuxInfo Figures(Figure1).AuxArray(Z)
'Next
'If CurPoint <= Locuses(CurLocus).LocusPointCount Then
'    tX = BasePoint(Figures(Figure1).Points(0)).X
'    tY = BasePoint(Figures(Figure1).Points(0)).Y
'    If Abs(tX) < 1000 And Abs(tY) < 1000 Then
'        Locuses(CurLocus).LocusPoints(CurPoint).X = tX
'        Locuses(CurLocus).LocusPoints(CurPoint).Y = tY
'        ToPhysical tX, tY
'        Locuses(CurLocus).LocusPixels(CurPoint).X = tX
'        Locuses(CurLocus).LocusPixels(CurPoint).Y = tY
'    End If
'End If
'
'For Z = 0 To Figures(Figure1).NumberOfChildren - 1
'    RecalcDynamicLocusChildren Figures(Figure1).Children(Z), CurPoint, CurLocus
'Next
'End Sub

'###############################################################
'Recalcs physical pixels in locus point array
'###############################################################
Public Sub RecalcLocus(ByVal Index As Long)
Dim tX As Double, tY As Double, Z As Long
With Locuses(Index)
    For Z = 1 To .LocusPointCount
        tX = .LocusPoints(Z).X
        tY = .LocusPoints(Z).Y
        ToPhysical tX, tY
        .LocusPixels(Z).X = tX
        .LocusPixels(Z).Y = tY
    Next
End With
End Sub

'###############################################################
'You know what it does :-)
'###############################################################
Public Sub RecalcLocuses()
Dim Z As Long
For Z = 1 To LocusCount
    RecalcLocus Z
Next
End Sub

'###############################################################
'Remove last point from locus
'If necessary, also deletes empty locus piece
'###############################################################
Public Sub RemovePointFromLocus(ByVal Locus1 As Long)
Dim LPC As Long

If Locus1 < 1 Or Locus1 > LocusCount Then Exit Sub

With Locuses(Locus1)
    LPC = .LocusPointCount
    If LPC = 0 Then Exit Sub
    
    .LocusPointCount = LPC - 1
    .LocusNumbers(.LocusNumber) = .LocusNumbers(.LocusNumber) - 1
    If LPC = 1 Then Exit Sub
    
    ReDim Preserve .LocusPoints(1 To LPC - 1)
    ReDim Preserve .LocusPixels(1 To LPC - 1)
    If .LocusNumbers(.LocusNumber) = 0 Then
        .LocusNumber = .LocusNumber - 1
        ReDim Preserve .LocusNumbers(1 To .LocusNumber)
    End If
End With
End Sub

'###############################################################
'Called when user selects "Clear Locus" from popup menu
'###############################################################

Public Sub ClearLocusMenu(ByVal Point1 As Long)
Dim Z As Long

If Not IsPoint(Point1) Then Exit Sub
If BasePoint(Point1).Locus = 0 Then Exit Sub

If Locuses(BasePoint(Point1).Locus).Dynamic Then
    DeleteFigure Locuses(BasePoint(Point1).Locus).ParentFigure
Else
    EraseLocus BasePoint(Point1).Locus
End If
End Sub

'###############################################################
'Called when user selects "Create Locus" from popup menu
'###############################################################

Public Sub CreateLocusMenu(ByVal Point1 As Long)
Dim pAction As Action

If Not IsPoint(Point1) Then Exit Sub

If BasePoint(Point1).Locus = 0 Then
    AddLocus Point1
    pAction.pPoint = Point1
    pAction.Type = actAddLocus
    pAction.pLocus = LocusCount
    RecordAction pAction
    Exit Sub
Else
    pAction.Type = actChangeAttrLocus
    pAction.pLocus = BasePoint(Point1).Locus
    pAction.pPoint = Point1
    ReDim pAction.sLocus(1 To 1)
    pAction.sLocus(1) = Locuses(BasePoint(Point1).Locus)
    RecordAction pAction
    
    Locuses(BasePoint(Point1).Locus).Enabled = Not Locuses(BasePoint(Point1).Locus).Enabled
    If Not Locuses(BasePoint(Point1).Locus).Enabled Then Locuses(BasePoint(Point1).Locus).ShouldBreak = True
End If
End Sub

Public Function GetLocusParentFigure(ByVal Point1 As Long)
Dim Z As Long

For Z = 0 To FigureCount - 1
    If Figures(Z).FigureType = dsDynamicLocus Then
        If Figures(Z).Points(0) = Point1 Then
            GetLocusParentFigure = Z
            Exit Function
        End If
    End If
Next

GetLocusParentFigure = -1
End Function
