Attribute VB_Name = "modDemo"
Option Explicit

Public Sub AutorunDemo()
If FormMain.tmrDemo.Enabled Then
    FormMain.tmrDemo.Enabled = False
Else
    FormMain.tmrDemo.Enabled = True
End If
DrawSituation
End Sub

Public Sub RunDemo()
LinearObjectListClear DemoSequence
DemoSequence = GenerateDemoSequence
If DemoSequence.Current = 0 Then
    MsgBox GetString(ResEmptyDemo), vbExclamation
    Exit Sub
End If
FormMain.EnterDemoMode
End Sub

Public Sub DemoFirstStep()
LinearObjectListFrameFirst DemoSequence
DrawSituation
End Sub

Public Sub DemoPreviousStep()
LinearObjectListFramePrevious DemoSequence
DrawSituation
End Sub

Public Sub DemoNextStep()
LinearObjectListFrameNext DemoSequence
DrawSituation
End Sub

Public Sub DemoLastStep()
LinearObjectListFrameLast DemoSequence
DrawSituation
End Sub

Public Sub EndDemo()
FormMain.ExitDemoMode
End Sub

Public Sub DemoShowDescription()
Dim S As String
S = GetString(ResDes_StepAofB)
S = Replace(S, "%1", DemoSequence.Step)
S = Replace(S, "%2", DemoSequence.StepCount)
S = S & " " & DemoSequence.Items(DemoSequence.Current).Description
FormMain.ShowStatusSpecial S
End Sub

Public Function GenerateDemoSequence() As LinearObjectList
Dim L As LinearObjectList
Dim Z As Long

For Z = 0 To FigureCount - 1
    AddFigureToDemoSequence L, Z
Next Z

For Z = 1 To PointCount
    AddPointToDemoSequence L, Z
Next Z

For Z = 1 To StaticGraphicCount
    LinearObjectListAdd L, gotSG, Z, StaticGraphics(Z).InDemo, StaticGraphics(Z).Description
Next

For Z = 1 To LabelCount
    LinearObjectListAdd L, gotLabel, Z, TextLabels(Z).InDemo, TextLabels(Z).Description
Next

For Z = 1 To ButtonCount
    LinearObjectListAdd L, gotButton, Z, Buttons(Z).InDemo, Buttons(Z).Description
Next

For Z = 1 To L.Count
    If L.Items(Z).Participate Then L.StepCount = L.StepCount + 1
Next

LinearObjectListFrameFirst L

GenerateDemoSequence = L
End Function

Public Sub LinearObjectListFrameFirst(L As LinearObjectList)
Dim Z As Long

For Z = 1 To L.Count
    If L.Items(Z).Participate Then
        L.Current = Z
        L.Step = 1
        Exit Sub
    End If
Next

End Sub

Public Sub LinearObjectListFrameLast(L As LinearObjectList)
Dim Z As Long

For Z = L.Count To 1 Step -1
    If L.Items(Z).Participate Then
        L.Current = Z
        L.Step = L.StepCount
        Exit Sub
    End If
Next

End Sub

Public Sub LinearObjectListFramePrevious(L As LinearObjectList)
Dim Z As Long

Z = L.Current - 1
Do While Z > 0
    If L.Items(Z).Participate Then
        L.Current = Z
        L.Step = L.Step - 1
        Exit Sub
    End If
    Z = Z - 1
Loop

End Sub

Public Sub LinearObjectListFrameNext(L As LinearObjectList)
Dim Z As Long

Z = L.Current + 1
Do While Z <= L.Count
    If L.Items(Z).Participate Then
        L.Current = Z
        L.Step = L.Step + 1
        Exit Sub
    End If
    Z = Z + 1
Loop

End Sub

Public Sub AddFigureToDemoSequence(L As LinearObjectList, ByVal Index As Long)
Dim Z As Long, Q As Long
If Not IsFigure(Index) Then Exit Sub
If LinearObjectListFindItem(L, gotFigure, Index) <> -1 Then Exit Sub

Q = GetProperParentNumber(Figures(Index).FigureType)
If Q > 0 Then
    For Z = 1 To Q
        AddFigureToDemoSequence L, Figures(Index).Parents(Z - 1)
    Next Z
End If

Q = Figures(Index).NumberOfPoints
If Q > 0 Then
    For Z = 1 To Q
        AddPointToDemoSequence L, Figures(Index).Points(Z - 1)
    Next
End If

If nActiveAxesAdded Then If Index = nActiveX Or Index = nActiveY Then Exit Sub

LinearObjectListAdd L, gotFigure, Index, Figures(Index).InDemo, Figures(Index).Description

End Sub

Public Sub AddPointToDemoSequence(L As LinearObjectList, ByVal Index As Long)
If Not IsPoint(Index) Then Exit Sub
If LinearObjectListFindItem(L, gotPoint, Index) <> -1 Then Exit Sub

If BasePoint(Index).Type <> dsPoint Then
    'AddFigureToDemoSequence L, BasePoint(Index).ParentFigure
Else
    LinearObjectListAdd L, gotPoint, Index, BasePoint(Index).InDemo, BasePoint(Index).Description
End If
End Sub

Public Sub LinearObjectListAdd(L As LinearObjectList, ByVal ObjectType As GeometryObjectType, ByVal Index As Long, Optional ByVal ShouldParticipate As Boolean = True, Optional ByVal Description As String = "")
With L
    .Count = .Count + 1
    ReDim Preserve .Items(1 To .Count)
    With .Items(.Count)
        .Index = Index
        .Type = ObjectType
        .Participate = ShouldParticipate
        .Description = Description
    End With
End With
End Sub

Public Sub LinearObjectListClear(L As LinearObjectList)
With L
    .Count = 0
    .Current = 0
    ReDim .Items(1 To 1)
End With
End Sub

Public Function LinearObjectListFindItem(L As LinearObjectList, ByVal ObjectType As GeometryObjectType, ByVal Index As Long) As Long
Dim Z As Long
For Z = 1 To L.Count
    If L.Items(Z).Type = ObjectType And L.Items(Z).Index = Index Then
        LinearObjectListFindItem = Z
        Exit Function
    End If
Next
LinearObjectListFindItem = -1
End Function

Public Function GetObjectDescription(ByVal ObjectType As GeometryObjectType, ByVal Index As Long) As String
Dim S As String, Z As Long

S = "Description."

Select Case ObjectType
Case gotButton
    Select Case Buttons(Index).Type
    Case butLaunchFile
        S = GetString(ResButton) & " " & Chr(34) & GetString(ResLaunchFile) & Chr(34)
    Case butMsgBox
        S = GetString(ResButton) & " " & Chr(34) & GetString(ResMessageButton) & Chr(34)
    Case butPlaySound
        S = GetString(ResButton) & " " & Chr(34) & GetString(ResPlaySound) & Chr(34)
    Case butShowHide
        S = GetString(ResButton) & " " & Chr(34) & GetString(ResShowHideObjects) & Chr(34)
    End Select
    
Case gotFigure
    With Figures(Index)
        Select Case .FigureType
        Case dsSegment
            S = GetString(ResDes_Segment)
            S = Replace(S, "%1", .Name)
            S = Replace(S, "%2", BasePoint(.Points(0)).Name)
            S = Replace(S, "%3", BasePoint(.Points(1)).Name)
        Case dsRay
            S = GetString(ResDes_Ray)
            S = Replace(S, "%1", .Name)
            S = Replace(S, "%2", BasePoint(.Points(0)).Name)
            S = Replace(S, "%3", BasePoint(.Points(1)).Name)
        Case dsLine_2Points
            S = GetString(ResDes_Line)
            S = Replace(S, "%1", .Name)
            S = Replace(S, "%2", BasePoint(.Points(0)).Name)
            S = Replace(S, "%3", BasePoint(.Points(1)).Name)
        Case dsLine_PointAndParallelLine
            S = GetString(ResDes_LineParallel)
            S = Replace(S, "%1", .Name)
            S = Replace(S, "%3", BasePoint(.Points(0)).Name)
            S = Replace(S, "%2", Figures(.Parents(0)).Name)
        Case dsLine_PointAndPerpendicularLine
            S = GetString(ResDes_LinePerpendicular)
            S = Replace(S, "%1", .Name)
            S = Replace(S, "%3", BasePoint(.Points(0)).Name)
            S = Replace(S, "%2", Figures(.Parents(0)).Name)
        Case dsBisector
            S = GetString(ResDes_Bisector)
            S = Replace(S, "%1", .Name)
            S = Replace(S, "%4", BasePoint(.Points(0)).Name)
            S = Replace(S, "%2", BasePoint(.Points(1)).Name)
            S = Replace(S, "%3", BasePoint(.Points(2)).Name)
            
        Case dsCircle_CenterAndCircumPoint
            S = GetString(ResDes_Circle)
            S = Replace(S, "%1", .Name)
            S = Replace(S, "%2", BasePoint(.Points(0)).Name)
            S = Replace(S, "%3", BasePoint(.Points(1)).Name)
        
        Case dsCircle_CenterAndTwoPoints
            S = GetString(ResDes_CircleByRadius)
            S = Replace(S, "%1", .Name)
            S = Replace(S, "%2", BasePoint(.Points(0)).Name)
            S = Replace(S, "%3", BasePoint(.Points(1)).Name)
            S = Replace(S, "%4", BasePoint(.Points(2)).Name)
        
        Case dsCircle_ArcCenterAndRadiusAndTwoPoints
            S = GetString(ResDes_Arc)
            S = Replace(S, "%1", .Name)
            S = Replace(S, "%2", BasePoint(.Points(0)).Name)
            S = Replace(S, "%3", BasePoint(.Points(1)).Name)
            S = Replace(S, "%4", BasePoint(.Points(2)).Name)
            S = Replace(S, "%5", BasePoint(.Points(3)).Name)
            S = Replace(S, "%6", BasePoint(.Points(4)).Name)
        
        Case dsAnCircle
            S = GetString(ResFigureAnCircle) & ": " & GetCircleEquation(Index)
        
        Case dsAnLineCanonic
            S = GetString(ResFigureAnLineCanonic) & ": " & GetLineEquationText(GetLineCoordinatesAbsolute(Index))
        
        Case dsAnLineGeneral
            S = GetString(ResFigureAnLineGeneral) & ": " & GetLineEquationText(GetLineCoordinatesAbsolute(Index))
        
        Case dsAnLineNormal
            S = GetString(ResFigureAnLineNormal) & ": " & GetLineEquationText(GetLineCoordinatesAbsolute(Index))
        
        Case dsAnLineNormalPoint
            S = GetString(ResFigureAnLineNormalPoint) & ": " & GetLineEquationText(GetLineCoordinatesAbsolute(Index))
        
        Case dsAnPoint
            S = GetString(ResDes_PointCoord)
            S = Replace(S, "%1", BasePoint(.Points(0)).Name)
            S = Replace(S, "%2", .XS)
            S = Replace(S, "%3", .YS)
        
        Case dsSimmPoint
            S = GetObjectDescription(gotPoint, .Points(0))
        Case dsSimmPointByLine
            S = GetObjectDescription(gotPoint, .Points(0))
        Case dsMiddlePoint
            S = GetObjectDescription(gotPoint, .Points(0))
        Case dsPointOnFigure
            S = GetObjectDescription(gotPoint, .Points(0))
        Case dsIntersect
            S = GetObjectDescription(gotPoint, .Points(0))
        Case dsInvert
            S = GetObjectDescription(gotPoint, .Points(0))
            
        Case dsDynamicLocus
            S = GetString(ResDes_DynamicLocus)
            S = Replace(S, "%1", BasePoint(.Points(0)).Name)
            S = Replace(S, "%2", BasePoint(.Points(1)).Name)
            S = Replace(S, "%3", Figures(Figures(BasePoint(.Points(1)).ParentFigure).Parents(0)).Name)
            
        Case dsMeasureDistance
            S = GetString(ResFigureMeasureDistance) & ": (" & BasePoint(.Points(0)).Name & "," & BasePoint(.Points(1)).Name & ")"
            
        Case dsMeasureAngle
            S = GetString(ResFigureMeasureAngle) & ": <(" & BasePoint(.Points(0)).Name & "," & BasePoint(.Points(1)).Name & "," & BasePoint(.Points(2)).Name & ")"
        
        End Select
    End With
    
Case gotLabel
    S = TextLabels(Index).Caption
    If InStr(S, vbCrLf) <> 0 Then
        S = Left(S, InStr(S, vbCrLf) - 1) & "..."
    End If
    S = GetString(ResLabel) & " " & Chr(34) & S & Chr(34)
Case gotLocus
'    If Locuses(Index).Dynamic Then
'        S = GetString(ResDes_DynamicLocus)
'        S = Replace(S, "%1", BasePoint(Locuses(Index).ParentPoint).Name)
'        S = Replace(S, "%2", Locuses(Index).ParentPoint)
'    Else
        S = GetString(ResDes_LocusOfPoint)
        S = Replace(S, "%1", BasePoint(Locuses(Index).ParentPoint).Name)
'    End If
Case gotPoint
    With BasePoint(Index)
        Select Case .Type
        Case dsPoint
            S = GetString(ResDes_Point)
            S = Replace(S, "%1", .Name)
        
        Case dsPointOnFigure
            S = GetString(ResDes_PointOnFigure)
            S = Replace(S, "%1", .Name)
            S = Replace(S, "%2", Figures(Figures(.ParentFigure).Parents(0)).Name)
            
        Case dsMiddlePoint
            S = GetString(ResDes_Middlepoint)
            S = Replace(S, "%1", .Name)
            S = Replace(S, "%2", BasePoint(Figures(.ParentFigure).Points(1)).Name)
            S = Replace(S, "%3", BasePoint(Figures(.ParentFigure).Points(2)).Name)
            
        Case dsSimmPoint
            S = GetString(ResDes_SimmPoint)
            S = Replace(S, "%1", .Name)
            S = Replace(S, "%2", BasePoint(Figures(.ParentFigure).Points(1)).Name)
            S = Replace(S, "%3", BasePoint(Figures(.ParentFigure).Points(2)).Name)
        
        Case dsSimmPointByLine
            S = GetString(ResDes_SimmPointByLine)
            S = Replace(S, "%1", .Name)
            S = Replace(S, "%2", BasePoint(Figures(.ParentFigure).Points(1)).Name)
            S = Replace(S, "%3", Figures(Figures(.ParentFigure).Parents(0)).Name)
        
        Case dsIntersect
            S = GetString(ResDes_Intersect)
            S = Replace(S, "%1", .Name)
            S = Replace(S, "%2", Figures(Figures(.ParentFigure).Parents(0)).Name)
            S = Replace(S, "%3", Figures(Figures(.ParentFigure).Parents(1)).Name)
            
        Case dsInvert
            S = GetString(ResDes_Invert)
            S = Replace(S, "%1", .Name)
            S = Replace(S, "%2", BasePoint(Figures(.ParentFigure).Points(1)).Name)
            S = Replace(S, "%3", Figures(Figures(.ParentFigure).Parents(0)).Name)
        
        Case dsCircle_ArcCenterAndRadiusAndTwoPoints
            S = GetString(ResDes_ArcEndPoint)
            S = Replace(S, "%1", .Name)
            S = Replace(S, "%2", Figures(.ParentFigure).Name)
        End Select
    End With

Case gotSG
    Select Case StaticGraphics(Index).Type
    Case sgBezier
        S = GetString(ResDes_Bezier)
        S = Replace(S, "%1", BasePoint(StaticGraphics(Index).Points(1)).Name)
        S = Replace(S, "%2", BasePoint(StaticGraphics(Index).Points(2)).Name)
        S = Replace(S, "%3", BasePoint(StaticGraphics(Index).Points(3)).Name)
        S = Replace(S, "%4", BasePoint(StaticGraphics(Index).Points(4)).Name)
        
    Case sgPolygon
        With StaticGraphics(Index)
            S = ""
            For Z = 1 To .NumberOfPoints
                S = S & BasePoint(.Points(Z)).Name & ","
            Next
            S = Trim(S)
            S = Left(S, Len(S) - 1)
            S = Replace(GetString(ResDes_Polygon), "%1", S)
        End With
        
    Case sgVector
        S = GetString(ResDes_Vector)
        S = Replace(S, "%1", BasePoint(StaticGraphics(Index).Points(1)).Name)
        S = Replace(S, "%2", BasePoint(StaticGraphics(Index).Points(2)).Name)
    End Select
End Select

GetObjectDescription = S
End Function

Public Sub RefillObjectDescriptions(Optional ByVal ShouldIgnoreUnempty As Boolean = True)
Dim Z As Long

If ShouldIgnoreUnempty Then
    For Z = 1 To PointCount
        If BasePoint(Z).Description = "" Then BasePoint(Z).Description = GetObjectDescription(gotPoint, Z)
    Next
    
    For Z = 0 To FigureCount - 1
        If Figures(Z).Description = "" Then Figures(Z).Description = GetObjectDescription(gotFigure, Z)
    Next
    
    For Z = 1 To StaticGraphicCount
        If StaticGraphics(Z).Description = "" Then StaticGraphics(Z).Description = GetObjectDescription(gotSG, Z)
    Next
    
    For Z = 1 To LabelCount
        If TextLabels(Z).Description = "" Then TextLabels(Z).Description = GetObjectDescription(gotLabel, Z)
    Next
    
    For Z = 1 To ButtonCount
        If Buttons(Z).Description = "" Then Buttons(Z).Description = GetObjectDescription(gotButton, Z)
    Next
Else
    For Z = 1 To PointCount
        BasePoint(Z).Description = GetObjectDescription(gotPoint, Z)
    Next
    
    For Z = 0 To FigureCount - 1
        Figures(Z).Description = GetObjectDescription(gotFigure, Z)
    Next
    
    For Z = 1 To StaticGraphicCount
        StaticGraphics(Z).Description = GetObjectDescription(gotSG, Z)
    Next
    
    For Z = 1 To LabelCount
        TextLabels(Z).Description = GetObjectDescription(gotLabel, Z)
    Next
    
    For Z = 1 To ButtonCount
        Buttons(Z).Description = GetObjectDescription(gotButton, Z)
    Next
End If
End Sub

Public Function GetObjectName(ByVal ObjectType As GeometryObjectType, ByVal Index As Long) As String
Dim S As String
S = ""

Select Case ObjectType
Case gotButton
    S = GetString(ResButton) & Index
Case gotFigure
    S = Figures(Index).Name
Case gotLabel
    S = GetString(ResLabel) & Index
Case gotLocus
    S = GetString(ResLocus) & Index
Case gotPoint
    S = BasePoint(Index).Name
Case gotSG
    S = GetString(ResStaticObjectBase + 2 * StaticGraphics(Index).Type) & Index
End Select

GetObjectName = S
End Function

Public Function GetGeometryObjectType(ByVal T As DrawState) As GeometryObjectType
Select Case T
Case dsPoint, dsPointOnFigure, dsMiddlePoint, dsSimmPoint, dsSimmPointByLine, dsIntersect, dsInvert
    GetGeometryObjectType = gotPoint
Case dsCircle_ArcCenterAndRadiusAndTwoPoints, dsCircle_CenterAndCircumPoint, dsCircle_CenterAndTwoPoints, _
                dsSegment To dsLine_PointAndPerpendicularLine, dsBisector
    GetGeometryObjectType = gotFigure
End Select
End Function

Public Sub UpdateObjectsWithLinearListData(L As LinearObjectList)
Dim Z As Long
For Z = 1 To L.Count
    With L.Items(Z)
        Select Case .Type
        Case gotButton
            Buttons(.Index).Description = .Description
            Buttons(.Index).InDemo = .Participate
        Case gotFigure
            Figures(.Index).Description = .Description
            Figures(.Index).InDemo = .Participate
        Case gotLabel
            TextLabels(.Index).Description = .Description
            TextLabels(.Index).InDemo = .Participate
        Case gotLocus
            Locuses(.Index).Description = .Description
            Locuses(.Index).InDemo = .Participate
        Case gotPoint
            BasePoint(.Index).Description = .Description
            BasePoint(.Index).InDemo = .Participate
        Case gotSG
            StaticGraphics(.Index).Description = .Description
            StaticGraphics(.Index).InDemo = .Participate
        End Select
    End With
Next
End Sub

Public Sub DrawSituation()
DrawLinearObjectList Paper.hDC, DemoSequence, , , False
If FormMain.tmrDemo.Enabled Then DrawCameraIcon
Paper.Refresh
DemoShowDescription
End Sub

Public Sub DrawCameraIcon()
Const Margin = 8
Dim C As IPictureDisp
Static i As Integer
i = 1 - i

Set C = LoadResPicture(ResBMPCamera + i, vbResBitmap)
TransparentBlt Paper.hDC, C.handle, PaperScaleWidth - Margin - Paper.ScaleX(C.Width, vbHimetric, vbPixels), Margin, vbWhite
End Sub
