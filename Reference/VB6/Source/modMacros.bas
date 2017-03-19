Attribute VB_Name = "modMacros"
'======================================================
' Module for work with macroses
' Implements mostly creation process
'======================================================
Option Explicit

'======================================================
'======================================================
' Entire Macro structure manipulation:
' Add, remove, organize
'======================================================
'======================================================

Public Sub AddMacro(tempMacro As Macro)
MacroCount = MacroCount + 1
ReDim Preserve Macros(1 To MacroCount)
Macros(MacroCount) = tempMacro

i_AddMacroRunMenu tempMacro.Name
End Sub

'======================================================
' Unload an already loaded macro
'======================================================

Public Sub RemoveMacro(ByVal Index As Long)
Dim Z As Long

If Index < 1 Or Index > MacroCount Then Exit Sub

i_RemoveMacroRunMenu Index

If Index < MacroCount Then
    For Z = Index To MacroCount - 1
        Macros(Z) = Macros(Z + 1)
    Next Z
End If

MacroCount = MacroCount - 1
If MacroCount > 0 Then ReDim Preserve Macros(1 To MacroCount) Else ReDim Macros(1 To 1)
End Sub

'======================================================
' Add a new virtual element to the macro's givens list
'======================================================

Public Function AddMacroGiven(tempMacro As Macro, ByVal GivenType As DrawState) As Long
Dim S As String

tempMacro.GivenCount = tempMacro.GivenCount + 1
ReDim Preserve tempMacro.Givens(1 To tempMacro.GivenCount)
tempMacro.Givens(tempMacro.GivenCount).Type = GivenType

If IsLineType(GivenType) Then
    S = GetString(ResGivenHintSegmentRayLine)
ElseIf IsCircleType(GivenType) Then
    S = GetString(ResGivenHintCircle)
Else
    S = GetString(ResGivenHintPoint)
End If
tempMacro.Givens(tempMacro.GivenCount).Description = S

AddMacroGiven = tempMacro.GivenCount
End Function

Public Function DeleteMacroGiven(tempMacro As Macro, ByVal Index As Long) As Boolean
' Delete a single given from a macro list
'======================================================
Dim Z As Long

For Z = 1 To PointCount
    ' GivenPoints(Z) indicates the index (order) of pointZ in givens list;
    ' 0 if PointZ doesn't belong to the givenslist
    If GivenPoints(Z) = Index Then GivenPoints(Z) = 0
    If GivenPoints(Z) > Index Then GivenPoints(Z) = GivenPoints(Z) - 1
Next

For Z = 0 To FigureCount - 1
    If GivenFigures(Z) = Index Then GivenFigures(Z) = 0
    If GivenFigures(Z) > Index Then GivenFigures(Z) = GivenFigures(Z) - 1
Next

If Index < tempMacro.GivenCount Then
    For Z = Index To tempMacro.GivenCount - 1
        tempMacro.Givens(Z) = tempMacro.Givens(Z + 1)
    Next
End If

tempMacro.GivenCount = tempMacro.GivenCount - 1
If tempMacro.GivenCount > 0 Then ReDim Preserve tempMacro.Givens(1 To tempMacro.GivenCount)
End Function

Public Sub MacroResetGivens()
' Erase already selected Givens
'======================================================
InitZeroTags
tempMacro.GivenCount = 0
ReDim tempMacro.Givens(1 To 1)
ReDim GivenPoints(1 To PointCount)
If FigureCount > 0 Then ReDim GivenFigures(0 To FigureCount - 1) Else ReDim GivenFigures(0 To 0)
End Sub

Public Sub MacroResetResults()
' Erase already selected Results
'======================================================
InitZeroTags
tempMacro.ResultCount = 0
tempMacro.FigurePointCount = 0
tempMacro.SGCount = 0
ReDim tempMacro.Results(1 To 1)
ReDim tempMacro.FigurePoints(1 To 1)
ReDim tempMacro.SG(1 To 1)
End Sub

Public Sub MacroFillResults()
' Macro Results are either figures or SGs.
' Generates macro result list; having info about
' ResultPoints(), ResultFigures(), ResultSGs(), ResultLoci()
'
Dim Z As Long
Dim FiguresToAdd As New Collection
Dim SGsToAdd As New Collection

MacroResetResults

For Z = 0 To FigureCount - 1
    If ResultFigures(Z) > 0 Then
        FiguresToAdd.Add Z
        'ResultFigures(Z) = 0
    End If
Next

For Z = 1 To LocusCount
    If ResultLoci(Z) > 0 Then
'        AP = Locuses(AL).ParentPoint
'        If IsPoint(AP) Then
'            If BasePoint(AP).Type <> dsPoint And Figures(BasePoint(AP).ParentFigure).Tag = 0 And GivenPoints(AP) = 0 Then
'                Figures(BasePoint(AP).ParentFigure).Tag = AddMacroResult(tempMacro, Figures(BasePoint(AP).ParentFigure))

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'   WARNING!
'   Actually adding a figurepoint that describes the locus,
'   and not the locus itself. Locus itself will be added some next time,
'   in AddMacroResult, when it recognizes this point as a locus
'   describing point
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        
        'FiguresToAdd.Add BasePoint(Locuses(Z).ParentPoint).ParentFigure
        If Locuses(Z).ParentFigure <> GetLocusParentFigure(Locuses(Z).ParentPoint) Then Locuses(Z).ParentFigure = GetLocusParentFigure(Locuses(Z).ParentPoint)
        FiguresToAdd.Add Locuses(Z).ParentFigure
        
    End If
Next

For Z = 1 To StaticGraphicCount
    If ResultSGs(Z) > 0 Then
        SGsToAdd.Add Z
        'ResultSGs(Z) = 0
    End If
Next

'#############################################

For Z = 1 To FiguresToAdd.Count
    Figures(FiguresToAdd(Z)).Tag = AddMacroResult(tempMacro, Figures(FiguresToAdd(Z)), False)
    If Figures(FiguresToAdd(Z)).Tag <= 0 Then
        ' will process macro errors later
    End If
Next

For Z = 1 To SGsToAdd.Count
    StaticGraphics(SGsToAdd(Z)).Tag = AddMacroSG(tempMacro, SGsToAdd(Z))
    If StaticGraphics(SGsToAdd(Z)).Tag <= 0 Then
        ' will process macro errors later
    End If
Next

InitZeroTags
End Sub

Public Sub MacroObjectSelected(ByVal Index As Long)
'======================================================
' Happens when user selects an ambiguity popup menu item
' in SelectGivens or SelectResults mode.
'======================================================
Dim Z As Long, PF As Long

If DragS.State = dscMacroStateGivens Then
    If Index < 0 Then
        If GivenPoints(-Index) = 0 Then
            GivenPoints(-Index) = AddMacroGiven(tempMacro, dsPoint)
        Else
            DeleteMacroGiven tempMacro, GivenPoints(-Index)
        End If
        PaperCls
        ShowAllWithGivens
    Else
        If GivenFigures(Index) = 0 Then
            GivenFigures(Index) = AddMacroGiven(tempMacro, Figures(Index).FigureType)
        Else
            DeleteMacroGiven tempMacro, GivenFigures(Index)
        End If
        PaperCls
        ShowAllWithGivens
    End If
End If

'======================================================
' F->ParentPoint[z]->ParentFigure;
If DragS.State = dscMacroStateResults Then
    If Index >= 0 Then
'        Figures(Index).Tag = AddMacroResult(tempMacro, Figures(Index))
'        If Figures(Index).Tag <= 0 Then i_CancelMacro GetString(ResMacroErrBase - 2 * Figures(Index).Tag): Exit Sub
        ResultFigures(Index) = 1 - ResultFigures(Index)
        If ResultFigures(Index) = 1 And IsVisual(Index) Then
            For Z = 0 To Figures(Index).NumberOfPoints - 1
                If Not IsChildPointPos(Figures(Index), Z) Then
                    If BasePoint(Figures(Index).Points(Z)).Type <> dsPoint Then
                        PF = BasePoint(Figures(Index).Points(Z)).ParentFigure
                        If CanBeAResult(PF) And ResultFigures(PF) = 0 Then MacroObjectSelected PF
                    End If
                End If
            Next
        End If
    Else
        Index = -Index
'        ResultSGs(Index) = AddMacroSG(tempMacro, Index)
        ResultSGs(Index) = 1 - ResultSGs(Index)
    End If
    PaperCls
    ShowAllWithResults
End If
End Sub

Public Function AddMacroResult(tempMacro As Macro, Result As Figure, Optional ByVal ShouldHide As Boolean = True) As Long
' Adds a specified result to the tempMacro's result list.
' Recursive.
'======================================================
Dim tempFigure As Figure, TP As Long, TPParent As Long, hResult As Long, Z As Long

If Result.Tag <> 0 Then
    Result.Hide = ShouldHide
    AddMacroResult = Result.Tag
    Exit Function
End If
tempFigure = Result
tempFigure.ZOrder = 0

'#######################################

For Z = 0 To tempFigure.NumberOfPoints - 1
    If Not IsChildPointPos(tempFigure, Z) Then
        TP = tempFigure.Points(Z)
        TPParent = BasePoint(TP).ParentFigure
        If GivenPoints(TP) <> 0 Then
            If tempFigure.FigureType = dsAnPoint Then
                ReplacePointInTree tempFigure.XTree, tempFigure.Points(Z), -GivenPoints(TP)
                ReplacePointInTree tempFigure.YTree, tempFigure.Points(Z), -GivenPoints(TP)
            End If
            tempFigure.Points(Z) = -GivenPoints(TP)
        Else
            If BasePoint(TP).Type = dsPoint Then
                AddMacroResult = -meNotAllGivensSelected
                Exit Function
            Else
                If BasePoint(TP).Tag = 0 Then
                    hResult = AddMacroResult(tempMacro, Figures(TPParent))
                    If hResult <= 0 Then AddMacroResult = hResult: Exit Function
                    Figures(TPParent).Tag = hResult
                End If
                If tempFigure.FigureType = dsAnPoint Then
                    ReplacePointInTree tempFigure.XTree, tempFigure.Points(Z), -BasePoint(TP).Tag
                    ReplacePointInTree tempFigure.YTree, tempFigure.Points(Z), -BasePoint(TP).Tag
                End If
                tempFigure.Points(Z) = -BasePoint(TP).Tag
            End If
        End If
    End If
    If tempFigure.FigureType = dsAnPoint Then
        tempFigure.XS = RestoreExpressionFromTree(tempFigure.XTree, True)
        tempFigure.YS = RestoreExpressionFromTree(tempFigure.YTree, True)
    End If
Next Z

'#############################################

For Z = 0 To GetProperParentNumber(tempFigure.FigureType) - 1
    If GivenFigures(tempFigure.Parents(Z)) <> 0 Then
        tempFigure.Parents(Z) = -GivenFigures(tempFigure.Parents(Z))
    Else
        If Figures(tempFigure.Parents(Z)).Tag = 0 Then Figures(tempFigure.Parents(Z)).Tag = AddMacroResult(tempMacro, Figures(tempFigure.Parents(Z)))
        If Figures(tempFigure.Parents(Z)).Tag < 0 Then AddMacroResult = Figures(tempFigure.Parents(Z)).Tag: Exit Function
        tempFigure.Parents(Z) = Figures(tempFigure.Parents(Z)).Tag
    End If
Next

'#############################################
        
For Z = 0 To tempFigure.NumberOfPoints - 1
    If IsChildPointPos(tempFigure, Z) Then
        If BasePoint(tempFigure.Points(Z)).Tag = 0 And GivenPoints(tempFigure.Points(Z)) = 0 Then
            tempMacro.FigurePointCount = tempMacro.FigurePointCount + 1
            ReDim Preserve tempMacro.FigurePoints(1 To tempMacro.FigurePointCount)
            tempMacro.FigurePoints(tempMacro.FigurePointCount) = BasePoint(tempFigure.Points(Z))
            
            If tempFigure.FigureType <> dsDynamicLocus Then
                If tempFigure.FigureType = dsIntersect Then
                    ' backing up Hide into Tag
                    tempMacro.FigurePoints(tempMacro.FigurePointCount).Tag = tempMacro.FigurePoints(tempMacro.FigurePointCount).Hide
                    tempMacro.FigurePoints(tempMacro.FigurePointCount).Hide = tempMacro.FigurePoints(tempMacro.FigurePointCount).Tag Or ShouldHide
                    tempMacro.FigurePoints(tempMacro.FigurePointCount).Enabled = Not tempMacro.FigurePoints(tempMacro.FigurePointCount).Hide
                Else
                    tempMacro.FigurePoints(tempMacro.FigurePointCount).Hide = ShouldHide
                    tempMacro.FigurePoints(tempMacro.FigurePointCount).Enabled = Not ShouldHide
                End If
            End If
            
            BasePoint(tempFigure.Points(Z)).Tag = -tempMacro.FigurePointCount
            tempMacro.FigurePoints(tempMacro.FigurePointCount).ParentFigure = tempMacro.ResultCount + 1
            tempFigure.Points(Z) = tempMacro.FigurePointCount
        Else
            AddMacroResult = -meNotAllGivensSelected
            Exit Function
        End If
    End If
Next
        
'#############################################

tempMacro.ResultCount = tempMacro.ResultCount + 1
ReDim Preserve tempMacro.Results(1 To tempMacro.ResultCount)
tempFigure.Hide = ShouldHide
tempMacro.Results(tempMacro.ResultCount) = tempFigure
AddMacroResult = tempMacro.ResultCount

' Now: if we are given a locus descriptor point, add the dynamic locus itself
'For Z = 0 To FigureCount - 1
'    If Figures(Z).FigureType = dsDynamicLocus Then
'        If BasePoint(Figures(Z).Points(0)).Tag <> 0 And BasePoint(Figures(Z).Points(1)).Tag <> 0 And Figures(Z).Tag = 0 And Figures(Z).Name <> Result.Name Then
'            hResult = AddMacroResult(tempMacro, Figures(Z))
'            If hResult <= 0 Then AddMacroResult = hResult: Exit Function
'            Figures(Z).Tag = hResult
'        End If
'    End If
'Next
End Function

'
'Public Function AddMacroResult(tempMacro As Macro, Result As Figure) As Long
'' Adds a specified result to the tempMacro's result list.
'' recursive;
''======================================================
'Dim tempFigure As Figure, TP As Long, TPParent As Long, hResult As Long, Z As Long
'
'If Result.Tag <> 0 Then AddMacroResult = Result.Tag: Exit Function
'tempFigure = Result
'tempFigure.ZOrder = 0
'
''#######################################
'Select Case tempFigure.FigureType
'    Case dsSegment, dsRay, dsLine_2Points, dsBisector, dsCircle_CenterAndCircumPoint, dsCircle_CenterAndTwoPoints, dsDynamicLocus
'        For Z = 0 To tempFigure.NumberOfPoints - 1
'            TP = tempFigure.Points(Z)
'            TPParent = BasePoint(TP).ParentFigure
'            If GivenPoints(TP) <> 0 Then
'                tempFigure.Points(Z) = -GivenPoints(TP)
'            Else
'                If BasePoint(TP).Type = dsPoint Then
'                    AddMacroResult = -meNotAllGivensSelected
'                    Exit Function
'                Else
'                    If BasePoint(TP).Tag = 0 Then
'                        hResult = AddMacroResult(tempMacro, Figures(TPParent))
'                        If hResult <= 0 Then AddMacroResult = hResult: Exit Function
'                        Figures(TPParent).Tag = hResult
'                    End If
'                    tempFigure.Points(Z) = -BasePoint(TP).Tag
'                End If
'            End If
'        Next Z
'    '#############################################3
'    Case dsLine_PointAndParallelLine, dsLine_PointAndPerpendicularLine
'        TP = tempFigure.Points(0)
'        TPParent = BasePoint(TP).ParentFigure
'        If GivenPoints(TP) <> 0 Then
'            tempFigure.Points(0) = -GivenPoints(TP)
'        Else
'            If BasePoint(TP).Type = dsPoint Then
'                AddMacroResult = -meNotAllGivensSelected
'                Exit Function
'            Else
'                If BasePoint(TP).Tag = 0 Then
'                    hResult = AddMacroResult(tempMacro, Figures(TPParent))
'                    If hResult <= 0 Then AddMacroResult = hResult: Exit Function
'                    Figures(TPParent).Tag = hResult
'                End If
'                tempFigure.Points(0) = -BasePoint(TP).Tag
'            End If
'        End If
'
'        If GivenFigures(tempFigure.Parents(0)) <> 0 Then
'            tempFigure.Parents(0) = -GivenFigures(tempFigure.Parents(0))
'        Else
'            If Figures(tempFigure.Parents(0)).Tag = 0 Then Figures(tempFigure.Parents(0)).Tag = AddMacroResult(tempMacro, Figures(tempFigure.Parents(0)))
'            If Figures(tempFigure.Parents(0)).Tag < 0 Then AddMacroResult = Figures(tempFigure.Parents(0)).Tag: Exit Function
'            tempFigure.Parents(0) = Figures(tempFigure.Parents(0)).Tag
'        End If
'
'    '#######################################
'    Case dsCircle_ArcCenterAndRadiusAndTwoPoints
'        For Z = 0 To 4
'            TP = tempFigure.Points(Z)
'            TPParent = BasePoint(TP).ParentFigure
'            If GivenPoints(TP) <> 0 Then
'                tempFigure.Points(Z) = -GivenPoints(TP)
'            Else
'                If BasePoint(TP).Type = dsPoint Then
'                    AddMacroResult = -meNotAllGivensSelected
'                    Exit Function
'                Else
'                    If BasePoint(TP).Tag = 0 Then
'                        hResult = AddMacroResult(tempMacro, Figures(TPParent))
'                        If hResult <= 0 Then AddMacroResult = hResult: Exit Function
'                        Figures(TPParent).Tag = hResult
'                    End If
'                    tempFigure.Points(Z) = -BasePoint(TP).Tag
'                End If
'            End If
'        Next Z
'
'        If BasePoint(tempFigure.Points(5)).Tag = 0 Then
'            tempMacro.FigurePointCount = tempMacro.FigurePointCount + 1
'            ReDim Preserve tempMacro.FigurePoints(1 To tempMacro.FigurePointCount)
'            tempMacro.FigurePoints(tempMacro.FigurePointCount) = BasePoint(tempFigure.Points(5))
'            BasePoint(tempFigure.Points(5)).Tag = -tempMacro.FigurePointCount
'            tempMacro.FigurePoints(tempMacro.FigurePointCount).ParentFigure = tempMacro.ResultCount + 1
'            tempFigure.Points(5) = tempMacro.FigurePointCount
'        End If
'        If BasePoint(tempFigure.Points(6)).Tag = 0 Then
'            tempMacro.FigurePointCount = tempMacro.FigurePointCount + 1
'            ReDim Preserve tempMacro.FigurePoints(1 To tempMacro.FigurePointCount)
'            tempMacro.FigurePoints(tempMacro.FigurePointCount) = BasePoint(tempFigure.Points(6))
'            BasePoint(tempFigure.Points(6)).Tag = -tempMacro.FigurePointCount
'            tempMacro.FigurePoints(tempMacro.FigurePointCount).ParentFigure = tempMacro.ResultCount + 1
'            tempFigure.Points(6) = tempMacro.FigurePointCount
'        End If
'
'    '###############################3
'    Case dsMiddlePoint, dsSimmPoint
'        For Z = 1 To 2
'            TP = tempFigure.Points(Z)
'            TPParent = BasePoint(TP).ParentFigure
'            If GivenPoints(TP) <> 0 Then
'                tempFigure.Points(Z) = -GivenPoints(TP)
'            Else
'                If BasePoint(TP).Type = dsPoint Then
'                    AddMacroResult = -meNotAllGivensSelected
'                    Exit Function
'                Else
'                    If BasePoint(TP).Tag = 0 Then
'                        hResult = AddMacroResult(tempMacro, Figures(TPParent))
'                        If hResult <= 0 Then AddMacroResult = hResult: Exit Function
'                        Figures(TPParent).Tag = hResult
'                    End If
'                    tempFigure.Points(Z) = -BasePoint(TP).Tag
'                End If
'            End If
'        Next Z
'
'        If BasePoint(tempFigure.Points(0)).Tag = 0 Then
'            tempMacro.FigurePointCount = tempMacro.FigurePointCount + 1
'            ReDim Preserve tempMacro.FigurePoints(1 To tempMacro.FigurePointCount)
'            tempMacro.FigurePoints(tempMacro.FigurePointCount) = BasePoint(tempFigure.Points(0))
'            BasePoint(tempFigure.Points(0)).Tag = -tempMacro.FigurePointCount
'            tempMacro.FigurePoints(tempMacro.FigurePointCount).ParentFigure = tempMacro.ResultCount + 1
'            tempFigure.Points(0) = tempMacro.FigurePointCount
'        End If
'
'    '###############################>8
'    Case dsAnPoint
'        For Z = 1 To tempFigure.NumberOfPoints - 1
'            TP = tempFigure.Points(Z)
'            TPParent = BasePoint(TP).ParentFigure
'            If GivenPoints(TP) <> 0 Then
'                ReplacePointInTree tempFigure.XTree, tempFigure.Points(Z), -GivenPoints(TP)
'                ReplacePointInTree tempFigure.YTree, tempFigure.Points(Z), -GivenPoints(TP)
'                tempFigure.Points(Z) = -GivenPoints(TP)
'            Else
'                If BasePoint(TP).Type = dsPoint Then
'                    AddMacroResult = -meNotAllGivensSelected
'                    Exit Function
'                Else
'                    If BasePoint(TP).Tag = 0 Then
'                        hResult = AddMacroResult(tempMacro, Figures(TPParent))
'                        If hResult <= 0 Then AddMacroResult = hResult: Exit Function
'                        Figures(TPParent).Tag = hResult
'                    End If
'                    ReplacePointInTree tempFigure.XTree, tempFigure.Points(Z), -BasePoint(TP).Tag
'                    ReplacePointInTree tempFigure.YTree, tempFigure.Points(Z), -BasePoint(TP).Tag
'                    tempFigure.Points(Z) = -BasePoint(TP).Tag
'                End If
'            End If
'            tempFigure.XS = RestoreExpressionFromTree(tempFigure.XTree, True)
'            tempFigure.YS = RestoreExpressionFromTree(tempFigure.YTree, True)
'        Next Z
'
'        If BasePoint(tempFigure.Points(0)).Tag = 0 Then
'            tempMacro.FigurePointCount = tempMacro.FigurePointCount + 1
'            ReDim Preserve tempMacro.FigurePoints(1 To tempMacro.FigurePointCount)
'            tempMacro.FigurePoints(tempMacro.FigurePointCount) = BasePoint(tempFigure.Points(0))
'            BasePoint(tempFigure.Points(0)).Tag = -tempMacro.FigurePointCount
'            tempMacro.FigurePoints(tempMacro.FigurePointCount).ParentFigure = tempMacro.ResultCount + 1
'            tempFigure.Points(0) = tempMacro.FigurePointCount
'        End If
'
'    '##############################
'    Case dsSimmPointByLine, dsInvert
'        TP = tempFigure.Points(1)
'        TPParent = BasePoint(TP).ParentFigure
'        If GivenPoints(TP) <> 0 Then
'            tempFigure.Points(1) = -GivenPoints(TP)
'        Else
'            If BasePoint(TP).Type = dsPoint Then
'                AddMacroResult = -meNotAllGivensSelected
'                Exit Function
'            Else
'                If BasePoint(TP).Tag = 0 Then
'                    hResult = AddMacroResult(tempMacro, Figures(TPParent))
'                    If hResult <= 0 Then AddMacroResult = hResult: Exit Function
'                    Figures(TPParent).Tag = hResult
'                End If
'                tempFigure.Points(1) = -BasePoint(TP).Tag
'            End If
'        End If
'
'        If GivenFigures(tempFigure.Parents(0)) <> 0 Then
'            tempFigure.Parents(0) = -GivenFigures(tempFigure.Parents(0))
'        Else
'            If Figures(tempFigure.Parents(0)).Tag = 0 Then Figures(tempFigure.Parents(0)).Tag = AddMacroResult(tempMacro, Figures(tempFigure.Parents(0)))
'            If Figures(tempFigure.Parents(0)).Tag < 0 Then AddMacroResult = Figures(tempFigure.Parents(0)).Tag: Exit Function
'            tempFigure.Parents(0) = Figures(tempFigure.Parents(0)).Tag
'        End If
'
'        If BasePoint(tempFigure.Points(0)).Tag = 0 Then
'            tempMacro.FigurePointCount = tempMacro.FigurePointCount + 1
'            ReDim Preserve tempMacro.FigurePoints(1 To tempMacro.FigurePointCount)
'            tempMacro.FigurePoints(tempMacro.FigurePointCount) = BasePoint(tempFigure.Points(0))
'            BasePoint(tempFigure.Points(0)).Tag = -tempMacro.FigurePointCount
'            tempMacro.FigurePoints(tempMacro.FigurePointCount).ParentFigure = tempMacro.ResultCount + 1
'            tempFigure.Points(0) = tempMacro.FigurePointCount
'        End If
'
'    '############################3
'    Case dsIntersect
'        If GivenFigures(tempFigure.Parents(0)) <> 0 Then
'            tempFigure.Parents(0) = -GivenFigures(tempFigure.Parents(0))
'        Else
'            If Figures(tempFigure.Parents(0)).Tag = 0 Then Figures(tempFigure.Parents(0)).Tag = AddMacroResult(tempMacro, Figures(tempFigure.Parents(0)))
'            If Figures(tempFigure.Parents(0)).Tag < 0 Then AddMacroResult = Figures(tempFigure.Parents(0)).Tag: Exit Function
'            tempFigure.Parents(0) = Figures(tempFigure.Parents(0)).Tag
'        End If
'        If GivenFigures(tempFigure.Parents(1)) <> 0 Then
'            tempFigure.Parents(1) = -GivenFigures(tempFigure.Parents(1))
'        Else
'            If Figures(tempFigure.Parents(1)).Tag = 0 Then Figures(tempFigure.Parents(1)).Tag = AddMacroResult(tempMacro, Figures(tempFigure.Parents(1)))
'            If Figures(tempFigure.Parents(1)).Tag < 0 Then AddMacroResult = Figures(tempFigure.Parents(1)).Tag: Exit Function
'            tempFigure.Parents(1) = Figures(tempFigure.Parents(1)).Tag
'        End If
'
'        If BasePoint(tempFigure.Points(0)).Tag = 0 And GivenPoints(tempFigure.Points(0)) = 0 Then
'            tempMacro.FigurePointCount = tempMacro.FigurePointCount + 1
'            ReDim Preserve tempMacro.FigurePoints(1 To tempMacro.FigurePointCount)
'            tempMacro.FigurePoints(tempMacro.FigurePointCount) = BasePoint(tempFigure.Points(0))
'            BasePoint(tempFigure.Points(0)).Tag = -tempMacro.FigurePointCount
'            tempMacro.FigurePoints(tempMacro.FigurePointCount).ParentFigure = tempMacro.ResultCount + 1
'            tempFigure.Points(0) = tempMacro.FigurePointCount
'        Else
'            AddMacroResult = -meNotAllGivensSelected
'            Exit Function
'        End If
'
'        If BasePoint(tempFigure.Points(1)).Tag = 0 And GivenPoints(tempFigure.Points(1)) = 0 Then
'            tempMacro.FigurePointCount = tempMacro.FigurePointCount + 1
'            ReDim Preserve tempMacro.FigurePoints(1 To tempMacro.FigurePointCount)
'            tempMacro.FigurePoints(tempMacro.FigurePointCount) = BasePoint(tempFigure.Points(1))
'            BasePoint(tempFigure.Points(1)).Tag = -tempMacro.FigurePointCount
'            tempMacro.FigurePoints(tempMacro.FigurePointCount).ParentFigure = tempMacro.ResultCount + 1
'            tempFigure.Points(1) = tempMacro.FigurePointCount
'        Else
'            AddMacroResult = -meNotAllGivensSelected
'            Exit Function
'        End If
'
'    '########################
'    Case dsPointOnFigure
'        If GivenFigures(tempFigure.Parents(0)) <> 0 Then
'            tempFigure.Parents(0) = -GivenFigures(tempFigure.Parents(0))
'        Else
'            If Figures(tempFigure.Parents(0)).Tag = 0 Then Figures(tempFigure.Parents(0)).Tag = AddMacroResult(tempMacro, Figures(tempFigure.Parents(0)))
'            If Figures(tempFigure.Parents(0)).Tag < 0 Then AddMacroResult = Figures(tempFigure.Parents(0)).Tag: Exit Function
'            tempFigure.Parents(0) = Figures(tempFigure.Parents(0)).Tag
'        End If
'
'        If BasePoint(tempFigure.Points(0)).Tag = 0 Then
'            tempMacro.FigurePointCount = tempMacro.FigurePointCount + 1
'            ReDim Preserve tempMacro.FigurePoints(1 To tempMacro.FigurePointCount)
'            tempMacro.FigurePoints(tempMacro.FigurePointCount) = BasePoint(tempFigure.Points(0))
'            BasePoint(tempFigure.Points(0)).Tag = -tempMacro.FigurePointCount
'            tempMacro.FigurePoints(tempMacro.FigurePointCount).ParentFigure = tempMacro.ResultCount + 1
'            tempFigure.Points(0) = tempMacro.FigurePointCount
'        End If
'End Select
''############################################
'
'tempMacro.ResultCount = tempMacro.ResultCount + 1
'ReDim Preserve tempMacro.Results(1 To tempMacro.ResultCount)
'tempMacro.Results(tempMacro.ResultCount) = tempFigure
'AddMacroResult = tempMacro.ResultCount
'
'For Z = 0 To FigureCount - 1
'    If Figures(Z).FigureType = dsDynamicLocus Then
'        If BasePoint(Figures(Z).Points(0)).Tag <> 0 And BasePoint(Figures(Z).Points(1)).Tag <> 0 And Figures(Z).Tag = 0 And Figures(Z).Name <> Result.Name Then
'            hResult = AddMacroResult(tempMacro, Figures(Z))
'            If hResult <= 0 Then AddMacroResult = hResult: Exit Function
'            Figures(Z).Tag = hResult
'        End If
'    End If
'Next
'End Function

Public Function AddMacroSG(tempMacro As Macro, ByVal Index As Long) As Long
Dim TPParent As Long, TP As Long, Z As Long, hResult As Long

With tempMacro
    .SGCount = .SGCount + 1
    ReDim Preserve .SG(1 To .SGCount)
    .SG(.SGCount) = StaticGraphics(Index)
   
    With .SG(.SGCount)
        For Z = 1 To .NumberOfPoints
            TP = .Points(Z)
            TPParent = BasePoint(TP).ParentFigure

            If GivenPoints(TP) <> 0 Then
                .Points(Z) = -GivenPoints(TP)
            Else
                If BasePoint(TP).Type = dsPoint Then
                    AddMacroSG = -meNotAllGivensSelected
                    Exit Function
                Else
                    If BasePoint(TP).Tag = 0 Then
                        hResult = AddMacroResult(tempMacro, Figures(TPParent))
                        If hResult <= 0 Then AddMacroSG = hResult: Exit Function
                        Figures(TPParent).Tag = hResult
                    End If
                    .Points(Z) = -BasePoint(TP).Tag
                End If
            End If

        Next Z
    End With
    
    AddMacroSG = .SGCount
End With
End Function

Public Sub RunMacro(ByVal Index As Long, MacroObjects() As Long)
Dim TL As Long, TPC As Long, pAction As Action, tPoints() As Long
Dim Z As Long, TFC As Long, Q As Long, i As Long, TempI As Long

'pAction.Type = actMacro
'MakeStructureSnapshot pAction
'ReDim pAction.AuxInfo(1 To 3)
'pAction.AuxInfo(1) = Macros(Index).FigurePointCount
'pAction.AuxInfo(2) = Macros(Index).ResultCount
'pAction.AuxInfo(3) = Macros(Index).SGCount
'RecordAction pAction

'============================================

TPC = PointCount
If PointCount + Macros(Index).FigurePointCount > MaxPointCount Then Exit Sub
PointCount = PointCount + Macros(Index).FigurePointCount
If FigureCount + Macros(Index).ResultCount > MaxFigureCount Then Exit Sub

If PointCount > 0 Then
    RedimPreserveBasePoint 1, PointCount
    For Z = 1 To Macros(Index).FigurePointCount
        BasePoint(TPC + Z) = Macros(Index).FigurePoints(Z)
        BasePoint(TPC + Z).Name = GenerateNewPointName
        'PointNames.Add BasePoint(TPC + Z).Name
        BasePoint(TPC + Z).ZOrder = GenerateNewPointZOrder
        BasePoint(TPC + Z).ParentFigure = FigureCount + BasePoint(TPC + Z).ParentFigure - 1
        BasePoint(TPC + Z).Locus = 0
        BasePoint(TPC + Z).Tag = 0
        'BasePoint(TPC + Z).Description = "He-he"
    Next Z
End If

TFC = FigureCount
FigureCount = FigureCount + Macros(Index).ResultCount
If FigureCount > 0 Then RedimPreserveFigures 0, FigureCount - 1
For Z = 0 To Macros(Index).ResultCount - 1
    Q = TFC + Z
    Figures(Q) = Macros(Index).Results(Z + 1)
    Select Case Figures(Q).FigureType
        Case dsSegment, dsRay, dsLine_2Points, dsBisector, dsCircle_CenterAndCircumPoint, dsCircle_CenterAndTwoPoints, dsMeasureDistance, dsMeasureAngle
            For i = 0 To Figures(Q).NumberOfPoints - 1
                If Figures(Q).Points(i) > 0 Then
                    Figures(Q).Points(i) = PointCount - Macros(Index).FigurePointCount + Figures(Q).Points(i)
                Else
                    Figures(Q).Points(i) = MacroObjects(-Figures(Q).Points(i))
                End If
            Next
        
        Case dsDynamicLocus
            For i = 0 To 1
                If Figures(Q).Points(i) > 0 Then
                    Figures(Q).Points(i) = PointCount - Macros(Index).FigurePointCount + Figures(Q).Points(i)
                Else
                    Figures(Q).Points(i) = MacroObjects(-Figures(Q).Points(i))
                End If
            Next
            
            For i = 0 To FigureCount - 1
                If Figures(i).FigureType = dsDynamicLocus Then
                    If Figures(i).Points(0) = Figures(Q).Points(0) Then
                        Figures(i).AuxInfo(4) = Infinity
                    End If
                End If
            Next
            
            Figures(Q).AuxInfo(4) = -Infinity
            
            If BasePoint(Figures(Q).Points(0)).Locus = 0 Then
                AddLocus Figures(Q).Points(0)
            Else
                EraseLocus BasePoint(Figures(Q).Points(0)).Locus
            End If
            
            With Locuses(BasePoint(Figures(Q).Points(0)).Locus)
                .Visible = True
                .Enabled = True
                .Dynamic = True
                .ParentPoint = Figures(Q).Points(0)
                .ParentFigure = GetLocusParentFigure(.ParentPoint)
                .LocusPointCount = setLocusDetails
                ReDim .LocusPixels(1 To .LocusPointCount)
                ReDim .LocusPoints(1 To .LocusPointCount)
            End With
        
        Case dsCircle_ArcCenterAndRadiusAndTwoPoints
            For i = 0 To 4
                If Figures(Q).Points(i) > 0 Then
                    Figures(Q).Points(i) = PointCount - Macros(Index).FigurePointCount + Figures(Q).Points(i)
                Else
                    Figures(Q).Points(i) = MacroObjects(-Figures(Q).Points(i))
                End If
            Next
            Figures(Q).Points(5) = PointCount - Macros(Index).FigurePointCount + Figures(Q).Points(5)
            Figures(Q).Points(6) = PointCount - Macros(Index).FigurePointCount + Figures(Q).Points(6)
        
        Case dsLine_PointAndParallelLine, dsLine_PointAndPerpendicularLine
            If Figures(Q).Points(0) > 0 Then
                Figures(Q).Points(0) = PointCount - Macros(Index).FigurePointCount + Figures(Q).Points(0)
            Else
                Figures(Q).Points(0) = MacroObjects(-Figures(Q).Points(0))
            End If
            If Figures(Q).Parents(0) > 0 Then
                Figures(Q).Parents(0) = Figures(Q).Parents(0) + TFC - 1
            Else
                Figures(Q).Parents(0) = MacroObjects(-Figures(Q).Parents(0))
                TL = Figures(Q).Parents(0)
                ReDim Preserve Figures(TL).Children(0 To Figures(TL).NumberOfChildren)
                Figures(TL).Children(Figures(TL).NumberOfChildren) = Q
                Figures(TL).NumberOfChildren = Figures(TL).NumberOfChildren + 1
            End If
        
        Case dsMiddlePoint, dsSimmPoint
            For i = 1 To 2
                If Figures(Q).Points(i) > 0 Then
                    Figures(Q).Points(i) = PointCount - Macros(Index).FigurePointCount + Figures(Q).Points(i)
                Else
                    Figures(Q).Points(i) = MacroObjects(-Figures(Q).Points(i))
                End If
            Next
            Figures(Q).Points(0) = PointCount - Macros(Index).FigurePointCount + Figures(Q).Points(0)
        
        Case dsAnPoint
            For i = 1 To Figures(Q).NumberOfPoints - 1
                If Figures(Q).Points(i) > 0 Then
                    SubstitutePointNamesInExpression Figures(Q).XS, Figures(Q).Points(i), PointCount - Macros(Index).FigurePointCount + Figures(Q).Points(i)
                    SubstitutePointNamesInExpression Figures(Q).YS, Figures(Q).Points(i), PointCount - Macros(Index).FigurePointCount + Figures(Q).Points(i)
                    Figures(Q).Points(i) = PointCount - Macros(Index).FigurePointCount + Figures(Q).Points(i)
                Else
                    SubstitutePointNamesInExpression Figures(Q).XS, Figures(Q).Points(i), MacroObjects(-Figures(Q).Points(i))
                    SubstitutePointNamesInExpression Figures(Q).YS, Figures(Q).Points(i), MacroObjects(-Figures(Q).Points(i))
                    Figures(Q).Points(i) = MacroObjects(-Figures(Q).Points(i))
                End If
            Next
            Figures(Q).XTree = BuildTree(Figures(Q).XS)
            Figures(Q).YTree = BuildTree(Figures(Q).YS)
            Figures(Q).Points(0) = PointCount - Macros(Index).FigurePointCount + Figures(Q).Points(0)
        
        Case dsSimmPointByLine, dsInvert
            If Figures(Q).Points(1) > 0 Then
                Figures(Q).Points(1) = PointCount - Macros(Index).FigurePointCount + Figures(Q).Points(1)
            Else
                Figures(Q).Points(1) = MacroObjects(-Figures(Q).Points(1))
            End If
            If Figures(Q).Parents(0) > 0 Then
                Figures(Q).Parents(0) = Figures(Q).Parents(0) + TFC - 1
            Else
                Figures(Q).Parents(0) = MacroObjects(-Figures(Q).Parents(0))
                TL = Figures(Q).Parents(0)
                ReDim Preserve Figures(TL).Children(0 To Figures(TL).NumberOfChildren)
                Figures(TL).Children(Figures(TL).NumberOfChildren) = Q
                Figures(TL).NumberOfChildren = Figures(TL).NumberOfChildren + 1
            End If
            Figures(Q).Points(0) = PointCount - Macros(Index).FigurePointCount + Figures(Q).Points(0)
        
        Case dsIntersect
            If Figures(Q).Parents(0) > 0 Then
                Figures(Q).Parents(0) = Figures(Q).Parents(0) + TFC - 1
            Else
                Figures(Q).Parents(0) = MacroObjects(-Figures(Q).Parents(0))
                TL = Figures(Q).Parents(0)
                ReDim Preserve Figures(TL).Children(0 To Figures(TL).NumberOfChildren)
                Figures(TL).Children(Figures(TL).NumberOfChildren) = Q
                Figures(TL).NumberOfChildren = Figures(TL).NumberOfChildren + 1
            End If
            Figures(Q).Points(0) = PointCount - Macros(Index).FigurePointCount + Figures(Q).Points(0)
            If Figures(Q).Parents(1) > 0 Then
                Figures(Q).Parents(1) = Figures(Q).Parents(1) + TFC - 1
            Else
                Figures(Q).Parents(1) = MacroObjects(-Figures(Q).Parents(1))
                TL = Figures(Q).Parents(1)
                ReDim Preserve Figures(TL).Children(0 To Figures(TL).NumberOfChildren)
                Figures(TL).Children(Figures(TL).NumberOfChildren) = Q
                Figures(TL).NumberOfChildren = Figures(TL).NumberOfChildren + 1
            End If
            Figures(Q).Points(1) = PointCount - Macros(Index).FigurePointCount + Figures(Q).Points(1)
        
        Case dsPointOnFigure
            If Figures(Q).Parents(0) > 0 Then
                Figures(Q).Parents(0) = Figures(Q).Parents(0) + TFC - 1
            Else
                Figures(Q).Parents(0) = MacroObjects(-Figures(Q).Parents(0))
                TL = Figures(Q).Parents(0)
                ReDim Preserve Figures(TL).Children(0 To Figures(TL).NumberOfChildren)
                Figures(TL).Children(Figures(TL).NumberOfChildren) = Q
                Figures(TL).NumberOfChildren = Figures(TL).NumberOfChildren + 1
            End If
            Figures(Q).Points(0) = PointCount - Macros(Index).FigurePointCount + Figures(Q).Points(0)
    End Select
    Figures(Q).Name = GenerateNewFigureName(Figures(Q).FigureType)
    'FigureNames.Add Figures(Q).Name
    Figures(Q).Tag = 0
    Figures(Q).ZOrder = GenerateNewFigureZOrder
    'Figures(Q).Description = "He-he"
Next Z

If Macros(Index).SGCount > 0 Then
    For Z = 1 To Macros(Index).SGCount
        tPoints = Macros(Index).SG(Z).Points
        For i = 1 To Macros(Index).SG(Z).NumberOfPoints
            If tPoints(i) > 0 Then
                tPoints(i) = PointCount - Macros(Index).FigurePointCount + tPoints(i)
            Else
                tPoints(i) = MacroObjects(-tPoints(i))
            End If
        Next
        
        AddStaticGraphic Macros(Index).SG(Z).Type, tPoints, False, False
        With StaticGraphics(StaticGraphicCount)
            .DrawMode = Macros(Index).SG(Z).DrawMode
            .DrawStyle = Macros(Index).SG(Z).DrawStyle
            .DrawWidth = Macros(Index).SG(Z).DrawWidth
            .FillColor = Macros(Index).SG(Z).FillColor
            .FillStyle = Macros(Index).SG(Z).FillStyle
            .ForeColor = Macros(Index).SG(Z).ForeColor
            .Visible = Macros(Index).SG(Z).Visible
        End With
    Next
End If

Z = 0
Do While Z < FigureCount
    If Figures(Z).AuxInfo(4) = Infinity Then
        TempI = BasePoint(Figures(Z).Points(0)).Locus
        DeleteFigure Z, False, False, True
        Locuses(TempI).Dynamic = True
    Else
        Z = Z + 1
    End If
Loop

For Z = 0 To FigureCount - 1
    If Figures(Z).FigureType = dsDynamicLocus And Figures(Z).AuxInfo(4) = -Infinity Then
        Figures(Z).AuxInfo(4) = 0
        BuildDynamicLocusDependency Figures(Z), Figures(Z).Points(0), Figures(Z).Points(1)
    End If
'    If Figures(Z).Description = "He-he" Then
'        Figures(Z).Description = GetObjectDescription(gotFigure, Z)
'    End If
Next

'For Z = 1 To PointCount
'    If BasePoint(Z).Description = "He-he" Then
'        BasePoint(Z).Description = GetObjectDescription(gotPoint, Z)
'    End If
'Next

FormMain.EnableMenus mnsStandard
RecalcAllAuxInfo
PaperCls
ShowAll
End Sub

Public Function SaveMacroAs(tMacro As Macro) As Boolean
Dim FName As String
CD.FileName = tMacro.Name & "." & extMAC
CD.Filter = GetString(ResDGMFiles) & "|*." & extMAC
CD.Flags = &H4 + &H2
If Not IsValidPath(LastMacroPath) Then LastMacroPath = ProgramPath
CD.InitDir = LastMacroPath
CD.DialogTitle = GetString(ResSaveAs)
CD.ShowSave
If CD.Cancelled Then Exit Function
FName = CD.FileName
If Right(UCase(FName), 4) <> "." & extMAC Then
    If InStr(RetrieveName(FName), ".") = 0 Then
        FName = FName & "." & extMAC
    Else
        FName = Left(FName, InStr(FName, ".")) & extMAC
    End If
End If
LastMacroPath = AddDirSep(RetrieveDir(FName))
If Not IsValidPath(LastMacroPath) Then LastMacroPath = ProgramPath
FormMain.Enabled = False
i_ShowStatus GetString(ResWorkingPleaseWait)
Screen.MousePointer = vbHourglass
DoEvents
SaveMacro tMacro, FName
i_ShowStatus
Screen.MousePointer = vbDefault
FormMain.Enabled = True

SaveMacroAs = True
End Function

'====================================================
' Returns True if actually added a macro, False otherwise.
'====================================================

Public Function LoadMacroAs(Optional ByVal Quiet As Boolean = False) As Boolean
CD.Filter = GetString(ResDGMFiles) & "|*." & extMAC
CD.Flags = &H1000 + &H4
If Not IsValidPath(LastMacroPath) Then LastMacroPath = ProgramPath

CD.InitDir = LastMacroPath
CD.DialogTitle = GetString(ResOpen)
CD.ShowOpen
If CD.Cancelled = True Or Dir(CD.FileName) = "" Then Exit Function

LastMacroPath = AddDirSep(RetrieveDir(CD.FileName))
If Not IsValidPath(LastMacroPath) Then LastMacroPath = ProgramPath
i_ShowStatus GetString(ResWorkingPleaseWait)

FormMain.Enabled = False
Screen.MousePointer = vbHourglass
DoEvents

LoadMacro CD.FileName, Not Quiet

i_ShowStatus
Screen.MousePointer = vbDefault
FormMain.Enabled = True
LoadMacroAs = True
End Function

Public Sub SaveMacro(tMacro As Macro, ByVal FName As String)
Dim Z As Long, Q As Long, NumOfParents As Long
On Local Error GoTo EH

Open FName For Output As #1
    Print #1, "[General]"
    Print #1, "FileFormatVersion=" & App.Major & "." & App.Minor & "." & App.Revision
    Print #1, "Name=" & tMacro.Name
    Print #1, "Description=" & ToSingleLine(tMacro.Description)
    Print #1, "GivenCount=" & tMacro.GivenCount
    Print #1, "FigurePointCount=" & tMacro.FigurePointCount
    Print #1, "ResultCount=" & tMacro.ResultCount
    Print #1, "SGCount=" & tMacro.SGCount
    Print #1,
    
    If tMacro.GivenCount > 0 Then
        For Z = 1 To tMacro.GivenCount
            Print #1, "[Given" & Z & "]"
            Print #1, "Type=" & tMacro.Givens(Z).Type
            Print #1, "Info=" & tMacro.Givens(Z).Description
            Print #1,
        Next Z
    End If
    
    If tMacro.FigurePointCount > 0 Then
        For Z = 1 To tMacro.FigurePointCount
            Print #1, "[Point" & Z & "]"
            
            Print #1, "FillColor=" & tMacro.FigurePoints(Z).FillColor
            Print #1, "FillStyle=" & tMacro.FigurePoints(Z).FillStyle
            Print #1, "ForeColor=" & tMacro.FigurePoints(Z).ForeColor
            If tMacro.FigurePoints(Z).Hide Then Print #1, "Hide=True"
            If Not tMacro.FigurePoints(Z).InDemo Then Print #1, "InDemo=False"
            Print #1, "Desc=" & tMacro.FigurePoints(Z).Description
            If tMacro.FigurePoints(Z).LabelOffsetX <> 0 Then Print #1, "LabelOffsetX=" & Trim(Str(BasePoint(Z).LabelOffsetX))
            If tMacro.FigurePoints(Z).LabelOffsetY <> -setdefPointSize \ 2 + 1 Then Print #1, "LabelOffsetY=" & Trim(Str(BasePoint(Z).LabelOffsetY))
            If tMacro.FigurePoints(Z).Locus <> 0 Then Print #1, "Locus=" & tMacro.FigurePoints(Z).Locus
            Print #1, "Name=" & tMacro.FigurePoints(Z).Name
            Print #1, "NameColor=" & tMacro.FigurePoints(Z).NameColor
            If tMacro.FigurePoints(Z).ParentFigure <> 0 Then Print #1, "ParentFigure=" & tMacro.FigurePoints(Z).ParentFigure
            Print #1, "PhysicalWidth=" & Trim(Str(tMacro.FigurePoints(Z).PhysicalWidth))
            Print #1, "Shape=" & tMacro.FigurePoints(Z).Shape
            Print #1, "ShowName=" & ToBoolean(tMacro.FigurePoints(Z).ShowName)
            If tMacro.FigurePoints(Z).ShowCoordinates Then Print #1, "ShowCoordinates=True"
            If tMacro.FigurePoints(Z).Tag <> 0 Then Print #1, "Tag=" & tMacro.FigurePoints(Z).Tag
            If Not tMacro.FigurePoints(Z).Enabled Then Print #1, "Enabled=False"
            Print #1, "Type=" & tMacro.FigurePoints(Z).Type
            Print #1, "Visible=" & ToBoolean(tMacro.FigurePoints(Z).Visible)
            Print #1, "X=" & Trim(Str(tMacro.FigurePoints(Z).X))
            Print #1, "Y=" & Trim(Str(tMacro.FigurePoints(Z).Y))
            Print #1, "ZOrder=" & tMacro.FigurePoints(Z).ZOrder
            Print #1,
        Next Z
    End If
    
    For Z = 1 To tMacro.ResultCount
        Print #1, "[Figure" & Z & "]"
        
        If tMacro.Results(Z).DrawMode <> defFigureDrawMode Then Print #1, "DrawMode=" & tMacro.Results(Z).DrawMode
        If tMacro.Results(Z).DrawStyle <> defFigureDrawStyle Or tMacro.Results(Z).FigureType = dsMeasureAngle Then Print #1, "DrawStyle=" & tMacro.Results(Z).DrawStyle
        If tMacro.Results(Z).FillColor <> colPolygonFillColor Then Print #1, "FillColor=" & tMacro.Results(Z).FillColor
        If tMacro.Results(Z).FillStyle <> 6 Then Print #1, "FillStyle=" & tMacro.Results(Z).FillStyle
        Print #1, "DrawWidth=" & tMacro.Results(Z).DrawWidth
        Print #1, "FigureName=" & tMacro.Results(Z).Name
        Print #1, "FigureType=" & tMacro.Results(Z).FigureType
        Print #1, "FigureTypeString=" & GetString(ResFigureBase + tMacro.Results(Z).FigureType * 2, langEnglish)
        Print #1, "ForeColor=" & tMacro.Results(Z).ForeColor
        Print #1, "NumberOfChildren=" & tMacro.Results(Z).NumberOfChildren
        Print #1, "NumberOfPoints=" & tMacro.Results(Z).NumberOfPoints
        Print #1, "Visible=" & ToBoolean(tMacro.Results(Z).Visible)
        If tMacro.Results(Z).Hide Then Print #1, "Hide=" & ToBoolean(tMacro.Results(Z).Hide)
        If Not tMacro.Results(Z).InDemo Then Print #1, "InDemo=False"
        Print #1, "Desc=" & tMacro.Results(Z).Description
        Print #1, "ZOrder=" & tMacro.Results(Z).ZOrder
        If tMacro.Results(Z).XS <> "" Then Print #1, "XS=" & tMacro.Results(Z).XS
        If tMacro.Results(Z).YS <> "" Then Print #1, "YS=" & tMacro.Results(Z).YS
        If tMacro.Results(Z).NumberOfChildren <> 0 Then
            For Q = 0 To tMacro.Results(Z).NumberOfChildren - 1
                Print #1, "Children" & (Q) & "=" & (tMacro.Results(Z).Children(Q))
            Next
        End If
        NumOfParents = GetProperParentNumber(tMacro.Results(Z).FigureType)
        If NumOfParents <> 0 Then
            For Q = 0 To NumOfParents - 1
                Print #1, "Parents" & (Q) & "=" & (tMacro.Results(Z).Parents(Q))
            Next
        End If
        For Q = 0 To tMacro.Results(Z).NumberOfPoints - 1
            Print #1, "Points" & (Q) & "=" & (tMacro.Results(Z).Points(Q))
        Next
        For Q = 1 To AuxCount
            If tMacro.Results(Z).AuxInfo(Q) <> 0 Then Print #1, "AuxInfo(" & Q & ")=" & Trim(Str(tMacro.Results(Z).AuxInfo(Q)))
            If tMacro.Results(Z).AuxPoints(Q).X <> 0 Then Print #1, "AuxPoints(" & Q & ").X=" & Trim(Str(tMacro.Results(Z).AuxPoints(Q).X))
            If tMacro.Results(Z).AuxPoints(Q).Y <> 0 Then Print #1, "AuxPoints(" & Q & ").Y=" & Trim(Str(tMacro.Results(Z).AuxPoints(Q).Y))
        Next
        
        Print #1,
    Next Z
    
    If tMacro.SGCount > 0 Then
        For Z = 1 To tMacro.SGCount
            Print #1, "[SG" & Z & "]"
            
            With tMacro.SG(Z)
                Print #1, "DrawMode=" & .DrawMode
                Print #1, "DrawStyle=" & .DrawStyle
                Print #1, "DrawWidth=" & .DrawWidth
                Print #1, "FillColor=" & .FillColor
                Print #1, "FillStyle=" & .FillStyle
                Print #1, "ForeColor=" & .ForeColor
                Print #1, "NumberOfPoints=" & .NumberOfPoints
                Print #1, "Type=" & .Type
                Print #1, "Visible=" & ToBoolean(.Visible)
                For Q = 1 To .NumberOfPoints
                    Print #1, "Point" & Q & "=" & .Points(Q)
                Next
            End With
            
            Print #1,
        Next
    End If
Close #1
Exit Sub

EH:
Reset
MsgBox GetString(ResUnableToSaveFile) & ":" & vbCrLf & ERR.Description, vbOKOnly + vbExclamation, GetString(ResError)
End Sub

Public Function LoadMacro(ByVal FName As String, Optional ByVal ShowProgress As Boolean = True) As Boolean
On Error GoTo EH:

Const ProgressStep = 100
Dim tempMacro As Macro, W As Double

Dim ErrStr As String, TotalProgress As Long, CurrentProgress As Long, ReadLength As Long, TotalLength As Long
Dim EqPos As Long, CurObj As Long, CurSect As String, CurSectType As Long
Dim AName As String, AValue As String, CurLine As Long
Dim BuildVersion As Long
Dim curChildren As Long, curParents As Long, curPoints As Long, curLetters As Long, curAuxInfo As Long, curAuxPointsX As Long, curAuxPointsY As Long
Dim tWEName As String, tWEExpression As String
Dim tX As Double, tY As Double
Dim tHide As Boolean
Dim A As String

If Dir$(FName) = "" Then ErrStr = "Cannot find file": GoTo EH

If ShowProgress Then
    TotalLength = FileLen(FName)
    ReadLength = 0
    ProgressShow GetString(ResLoading) & " " & RetrieveName(FName) & "..."
End If

BuildVersion = App.Revision
CurSect = "General"
CurSectType = 0
CurLine = 0

Open FName For Input As #1
    Do While True
        Do
            If EOF(1) Then A = "The End": Exit Do
            Line Input #1, A
            CurLine = CurLine + 1
            
            If ShowProgress Then
                ReadLength = ReadLength + Len(A) + 2
                If CurLine Mod ProgressStep = 0 Then ProgressUpdate ReadLength / TotalLength
            End If
            
            If IsSectionHeader(A) Then
'                If CurSectType = 2 Then
'                    If Not tHide Then
'                        tempMacro.Results(CurObj).Hide = Not tempMacro.Results(CurObj).Visible
'                    Else
'                        tHide = False
'                    End If
'                End If
                CurSect = GetSectionHeader(A): A = "//" & A
                If CurSect = "General" Then
                    CurSectType = 0
                    CurObj = 0
                ElseIf Left(CurSect, 5) = "Given" Then
                    CurSectType = 1
                    CurObj = Val(Right(CurSect, Len(CurSect) - 5))
                    '
                ElseIf Left(CurSect, 5) = "Point" Then
                    CurSectType = 2
                    CurObj = Val(Right(CurSect, Len(CurSect) - 5))
                    FillPointWithDefaults tempMacro.FigurePoints(CurObj)
                    '
                ElseIf Left(CurSect, 6) = "Figure" Then
                    CurSectType = 3
                    CurObj = Val(Right(CurSect, Len(CurSect) - 6))
                    FillFigureWithDefaults tempMacro.Results(CurObj)
                    curAuxInfo = 0
                    curAuxPointsX = 0
                    curAuxPointsY = 0
                    curChildren = 0
                    curLetters = 0
                    curParents = 0
                    curParents = 0
                    tHide = False
                    
                    '
                ElseIf Left(CurSect, 2) = "SG" Then
                    CurSectType = 4
                    CurObj = Val(Right(CurSect, Len(CurSect) - 2))
                    '
                End If
            End If
            EqPos = InStr(A, "=")
        Loop Until (Not IsComment(A)) And EqPos <> 0
        
        If A = "The End" Then
            '
            Exit Do
        End If
        
        AName = Left(A, EqPos - 1)
        AValue = Right(A, Len(A) - EqPos)
        
        Select Case CurSectType
            Case 0
                Select Case AName
                    Case "Name"
                        tempMacro.Name = AValue
                    Case "Description"
                        tempMacro.Description = ToMultiLine(AValue)
                    'Case "FileFormatVersion"
                    '    BuildVersion = Val(Right(AValue, Len(AValue) - InStrRev(AValue, ".")))
                    Case "GivenCount"
                        tempMacro.GivenCount = Val(AValue)
                        If tempMacro.GivenCount > 0 Then
                            ReDim tempMacro.Givens(1 To tempMacro.GivenCount)
                        End If
                    Case "FigurePointCount"
                        tempMacro.FigurePointCount = Val(AValue)
                        If tempMacro.FigurePointCount > 0 Then ReDim tempMacro.FigurePoints(1 To tempMacro.FigurePointCount)
                    Case "ResultCount"
                        tempMacro.ResultCount = Val(AValue)
                        If tempMacro.ResultCount > 0 Then ReDim tempMacro.Results(1 To tempMacro.ResultCount)
                    Case "SGCount"
                        tempMacro.SGCount = Val(AValue)
                        If tempMacro.SGCount > 0 Then ReDim tempMacro.SG(1 To tempMacro.SGCount)
                End Select
            
            Case 1
                If AName = "Type" Then tempMacro.Givens(CurObj).Type = Val(AValue)
                If AName = "Info" Then tempMacro.Givens(CurObj).Description = AValue
            
            Case 2
                With tempMacro.FigurePoints(CurObj)
                    Select Case AName
                        Case "Enabled"
                            .Enabled = FromBoolean(AValue)
                        Case "FillColor"
                            .FillColor = Val(AValue)
                        Case "FillStyle"
                            .FillStyle = Val(AValue)
                        Case "ForeColor"
                            .ForeColor = Val(AValue)
'                        Case "Hide"
'                            .Hide = FromBoolean(AValue)
                        Case "InDemo"
                            .InDemo = FromBoolean(AValue)
                        Case "Desc"
                            .Description = AValue
                        Case "LabelOffsetX"
                            .LabelOffsetX = Val(AValue)
                        Case "LabelOffsetY"
                            .LabelOffsetY = Val(AValue)
                        Case "Locus"
                            .Locus = Val(AValue)
                        Case "Name"
                            .Name = AValue
                            .LabelLength = Len(AValue)
                            .LabelHeight = Paper.TextHeight(.Name)
                            .LabelWidth = Paper.TextWidth(.Name)
                        Case "NameColor"
                            .NameColor = Val(AValue)
                        Case "ParentFigure"
                            .ParentFigure = Val(AValue)
                        Case "PhysicalWidth"
                            .PhysicalWidth = Val(AValue)
                            .Width = .PhysicalWidth
                            ToLogicalLength .Width
                        Case "Shape"
                            .Shape = Val(AValue)
                        Case "ShowCoordinates"
                            .ShowCoordinates = FromBoolean(AValue)
                        Case "ShowName"
                            .ShowName = FromBoolean(AValue)
                        Case "Tag"
                            .Tag = AValue
                        Case "Type"
                            .Type = Val(AValue)
                        Case "Visible"
                            .Visible = FromBoolean(AValue)
                        Case "X"
                            .X = Val(AValue)
                        Case "Y"
                            .Y = Val(AValue)
                        Case "ZOrder"
                            .ZOrder = Val(AValue)
                        Case "Hide"
                            .Hide = FromBoolean(AValue)
                    End Select
                End With
            
            Case 3
                With tempMacro.Results(CurObj)
                    If Left(AName, 8) = "Children" Then
                        curChildren = Right(AName, Len(AName) - 8)
                        If curChildren < .NumberOfChildren Then .Children(curChildren) = Val(AValue)
                    ElseIf Left(AName, 7) = "Parents" Then
                        curParents = Right(AName, Len(AName) - 7)
                        If curParents < GetProperParentNumber(.FigureType) Then .Parents(curParents) = Val(AValue)
                    ElseIf Left(AName, 6) = "Points" Then
                        curPoints = Right(AName, Len(AName) - 6)
                        If curPoints < .NumberOfPoints Then .Points(curPoints) = Val(AValue)
                    ElseIf Left(AName, 7) = "AuxInfo" Then
                        curAuxInfo = Mid(AName, 9, Len(AName) - 9)
                        If curAuxInfo <= AuxCount Then .AuxInfo(curAuxInfo) = Val(AValue)
                    ElseIf Left(AName, 9) = "AuxPoints" And Right(AName, 1) = "X" Then
                        curAuxPointsX = Mid(AName, 11, Len(AName) - 13)
                        If curAuxPointsX <= AuxCount Then .AuxPoints(curAuxPointsX).X = Val(AValue)
                    ElseIf Left(AName, 9) = "AuxPoints" And Right(AName, 1) = "Y" Then
                        curAuxPointsY = Mid(AName, 11, Len(AName) - 13)
                        If curAuxPointsY <= AuxCount Then .AuxPoints(curAuxPointsY).Y = Val(AValue)
                    End If
                    Select Case AName
                        Case "DrawMode"
                            .DrawMode = Val(AValue)
                        Case "DrawStyle"
                            .DrawStyle = Val(AValue)
                        Case "DrawWidth"
                            .DrawWidth = Val(AValue)
                        Case "FillColor"
                            .FillColor = Val(AValue)
                        Case "FillStyle"
                            .FillStyle = Val(AValue)
                        Case "FigureName"
                            .Name = AValue
                        Case "FigureType"
                            .FigureType = Val(AValue)
                            If GetProperParentNumber(.FigureType) > 0 Then ReDim .Parents(0 To GetProperParentNumber(.FigureType) - 1)
                            If .FigureType = dsMeasureAngle Then
                                .AuxInfo(2) = defAngleMarkRadius
                            End If
                        Case "FigureTypeString"
                            'do nothing
                        Case "Desc"
                            .Description = AValue
                        Case "ForeColor"
                            .ForeColor = Val(AValue)
                        Case "NumberOfChildren"
                            .NumberOfChildren = Val(AValue)
                            If .NumberOfChildren > 0 Then ReDim .Children(0 To .NumberOfChildren - 1)
                        Case "NumberOfPoints"
                            .NumberOfPoints = Val(AValue)
                            If .NumberOfPoints > 0 Then
                                ReDim .Points(0 To .NumberOfPoints - 1)
                            End If
                        Case "Visible"
                            .Visible = FromBoolean(AValue)
                        Case "XS"
                            .XS = AValue
                            If .FigureType = dsAnPoint Then .XTree = BuildTree(.XS)
                        Case "YS"
                            .YS = AValue
                            If .FigureType = dsAnPoint Then .YTree = BuildTree(.YS)
                        Case "Hide"
                            .Hide = FromBoolean(AValue)
                            tHide = True
                        Case "InDemo"
                            .InDemo = FromBoolean(AValue)
'                        Case "Desc"
'                            .Description = AValue
                        Case "ZOrder"
                            .ZOrder = Val(AValue)
                    End Select
                End With
            Case 4
                With tempMacro.SG(CurObj)
                    If Left(AName, 5) = "Point" Then
                        curPoints = Right(AName, Len(AName) - 5)
                        If curPoints <= .NumberOfPoints Then .Points(curPoints) = Val(AValue)
                    End If
                    Select Case AName
                        Case "DrawMode"
                            .DrawMode = Val(AValue)
                        Case "DrawStyle"
                            .DrawStyle = Val(AValue)
                        Case "DrawWidth"
                            .DrawWidth = Val(AValue)
                        Case "FillColor"
                            .FillColor = Val(AValue)
                        Case "FillStyle"
                            .FillStyle = Val(AValue)
                        Case "ForeColor"
                            .ForeColor = Val(AValue)
                        Case "NumberOfPoints"
                            .NumberOfPoints = Val(AValue)
                            If .NumberOfPoints > 0 Then
                                ReDim .Points(1 To .NumberOfPoints)
                                ReDim .ObjectPoints(1 To .NumberOfPoints)
                                ReDim .ObjectPixels(1 To .NumberOfPoints)
                            End If
                        Case "Type"
                            .Type = Val(AValue)
                        Case "Visible"
                            .Visible = FromBoolean(AValue)
                        Case "Desc"
                            .Description = AValue
                    End Select
                End With
        End Select
    Loop
Close #1

If ShowProgress Then
    ProgressClose
    If FormMain.Enabled And FormMain.Visible Then FormMain.SetFocus
End If

If tempMacro.ResultCount = 0 And tempMacro.SGCount = 0 Then MsgBox GetString(ResMsgCannotOpenFile) & " (" & FName & ")" & vbCrLf & "No figures or polygons.", vbExclamation: Close: Exit Function

AddMacro tempMacro
LoadMacro = True
Exit Function

EH:
Reset
If ShowProgress Then
    ProgressClose
    If FormMain.Enabled And FormMain.Visible Then FormMain.SetFocus
End If
MsgBox GetString(ResMsgCannotOpenFile) & vbCrLf & GetString(ResError) & ": " & ErrStr & vbCrLf & ERR.Description & vbCrLf & "Line number " & CurLine & ":" & vbCrLf & A, vbOKOnly + vbCritical, GetString(ResError)
End Function

'
''
''Public Sub SaveMacro(tMacro As Macro, ByVal FName As String)
''If Dir(FName, 23) <> "" Then Kill FName
''
''WritePrivateProfileStringByKeyName "General", "FileFormatVersion", App.Major & "." & App.Minor & "." & App.Revision, FName
''WritePrivateProfileStringByKeyName "General", "Name", tMacro.Name, FName
''WritePrivateProfileStringByKeyName "General", "Description", Replace(tMacro.Description, vbCrLf, "~~"), FName
''WritePrivateProfileStringByKeyName "General", "GivenCount", tMacro.GivenCount, FName
''WritePrivateProfileStringByKeyName "General", "FigurePointCount", tMacro.FigurePointCount, FName
''WritePrivateProfileStringByKeyName "General", "ResultCount", tMacro.ResultCount, FName
''
''If tMacro.GivenCount > 0 Then
''    For Z = 1 To tMacro.GivenCount
''        WritePrivateProfileStringByKeyName "Given" & Z, "Type", tMacro.Givens(Z).Type, FName
''    Next Z
''End If
''
''If tMacro.FigurePointCount > 0 Then
''    For Z = 1 To tMacro.FigurePointCount
''        WritePrivateProfileStringByKeyName "Point" & Z, "DrawWidth", tMacro.FigurePoints(Z).DrawWidth, FName
''        WritePrivateProfileStringByKeyName "Point" & Z, "FillColor", tMacro.FigurePoints(Z).FillColor, FName
''        WritePrivateProfileStringByKeyName "Point" & Z, "FillStyle", tMacro.FigurePoints(Z).FillStyle, FName
''        WritePrivateProfileStringByKeyName "Point" & Z, "ForeColor", tMacro.FigurePoints(Z).ForeColor, FName
''        WritePrivateProfileStringByKeyName "Point" & Z, "Height", Str(tMacro.FigurePoints(Z).Height), FName
''        WritePrivateProfileStringByKeyName "Point" & Z, "Hide", tMacro.FigurePoints(Z).Hide, FName
''        WritePrivateProfileStringByKeyName "Point" & Z, "LabelOffsetX", Str(BasePoint(Z).LabelOffsetX), FName
''        WritePrivateProfileStringByKeyName "Point" & Z, "LabelOffsetY", Str(BasePoint(Z).LabelOffsetY), FName
''        WritePrivateProfileStringByKeyName "Point" & Z, "Locus", tMacro.FigurePoints(Z).Locus, FName
''        WritePrivateProfileStringByKeyName "Point" & Z, "Name", tMacro.FigurePoints(Z).Name, FName
''        WritePrivateProfileStringByKeyName "Point" & Z, "NameColor", tMacro.FigurePoints(Z).NameColor, FName
''        WritePrivateProfileStringByKeyName "Point" & Z, "ParentFigure", tMacro.FigurePoints(Z).ParentFigure, FName
''        WritePrivateProfileStringByKeyName "Point" & Z, "PhysicalWidth", Str(tMacro.FigurePoints(Z).PhysicalWidth), FName
''        WritePrivateProfileStringByKeyName "Point" & Z, "Shape", tMacro.FigurePoints(Z).Shape, FName
''        WritePrivateProfileStringByKeyName "Point" & Z, "ShowName", tMacro.FigurePoints(Z).ShowName, FName
''        WritePrivateProfileStringByKeyName "Point" & Z, "ShowCoordinates", tMacro.FigurePoints(Z).ShowCoordinates, FName
''        WritePrivateProfileStringByKeyName "Point" & Z, "Tag", tMacro.FigurePoints(Z).Tag, FName
''        WritePrivateProfileStringByKeyName "Point" & Z, "Type", tMacro.FigurePoints(Z).Type, FName
''        WritePrivateProfileStringByKeyName "Point" & Z, "Visible", tMacro.FigurePoints(Z).Visible, FName
''        WritePrivateProfileStringByKeyName "Point" & Z, "Width", Str(tMacro.FigurePoints(Z).Width), FName
''        WritePrivateProfileStringByKeyName "Point" & Z, "X", Str(tMacro.FigurePoints(Z).X), FName
''        WritePrivateProfileStringByKeyName "Point" & Z, "Y", Str(tMacro.FigurePoints(Z).Y), FName
''        WritePrivateProfileStringByKeyName "Point" & Z, "ZOrder", tMacro.FigurePoints(Z).ZOrder, FName
''    Next Z
''End If
''
''For Z = 1 To tMacro.ResultCount
''    WritePrivateProfileStringByKeyName "Figure" & Z, "DrawMode", tMacro.Results(Z).DrawMode, FName
''    WritePrivateProfileStringByKeyName "Figure" & Z, "DrawStyle", tMacro.Results(Z).DrawStyle, FName
''    WritePrivateProfileStringByKeyName "Figure" & Z, "DrawWidth", tMacro.Results(Z).DrawWidth, FName
''    WritePrivateProfileStringByKeyName "Figure" & Z, "FigureName", tMacro.Results(Z).Name, FName
''    WritePrivateProfileStringByKeyName "Figure" & Z, "FigureType", tMacro.Results(Z).FigureType, FName
''    WritePrivateProfileStringByKeyName "Figure" & Z, "FigureTypeString", GetString(ResFigureBase + tMacro.Results(Z).FigureType * 2, langEnglish), FName
''    WritePrivateProfileStringByKeyName "Figure" & Z, "ForeColor", tMacro.Results(Z).ForeColor, FName
''    WritePrivateProfileStringByKeyName "Figure" & Z, "NumberOfChildren", tMacro.Results(Z).NumberOfChildren, FName
''    WritePrivateProfileStringByKeyName "Figure" & Z, "NumberOfPoints", tMacro.Results(Z).NumberOfPoints, FName
''    WritePrivateProfileStringByKeyName "Figure" & Z, "Visible", tMacro.Results(Z).Visible, FName
''    WritePrivateProfileStringByKeyName "Figure" & Z, "Hide", tMacro.Results(Z).Hide, FName
''    WritePrivateProfileStringByKeyName "Figure" & Z, "ZOrder", tMacro.Results(Z).ZOrder, FName
''    WritePrivateProfileStringByKeyName "Figure" & Z, "XS", tMacro.Results(Z).XS, FName
''    WritePrivateProfileStringByKeyName "Figure" & Z, "YS", tMacro.Results(Z).YS, FName
''    If tMacro.Results(Z).NumberOfChildren <> 0 Then
''        For Q = 0 To tMacro.Results(Z).NumberOfChildren - 1
''            WritePrivateProfileStringByKeyName "Figure" & Z, "Children" & (Q), (tMacro.Results(Z).Children(Q)), FName
''        Next
''    End If
''    NumOfParents = GetProperParentNumber(tMacro.Results(Z).FigureType)
''    If NumOfParents <> 0 Then
''        For Q = 0 To NumOfParents - 1
''            WritePrivateProfileStringByKeyName "Figure" & Z, "Parents" & (Q), (tMacro.Results(Z).Parents(Q)), FName
''        Next
''    End If
''    For Q = 0 To tMacro.Results(Z).NumberOfPoints - 1
''        WritePrivateProfileStringByKeyName "Figure" & Z, "Points" & (Q), (tMacro.Results(Z).Points(Q)), FName
''        WritePrivateProfileStringByKeyName "Figure" & Z, "Letters" & (Q), (tMacro.Results(Z).Letters(Q)), FName
''    Next
''    For Q = 1 To AuxCount
''        WritePrivateProfileStringByKeyName "Figure" & Z, "AuxInfo(" & Q & ")", Str(tMacro.Results(Z).AuxInfo(Q)), FName
''        WritePrivateProfileStringByKeyName "Figure" & Z, "AuxPoints(" & Q & ").X", Str(tMacro.Results(Z).AuxPoints(Q).X), FName
''        WritePrivateProfileStringByKeyName "Figure" & Z, "AuxPoints(" & Q & ").Y", Str(tMacro.Results(Z).AuxPoints(Q).Y), FName
''    Next
''Next Z
''End Sub
'

Public Sub AutoLoadMacros()
Dim Z As Long
On Local Error Resume Next
Dim tFileList() As String, S As String
If setMacroAutoloadPath = "" Or Dir(setMacroAutoloadPath, 23) = "" Then Exit Sub
ProgressShow GetString(ResPleaseWaitLoadingMacros)
ReDim tFileList(0 To 0)
tFileList = PrepareFileList(setMacroAutoloadPath, "*." & extMAC)
For Z = 1 To UBound(tFileList)
    If Not LoadMacro(tFileList(Z), False) Then ProgressRefresh
    ProgressUpdate Z / UBound(tFileList)
Next Z
ProgressClose
End Sub

Public Sub MacroRunMenu(ByVal Index As Long)
Dim Z As Long

RecordGenericAction ResUndoMacro

DragS.MacroObjectCount = Macros(Index).GivenCount
i_SelectTool dsSelect

If DragS.MacroObjectCount > 0 Then
    i_EnterMacroRunMode
    
    ReDim DragS.MacroObjects(1 To DragS.MacroObjectCount)
    ReDim DragS.MacroObjectType(1 To DragS.MacroObjectCount)
    ReDim DragS.MacroObjectDescription(1 To DragS.MacroObjectCount)
    For Z = 1 To DragS.MacroObjectCount
        DragS.MacroObjectType(Z) = Macros(Index).Givens(Z).Type
        DragS.MacroObjectDescription(Z) = Macros(Index).Givens(Z).Description
    Next
    DragS.MacroCurrentObject = 1
    DragS.State = dscMacroStateRun
    DragS.Number(1) = Index
Else
    ReDim DragS.MacroObjects(1 To 1)
    RunMacro Index, DragS.MacroObjects
End If
End Sub

Public Function MacroCanBeCreated() As Boolean
If FigureCount = 0 And StaticGraphicCount = 0 Then
    MsgBox GetString(ResMacroErrBase + 2 * MacroErrors.meAtLeastOneFigureNeeded), vbCritical
    MacroCanBeCreated = False
    Exit Function
End If
MacroCanBeCreated = True
End Function

Public Sub MacroCancel(ByVal ErrStr As String)
Dim Z As Long

DragS.State = dscNormalState
For Z = 1 To MaxDragNumbers: DragS.Number(Z) = 0:  Next

InitZeroTags

PaperCls
ShowAll
ImitateMouseMove

If ErrStr <> "" Then MsgBox ErrStr, vbOKOnly + vbExclamation, GetString(ResMacroErrBase + 2 * MacroErrors.meErrorCreatingMacro)
End Sub

Public Sub MacroCreateInit()
If Not MacroCanBeCreated Then Exit Sub

If setShowMacroCreateDialog Then
    frmMacroCreate.Show vbModal
Else
    MacroCreateBegin
End If
End Sub

Public Sub MacroCreateBegin()
InitZeroTags

' Clear tempMacro
tempMacro.GivenCount = 0
tempMacro.ResultCount = 0
tempMacro.FigurePointCount = 0
tempMacro.SGCount = 0
ReDim tempMacro.Givens(1 To 1)
ReDim tempMacro.Results(1 To 1)
ReDim tempMacro.FigurePoints(1 To 1)
ReDim tempMacro.SG(1 To 1)
tempMacro.Description = ""
tempMacro.Name = ""

' Clear GivenPoints and GivenFigures
If PointCount > 0 Then ReDim GivenPoints(1 To PointCount)
If FigureCount > 0 Then ReDim GivenFigures(0 To FigureCount - 1) Else ReDim GivenFigures(0 To 0)

' Notify environment of state changes
i_SelectTool dsSelect
i_EnterMacroGivenSelectMode

PaperCls
ShowAllWithGivens
End Sub

Public Sub MacroCreateResultsInit()
If setShowMacroResultsDialog Then
    frmMacroSelectResults.Show vbModal
Else
    MacroCreateResults
End If
End Sub

Public Sub MacroSaveInit()
frmMacroSave.Show vbModal
End Sub

Public Sub MacroCreateResults()
Dim Z As Long

If FigureCount > 0 Then ReDim ResultFigures(0 To FigureCount - 1) Else ReDim ResultFigures(0 To 0)
If StaticGraphicCount > 0 Then ReDim ResultSGs(1 To StaticGraphicCount) Else ReDim ResultSGs(1 To 1)
If LocusCount > 0 Then ReDim ResultLoci(1 To LocusCount) Else ReDim ResultLoci(1 To 1)
If PointCount > 0 Then ReDim ResultPoints(1 To PointCount) Else ReDim ResultPoints(1 To 1)

For Z = 0 To FigureCount - 1
    ResultFigures(Z) = Not CanBeAResult(Z)
Next

For Z = 1 To StaticGraphicCount
    ResultSGs(Z) = Not CanBeAResultSG(Z)
Next

For Z = 1 To LocusCount
    ResultLoci(Z) = Not CanBeAResultLocus(Z)
Next

For Z = 1 To PointCount
    ResultPoints(Z) = Not CanBeAResultPoint(Z)
Next

i_EnterMacroResultSelectMode

PaperCls
ShowAllWithResults
End Sub

Public Sub MacroCreateSave()
Dim Z As Long, S As String

InitZeroTags

i_ExitMacroResultSelectMode

PaperCls
ShowAll

If tempMacro.ResultCount = 0 And tempMacro.SGCount = 0 Then MsgBox GetString(ResMacroErrBase), vbExclamation: Exit Sub

'S = GetString(ResEnterMacroName)
'Do While tempMacro.Name = ""
'    tempMacro.Name = InputBox(S, GetString(ResMacro), GetString(ResMacro) & (MacroCount + 1))
'    If tempMacro.Name = "" Then If MsgBox(GetString(ResMsgDoYouReallyWantToCancel), vbYesNo, GetString(ResMacro)) = vbYes Then Exit Sub
'    S = GetString(ResEnterMacroName)
'    For Z = 1 To MacroCount
'        If Macros(Z).Name = tempMacro.Name Then
'            S = tempMacro.Name & GetString(ResMsgObjectAlreadyExists)
'            tempMacro.Name = ""
'        End If
'    Next
'Loop

If Not SaveMacroAs(tempMacro) Then Exit Sub
AddMacro tempMacro

End Sub

'=======================================================
'           Fills Figures(x).Tag with zeroes; Basepoint(x).Tag and SGs(x).Tag
' BitField indicates whether to fill:
'   Figures.Tag                 -     bit 0 set (+1)
'   Basepoint.Tag             -     bit 1 set (+2)
'   StaticGraphics.Tag      -     bit 2 set (+4)
'   _________________________________
'                                                            = BitField
'=======================================================

Public Sub InitZeroTags(Optional ByVal BitField As Long = 7)
Dim Z As Long

If (BitField And 1) = 1 Then
    For Z = 0 To FigureCount - 1
        Figures(Z).Tag = 0
    Next
End If

If (BitField And 2) = 2 Then
    For Z = 1 To PointCount
        BasePoint(Z).Tag = 0
    Next
End If

If (BitField And 4) = 4 Then
    For Z = 1 To StaticGraphicCount
        StaticGraphics(Z).Tag = 0
    Next
End If
End Sub

Public Sub ShowAllWithGivens(Optional ByVal hDC As Long = 0)
Dim Z As Long
If hDC = 0 Then hDC = Paper.hDC

ShowAll , True, False

For Z = 0 To FigureCount - 1
    If GivenFigures(Z) > 0 Then
        ShowSelectedFigure hDC, Z
    Else
        DrawFigure hDC, Z
    End If
Next

For Z = 1 To PointCount
    If GivenPoints(Z) > 0 Then
        ShowSelectedPoint hDC, Z
    Else
        ShowPoint hDC, Z
    End If
Next

Paper.Refresh
End Sub

Public Sub ShowAllWithResults(Optional ByVal hDC As Long = 0)
Dim Z As Long
If hDC = 0 Then hDC = Paper.hDC

If setWallpaper <> "" Then DrawWallPaper hDC Else If nGradientPaper Then Gradient hDC, nPaperColor1, nPaperColor2, 0, 0, PaperScaleWidth, PaperScaleHeight, False ' Else PaperCls
If nShowGrid Then ShowGrid hDC, True
If nShowAxes Then ShowAxes hDC, True

For Z = 1 To LabelCount
    ShowLabel hDC, Z, , , , True
Next

For Z = 1 To StaticGraphicCount
    If ResultSGs(Z) > 0 Then
        DrawStaticGraphicSelected hDC, Z
    Else
        DrawStaticGraphic hDC, Z, , , ResultSGs(Z) = -1
    End If
Next

For Z = 1 To LocusCount
    If ResultLoci(Z) > 0 Then
        ShowLocusSelected hDC, Z
    Else
        ShowLocus hDC, Z, , , , ResultLoci(Z) = -1
    End If
Next

For Z = 0 To FigureCount - 1
    DrawFigure hDC, Z, False, , ResultFigures(Z) = -1
Next

For Z = 1 To PointCount
    ShowPoint hDC, Z, , , , ResultPoints(Z) = -1
Next

For Z = 1 To ButtonCount
    ShowButton hDC, Z, , , , , , True
Next

'==============================================


For Z = 0 To FigureCount - 1
    If ResultFigures(Z) > 0 Then
        If IsVisual(Z) And Not tempMacro.Results(ResultFigures(Z)).Hide Then
            ShowSelectedFigure hDC, Z
        End If
    End If
Next

For Z = 0 To FigureCount - 1
    If ResultFigures(Z) > 0 Then
        If Not IsVisual(Z) And Not tempMacro.Results(ResultFigures(Z)).Hide Then
            ShowSelectedFigure hDC, Z
        End If
    End If
Next

Paper.Refresh
End Sub

Public Function IsResultSelected(ByVal Index As Long)
Dim tBool As Boolean
tBool = IsResultAdded(Index)
If tBool Then tBool = Not tempMacro.Results(ResultFigures(Index)).Hide
IsResultSelected = tBool
End Function

Public Function IsResultAdded(ByVal Index As Long) As Boolean
IsResultAdded = ResultFigures(Index) > 0
End Function

Public Function IsSGAdded(ByVal Index As Long) As Boolean
IsSGAdded = ResultSGs(Index) > 0
End Function

Public Function CanBeAResult(ByVal Index As Long) As Boolean
Dim TP As Long, TPParent As Long, Z As Long

If Not IsFigure(Index) Then Exit Function

With Figures(Index)
    For Z = 0 To .NumberOfPoints - 1
        If Not IsChildPointPos(Figures(Index), Z) Then
            If Not CanBeAResultParentPoint(.Points(Z)) Then
                CanBeAResult = False
                Exit Function
            End If
        End If
    Next Z
            
    '#############################################3
            
    For Z = 0 To GetProperParentNumber(.FigureType) - 1
        If Not CanBeAResultParentFigure(.Parents(Z)) Then
            CanBeAResult = False
            Exit Function
        End If
    Next
            
End With

CanBeAResult = True
End Function

Public Function CanBeAResultParentPoint(ByVal TP As Long) As Boolean
If Not IsPoint(TP) Then Exit Function

If GivenPoints(TP) = 0 Then
    If BasePoint(TP).Type = dsPoint Then
        CanBeAResultParentPoint = False
        Exit Function
    Else
        If Not CanBeAResultParentFigure(BasePoint(TP).ParentFigure) Then
            CanBeAResultParentPoint = False
            Exit Function
        End If
    End If
End If

CanBeAResultParentPoint = True
End Function

Public Function CanBeAResultParentFigure(ByVal Index As Long) As Boolean
If Not IsFigure(Index) Then Exit Function

If GivenFigures(Index) > 0 Then CanBeAResultParentFigure = True: Exit Function

CanBeAResultParentFigure = CanBeAResult(Index)
End Function

Public Function CanBeAResultSG(ByVal Index As Long) As Boolean
Dim Z As Long

If Not IsSG(Index) Then Exit Function

With StaticGraphics(Index)
    For Z = 1 To .NumberOfPoints
        If Not CanBeAResultParentPoint(.Points(Z)) Then
            CanBeAResultSG = False
            Exit Function
        End If
    Next
End With

CanBeAResultSG = True
End Function

Public Function CanBeAResultLocus(ByVal Index As Long) As Boolean
If Not IsLocus(Index) Then Exit Function
If Not Locuses(Index).Dynamic Then Exit Function

CanBeAResultLocus = CanBeAResult(Locuses(Index).ParentFigure)
End Function

Public Function CanBeAResultPoint(ByVal Index As Long) As Boolean
If Not IsPoint(Index) Then Exit Function

If BasePoint(Index).Type = dsPoint Or GivenPoints(Index) > 0 Then CanBeAResultPoint = False: Exit Function
CanBeAResultPoint = CanBeAResult(BasePoint(Index).ParentFigure)
End Function

Public Sub MacroOrganizeShow()
frmMacrosOrganize.Show vbModal
End Sub
