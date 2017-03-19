Attribute VB_Name = "modObjectList"
Option Explicit

Public GlobalCalc As ctlCalculator

Public Function ObjectSelectionBegin(ByVal oType As ObjectSelectionType, Optional ByVal ShouldClearTempObjectSelection As Boolean = True, Optional ByVal Caller As ObjectSelectionCaller = oscButton) As Boolean
If ShouldClearTempObjectSelection Then ObjectListClear TempObjectSelection
TempObjectSelection.SubType = Caller
TempObjectSelection.Type = oType
FormMain.EnterObjectSelectionMode oType
End Function

Public Function ObjectSelectionComplete()
Select Case TempObjectSelection.Type
Case ostShowHideObjects
    FormMain.Enabled = False
    ObjectSelectionEnd
    Select Case TempObjectSelection.SubType
    Case oscButton
        frmButtonProps.ObjectSelectionComplete_ShowHide
    Case Else
    End Select
Case ostCalcPoints
    FormMain.Enabled = False
    ObjectSelectionEnd
    GlobalCalc.ObjectSelectionComplete_CalcPoints
'    Select Case TempObjectSelection.SubType
'    Case oscCalcLabels
'        frmLabelProps.ctlCalculator1.ObjectSelectionComplete_CalcPoints
'    Case oscCalcAnPoint
'        frmAnPoint.ctlCalculator1.ObjectSelectionComplete_CalcPoints
'    Case oscCalculator
'        frmCalculator.ctlCalculator1.ObjectSelectionComplete_CalcPoints
'    Case Else
'    End Select
End Select
End Function

Public Function ObjectSelectionCancel()
Select Case TempObjectSelection.Type
Case ostShowHideObjects
    FormMain.Enabled = False
    ObjectSelectionEnd
    Select Case TempObjectSelection.SubType
    Case oscButton
        frmButtonProps.ObjectSelectionCancel_ShowHide
    End Select
Case ostCalcPoints
    FormMain.Enabled = False
    ObjectSelectionEnd
    GlobalCalc.ObjectSelectionCancel_CalcPoints
'    Select Case TempObjectSelection.SubType
'    Case oscCalcLabels
'        frmLabelProps.ctlCalculator1.ObjectSelectionCancel_CalcPoints
'    Case oscCalcAnPoint
'        frmAnPoint.ctlCalculator1.ObjectSelectionCancel_CalcPoints
'    End Select
End Select
End Function

Public Function ObjectSelectionEnd()
FormMain.ExitObjectSelectionMode
End Function

Public Sub ObjectListClear(objList As ObjectList)
Dim tempTwoPoints As TwoPoints
With objList
    .BoundingRect = tempTwoPoints
    .FigureCount = 0
    .LabelCount = 0
    .LocusCount = 0
    .PointCount = 0
    .SGCount = 0
    .WECount = 0
    .ButtonCount = 0
    .TotalCount = 0
    .FigureCountMax = 0
    .LabelCountMax = 0
    .LocusCountMax = 0
    .PointCountMax = 0
    .SGCountMax = 0
    .WECountMax = 0
    .ButtonCountMax = 0
    .TotalCountMax = 0
    ReDim .Figures(1 To 1)
    ReDim .Labels(1 To 1)
    ReDim .Loci(1 To 1)
    ReDim .Points(1 To 1)
    ReDim .SGs(1 To 1)
    ReDim .WEs(1 To 1)
    ReDim .Buttons(1 To 1)
    .Type = ostShowHideObjects
End With
End Sub

Public Function ObjectListFind(objList As ObjectList, Category As GeometryObjectType, ByVal Index As Long) As Long
Dim Z As Long

With objList
    Select Case Category
    Case gotGeneric
    Case gotPoint
        For Z = 1 To .PointCount
            If .Points(Z) = Index Then ObjectListFind = Z: Exit Function
        Next
    Case gotFigure
        For Z = 1 To .FigureCount
            If .Figures(Z) = Index Then ObjectListFind = Z: Exit Function
        Next
    Case gotLabel
        For Z = 1 To .LabelCount
            If .Labels(Z) = Index Then ObjectListFind = Z: Exit Function
        Next
    Case gotLocus
        For Z = 1 To .LocusCount
            If .Loci(Z) = Index Then ObjectListFind = Z: Exit Function
        Next
    Case gotSG
        For Z = 1 To .SGCount
            If .SGs(Z) = Index Then ObjectListFind = Z: Exit Function
        Next
    Case gotWE
        For Z = 1 To .WECount
            If .WEs(Z) = Index Then ObjectListFind = Z: Exit Function
        Next
    Case gotButton
        For Z = 1 To .ButtonCount
            If .Buttons(Z) = Index Then ObjectListFind = Z: Exit Function
        Next
    End Select
End With
End Function

Public Function ObjectListReplace(objList As ObjectList, Category As GeometryObjectType, ByVal Index As Long, ByVal NewValue As Long) As Long
Dim Z As Long

With objList
    Select Case Category
    Case gotGeneric
    Case gotPoint
        For Z = 1 To .PointCount
            If .Points(Z) = Index Then .Points(Z) = NewValue
        Next
    Case gotFigure
        For Z = 1 To .FigureCount
            If .Figures(Z) = Index Then .Figures(Z) = NewValue
        Next
    Case gotLabel
        For Z = 1 To .LabelCount
            If .Labels(Z) = Index Then .Labels(Z) = NewValue
        Next
    Case gotLocus
        For Z = 1 To .LocusCount
            If .Loci(Z) = Index Then .Loci(Z) = NewValue
        Next
    Case gotSG
        For Z = 1 To .SGCount
            If .SGs(Z) = Index Then .SGs(Z) = NewValue
        Next
    Case gotWE
        For Z = 1 To .WECount
            If .WEs(Z) = Index Then .WEs(Z) = NewValue
        Next
    Case gotButton
        For Z = 1 To .ButtonCount
            If .Buttons(Z) = Index Then .Buttons(Z) = NewValue
        Next
    End Select
End With
End Function

Public Function ObjectListAdd(objList As ObjectList, Category As GeometryObjectType, ByVal Index As Long) As Long
If ObjectListFind(objList, Category, Index) <> 0 Then Exit Function

With objList
    .TotalCount = .TotalCount + 1
    Select Case Category
    Case gotGeneric
    Case gotPoint
        .PointCount = .PointCount + 1
        ReDim Preserve .Points(1 To .PointCount)
        .Points(.PointCount) = Index
    Case gotFigure
        .FigureCount = .FigureCount + 1
        ReDim Preserve .Figures(1 To .FigureCount)
        .Figures(.FigureCount) = Index
    Case gotLabel
        .LabelCount = .LabelCount + 1
        ReDim Preserve .Labels(1 To .LabelCount)
        .Labels(.LabelCount) = Index
    Case gotLocus
        .LocusCount = .LocusCount + 1
        ReDim Preserve .Loci(1 To .LocusCount)
        .Loci(.LocusCount) = Index
    Case gotSG
        .SGCount = .SGCount + 1
        ReDim Preserve .SGs(1 To .SGCount)
        .SGs(.SGCount) = Index
    Case gotWE
        .WECount = .WECount + 1
        ReDim Preserve .WEs(1 To .WECount)
        .WEs(.WECount) = Index
    Case gotButton
        .ButtonCount = .ButtonCount + 1
        ReDim Preserve .Buttons(1 To .ButtonCount)
        .Buttons(.ButtonCount) = Index
    End Select
End With
End Function

Public Function ObjectListDelete(objList As ObjectList, Category As GeometryObjectType, ByVal Index As Long) As Long
Dim Z As Long

If Index = 0 Then Exit Function
ObjectListDelete = Index

With objList
    .TotalCount = .TotalCount - 1
    Select Case Category
    Case gotGeneric
    Case gotPoint
        If Index < .PointCount Then
            For Z = Index To .PointCount - 1
                .Points(Z) = .Points(Z + 1)
            Next
        End If
        .PointCount = .PointCount - 1
        If .PointCount > 0 Then ReDim Preserve .Points(1 To .PointCount)
    Case gotFigure
        If Index < .FigureCount Then
            For Z = Index To .FigureCount - 1
                .Figures(Z) = .Figures(Z + 1)
            Next
        End If
        .FigureCount = .FigureCount - 1
        If .FigureCount > 0 Then ReDim Preserve .Figures(1 To .FigureCount)
    Case gotLabel
        If Index < .LabelCount Then
            For Z = Index To .LabelCount - 1
                .Labels(Z) = .Labels(Z + 1)
            Next
        End If
        .LabelCount = .LabelCount - 1
        If .LabelCount > 0 Then ReDim Preserve .Labels(1 To .LabelCount)
    Case gotLocus
        If Index < .LocusCount Then
            For Z = Index To .LocusCount - 1
                .Loci(Z) = .Loci(Z + 1)
            Next
        End If
        .LocusCount = .LocusCount - 1
        If .LocusCount > 0 Then ReDim Preserve .Loci(1 To .LocusCount)
    Case gotSG
        If Index < .SGCount Then
            For Z = Index To .SGCount - 1
                .SGs(Z) = .SGs(Z + 1)
            Next
        End If
        .SGCount = .SGCount - 1
        If .SGCount > 0 Then ReDim Preserve .SGs(1 To .SGCount)
    Case gotWE
        If Index < .WECount Then
            For Z = Index To .WECount - 1
                .WEs(Z) = .WEs(Z + 1)
            Next
        End If
        .WECount = .WECount - 1
        If .WECount > 0 Then ReDim Preserve .WEs(1 To .WECount)
    Case gotButton
        If Index < .ButtonCount Then
            For Z = Index To .ButtonCount - 1
                .Buttons(Z) = .Buttons(Z + 1)
            Next
        End If
        .ButtonCount = .ButtonCount - 1
        If .ButtonCount > 0 Then ReDim Preserve .Buttons(1 To .ButtonCount)
    End Select
End With
End Function

Public Function ObjectListRemove(objList As ObjectList, Category As GeometryObjectType, ByVal Index As Long) As Long
Dim Z As Long

Index = ObjectListFind(objList, Category, Index)
ObjectListRemove = Index
If Index = 0 Then Exit Function

With objList
    .TotalCount = .TotalCount - 1
    Select Case Category
    Case gotGeneric
    Case gotPoint
        If Index < .PointCount Then
            For Z = Index To .PointCount - 1
                .Points(Z) = .Points(Z + 1)
            Next
        End If
        .PointCount = .PointCount - 1
        If .PointCount > 0 Then ReDim Preserve .Points(1 To .PointCount)
    Case gotFigure
        If Index < .FigureCount Then
            For Z = Index To .FigureCount - 1
                .Figures(Z) = .Figures(Z + 1)
            Next
        End If
        .FigureCount = .FigureCount - 1
        If .FigureCount > 0 Then ReDim Preserve .Figures(1 To .FigureCount)
    Case gotLabel
        If Index < .LabelCount Then
            For Z = Index To .LabelCount - 1
                .Labels(Z) = .Labels(Z + 1)
            Next
        End If
        .LabelCount = .LabelCount - 1
        If .LabelCount > 0 Then ReDim Preserve .Labels(1 To .LabelCount)
    Case gotLocus
        If Index < .LocusCount Then
            For Z = Index To .LocusCount - 1
                .Loci(Z) = .Loci(Z + 1)
            Next
        End If
        .LocusCount = .LocusCount - 1
        If .LocusCount > 0 Then ReDim Preserve .Loci(1 To .LocusCount)
    Case gotSG
        If Index < .SGCount Then
            For Z = Index To .SGCount - 1
                .SGs(Z) = .SGs(Z + 1)
            Next
        End If
        .SGCount = .SGCount - 1
        If .SGCount > 0 Then ReDim Preserve .SGs(1 To .SGCount)
    Case gotWE
        If Index < .WECount Then
            For Z = Index To .WECount - 1
                .WEs(Z) = .WEs(Z + 1)
            Next
        End If
        .WECount = .WECount - 1
        If .WECount > 0 Then ReDim Preserve .WEs(1 To .WECount)
    Case gotButton
        If Index < .ButtonCount Then
            For Z = Index To .ButtonCount - 1
                .Buttons(Z) = .Buttons(Z + 1)
            Next
        End If
        .ButtonCount = .ButtonCount - 1
        If .ButtonCount > 0 Then ReDim Preserve .Buttons(1 To .ButtonCount)
    End Select
End With
End Function

Public Function ObjectListAddRemove(objList As ObjectList, Category As GeometryObjectType, ByVal Index As Long) As Boolean
Dim Z As Long
Z = ObjectListFind(objList, Category, Index)
If Z = 0 Then
    ObjectListAdd objList, Category, Index
    ObjectListAddRemove = True
Else
    ObjectListRemove objList, Category, Index
    ObjectListAddRemove = False
End If
End Function

Public Sub FillListBoxWithObjectList(lstListBox As ListBox, tObjectlist As ObjectList, Optional ByVal AnySpecificType As GeometryObjectType)
Dim Z As Long

With tObjectlist
    lstListBox.Clear
    Select Case AnySpecificType
    
    Case gotGeneric
        For Z = 1 To .PointCount
            lstListBox.AddItem BasePoint(.Points(Z)).Name
        Next
        For Z = 1 To .FigureCount
            lstListBox.AddItem Figures(.Figures(Z)).Name
        Next
        For Z = 1 To .LabelCount
            lstListBox.AddItem GetString(ResLabel) & .Labels(Z)
        Next
        For Z = 1 To .LocusCount
            lstListBox.AddItem GetString(ResFigureBase + 2 * dsDynamicLocus) & .Loci(Z)
        Next
        For Z = 1 To .SGCount
            lstListBox.AddItem GetString(ResStaticObjectBase + StaticGraphics(.SGs(Z)).Type * 2) & .SGs(Z)
        Next
        For Z = 1 To .WECount
            lstListBox.AddItem WatchExpressions(.WEs(Z)).Name
        Next
        For Z = 1 To .ButtonCount
            lstListBox.AddItem Buttons(.Buttons(Z)).Caption
        Next
    
    Case gotPoint
        For Z = 1 To .PointCount
            lstListBox.AddItem BasePoint(.Points(Z)).Name
        Next
    
    Case gotFigure
        For Z = 1 To .FigureCount
            lstListBox.AddItem Figures(.Figures(Z)).Name
        Next
    
    Case gotLabel
        For Z = 1 To .LabelCount
            lstListBox.AddItem GetString(ResLabel) & Z
        Next
    
    Case gotLocus
        For Z = 1 To .LocusCount
            lstListBox.AddItem GetString(ResFigureBase + 2 * dsDynamicLocus) & Z
        Next
    
    Case gotSG
        For Z = 1 To .SGCount
            lstListBox.AddItem GetString(ResStaticObjectBase + StaticGraphics(.SGs(Z)).Type * 2) & Z
        Next
    
    Case gotWE
        For Z = 1 To .WECount
            lstListBox.AddItem WatchExpressions(.WEs(Z)).Name
        Next
    
    Case gotButton
        For Z = 1 To .ButtonCount
            lstListBox.AddItem Buttons(.Buttons(Z)).Caption
        Next
    End Select
End With
AddListboxScrollbar lstListBox
End Sub

Public Sub ObjectListGetItemFromGenericIndex(objList As ObjectList, ByVal Index As Long, objIndex As Long, ObjType As GeometryObjectType)
If Index = 0 Then objIndex = 0: Exit Sub

With objList
    If Index <= .PointCount Then
        ObjType = gotPoint
        objIndex = Index
    Else
        Index = Index - .PointCount
        If Index <= .FigureCount Then
            ObjType = gotFigure
            objIndex = Index
        Else
            Index = Index - .FigureCount
            If Index <= .LabelCount Then
                ObjType = gotLabel
                objIndex = Index - .PointCount - .LabelCount
            Else
                Index = Index - .LabelCount
                If Index <= .LocusCount Then
                    ObjType = gotLocus
                    objIndex = Index - .PointCount - .LabelCount - .LocusCount
                Else
                    Index = Index - .LocusCount
                    If Index <= .SGCount Then
                        ObjType = gotSG
                        objIndex = Index - .PointCount - .LabelCount - .LocusCount - .SGCount
                    Else
                        Index = Index - .SGCount
                        If Index <= .WECount Then
                            ObjType = gotWE
                            objIndex = Index - .PointCount - .LabelCount - .LocusCount - .SGCount - .WECount
                        Else
                            Index = Index - .SGCount
                            If Index <= .ButtonCount Then
                                ObjType = gotButton
                                objIndex = Index - .PointCount - .LabelCount - .LocusCount - .SGCount - .WECount - .ButtonCount
                            Else
                                ObjType = gotGeneric
                                objIndex = Index
                            End If ' ButtonCount
                        End If ' WECount
                    End If ' SGCount
                End If ' LocusCount
            End If ' LabelCount
        End If ' FigureCount
    End If ' PointCount
End With
End Sub

Public Sub ObjectListShowHideAll(objList As ObjectList, Optional ByVal bShow As Boolean = True, Optional ByVal ShouldRepaint As Boolean = True)
Dim Z As Long
bShow = Not bShow

With objList
    For Z = 1 To .PointCount
        BasePoint(.Points(Z)).Hide = bShow
    Next
    For Z = 1 To .FigureCount
        Figures(.Figures(Z)).Hide = bShow
    Next
    For Z = 1 To .LabelCount
        TextLabels(.Labels(Z)).Hide = bShow
    Next
    For Z = 1 To .LocusCount
        Locuses(.Loci(Z)).Hide = bShow
    Next
    For Z = 1 To .SGCount
        StaticGraphics(.SGs(Z)).Hide = bShow
    Next
    For Z = 1 To .ButtonCount
        Buttons(.Buttons(Z)).Hide = bShow
    Next
End With

If ShouldRepaint Then
    PaperCls
    ShowAll
End If
End Sub

Public Function GetObjectsFromPoint(objList As ObjectList, ByVal X As Double, ByVal Y As Double, Optional ByVal IncludeHidden As Boolean = False, Optional ByVal ShouldClearList As Boolean = False)
Dim Z As Long, ValidDistance As Double
Dim OX As Double, OY As Double, hRgn As Long, P As OnePoint
Dim lpRect As RECT

OX = X
OY = Y
ToPhysical OX, OY

'==============================================

If ShouldClearList Then ObjectListClear objList

'==============================================

With objList
    '==============================================
    '   Catch Points
    '==============================================
    
    If .PointCountMax >= 0 And PointCount > 0 Then
        For Z = 1 To PointCount
            If IsInBasePoint(X, Y, Z) Then
                ObjectListAdd objList, gotPoint, Z
                '.PointCount = .PointCount + 1
                'ReDim Preserve .Points(1 To .PointCount)
                '.Points(.PointCount) = Z
            End If
        Next Z
    End If
    
    '==============================================
    '   Catch Figures
    '==============================================
    
    If .FigureCountMax >= 0 And FigureCount > 0 Then
        For Z = 0 To FigureCount - 1
            If PointBelongsToFigure(X, Y, Z, IncludeHidden) Then
                ObjectListAdd objList, gotFigure, Z
                '.FigureCount = .FigureCount + 1
                'ReDim Preserve .Figures(1 To .FigureCount)
                '.Figures(.FigureCount) = Z
            End If
        Next Z
    End If
    
    '==============================================
    '   Catch Static graphics
    '==============================================
    
    If .SGCountMax >= 0 And StaticGraphicCount > 0 Then
        For Z = 1 To StaticGraphicCount
            If StaticGraphics(Z).Visible Then
                Select Case StaticGraphics(Z).Type
                Case sgPolygon
                    hRgn = CreatePolygonRgn(StaticGraphics(Z).ObjectPixels(1), StaticGraphics(Z).NumberOfPoints, ALTERNATE)
                    If PtInRegion(hRgn, OX, OY) <> 0 Then
                        ObjectListAdd objList, gotSG, Z
                        '.SGCount = .SGCount + 1
                        'ReDim Preserve .SGs(1 To .SGCount)
                        '.SGs(.SGCount) = Z
                    End If
                    DeleteObject hRgn
                    
                Case sgBezier
                    If PtInBezier(Z, X, Y) Then
                        ObjectListAdd objList, gotSG, Z
                        '.SGCount = .SGCount + 1
                        'ReDim Preserve .SGs(1 To .SGCount)
                        '.SGs(.SGCount) = Z
                    End If
                
                Case sgVector
                    ValidDistance = 1
                    ToLogicalLength ValidDistance
                    ValidDistance = ValidDistance + Sensitivity
                    
                    If PointBelongsToSegment(X, Y, BasePoint(StaticGraphics(Z).Points(1)).X, BasePoint(StaticGraphics(Z).Points(1)).Y, BasePoint(StaticGraphics(Z).Points(2)).X, BasePoint(StaticGraphics(Z).Points(2)).Y, ValidDistance) Then
                        ObjectListAdd objList, gotSG, Z
                        '.SGCount = .SGCount + 1
                        'ReDim Preserve .SGs(1 To .SGCount)
                        '.SGs(.SGCount) = Z
                    End If
                End Select
            End If
        Next
    End If
    
    '==============================================
    '   Catch Labels
    '==============================================
    
    If .LabelCountMax >= 0 And LabelCount > 0 Then
        For Z = 1 To LabelCount
            If PointInRectangle(X, Y, TextLabels(Z).LogicalPosition.P1.X, TextLabels(Z).LogicalPosition.P1.Y, TextLabels(Z).LogicalPosition.P2.X, TextLabels(Z).LogicalPosition.P2.Y) And (Not TextLabels(Z).Hide Or IncludeHidden) Then
                ObjectListAdd objList, gotLabel, Z
                '.LabelCount = .LabelCount + 1
                'ReDim Preserve .Labels(1 To .LabelCount)
                '.Labels(.LabelCount) = Z
            End If
        Next
    End If
    
    '==============================================
    '   Catch Loci
    '==============================================
    
    If .LocusCountMax >= 0 And LocusCount > 0 Then
        For Z = 1 To LocusCount
            P = GetPerpPointPolyline(X, Y, Locuses(Z).LocusPoints)
            
            ValidDistance = Locuses(Z).DrawWidth
            ToLogicalLength ValidDistance
            ValidDistance = ValidDistance + Sensitivity
            
            If Distance(X, Y, P.X, P.Y) < ValidDistance And (Not Locuses(Z).Hide Or IncludeHidden) Then
                ObjectListAdd objList, gotLocus, Z
                '.LocusCount = .LocusCount + 1
                'ReDim Preserve .Loci(1 To .LocusCount)
                '.Loci(.LocusCount) = Z
            End If
        Next
    End If

    '==============================================
    '   Catch Buttons
    '==============================================
    
    If .ButtonCountMax >= 0 And ButtonCount > 0 Then
        For Z = 1 To ButtonCount
            lpRect = Buttons(Z).Position
            lpRect.Left = Buttons(Z).Position.Left - FrameWidth - 2
            lpRect.Top = Buttons(Z).Position.Top - FrameHeight
            lpRect.Right = Buttons(Z).Position.Left + Buttons(Z).Position.Right + FrameWidth + 1
            lpRect.Bottom = Buttons(Z).Position.Top + Buttons(Z).Position.Bottom + FrameHeight + 1
            If PointInRectangle(OX, OY, lpRect.Left, lpRect.Top, lpRect.Right, lpRect.Bottom) Then
                ObjectListAdd objList, gotButton, Z
                '.ButtonCount = .ButtonCount + 1
                'ReDim Preserve .Buttons(1 To .ButtonCount)
                '.Buttons(.ButtonCount) = Z
            End If
        Next
    End If
    
End With

End Function

Public Function ObjectListGetUpperPoint(objList As ObjectList) As Long
Dim Z As Long, MaxZOrder As Long

If objList.PointCount = 0 Then ObjectListGetUpperPoint = 0: Exit Function

MaxZOrder = 1

For Z = 2 To objList.PointCount
    If BasePoint(objList.Points(Z)).ZOrder > BasePoint(objList.Points(MaxZOrder)).ZOrder Then MaxZOrder = Z
Next

If MaxZOrder > 0 Then ObjectListGetUpperPoint = objList.Points(MaxZOrder) Else ObjectListGetUpperPoint = 0
End Function

Public Function ObjectListGetUpperFigure(objList As ObjectList) As Long
Dim Z As Long, MaxZOrder As Long

If objList.FigureCount = 0 Then ObjectListGetUpperFigure = -1: Exit Function

MaxZOrder = 1

For Z = 2 To objList.FigureCount
    If Figures(objList.Figures(Z)).ZOrder > Figures(objList.Figures(MaxZOrder)).ZOrder Then MaxZOrder = Z
Next

If MaxZOrder > 0 Then ObjectListGetUpperFigure = objList.Figures(MaxZOrder) Else ObjectListGetUpperFigure = -1
End Function

Public Function ObjectListGetUpperSG(objList As ObjectList) As Long
If objList.SGCount > 0 Then ObjectListGetUpperSG = objList.SGs(objList.SGCount)
End Function

Public Function ObjectListGetUpperLocus(objList As ObjectList) As Long
If objList.LocusCount > 0 Then ObjectListGetUpperLocus = objList.Loci(objList.LocusCount)
End Function

Public Function ObjectListGetUpperLabel(objList As ObjectList) As Long
If objList.LabelCount > 0 Then ObjectListGetUpperLabel = objList.Labels(objList.LabelCount)
End Function

Public Function ObjectListGetUpperButton(objList As ObjectList) As Long
If objList.ButtonCount > 0 Then ObjectListGetUpperButton = objList.Buttons(objList.ButtonCount)
End Function
