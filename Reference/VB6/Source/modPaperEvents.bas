Attribute VB_Name = "modInterfaceIn"
'========================================================
'                               modInterfaceIn
'Module to link concrete environment events to abstract kernel
'and make kernel do things in response to mouse or keyboard events
'========================================================

Option Explicit

Public Enum MenuStates 'Main window menu-enabled states
    mnsStandard
    mnsFigureCreate
    mnsObjectSelection
    mnsDemo
    mnsFullScreen
    mnsCompleteDisable
    mnsMacroGivens
    mnsMacroResults
    mnsMacroRun
End Enum

Public Const MaxDragNumbers = 5
Public Enum DragStateConstants
    dscNormalState
    dscDraggingState
    dscMovingState
    dscErrorState
    dscPointLabelDrag
    dscMacroStateGivens
    dscMacroStateResults
    dscMacroStateRun
    dscCreateStaticGraphic
    dscScroll
    dscSelectObjects
    dscPushingState
    dscMeasureDrag
    dscDemo
End Enum

Public Type DragState
    State As DragStateConstants
    WhatDoIDrag As GeometryObjectType
    Button As Integer
    Shift As Integer
    OX As Double
    OY As Double
    X As Double
    Y As Double
    NumOfMouseDowns As Integer
    NumOfMouseUps As Integer
    Points() As Long
    Pixels() As POINTAPI
    NumberOfPoints As Long
    TypeOfStaticGraphic As StaticGraphicType
    Figures() As Long
    NumberOfFigures As Long
    Number(1 To MaxDragNumbers) As Integer
    MacroObjects() As Long
    MacroObjectType() As DrawState
    MacroObjectCount As Long
    MacroCurrentObject As Long
    MacroObjectDescription() As String
    OldAutoRedraw As Boolean
    ShouldComplete As Boolean
    ShouldSkipUndo As Boolean
    ShouldHideFirstPointName As Boolean
End Type

Public DragS As DragState
Public tempMacro As Macro
Public WasKeyDown As Boolean

Public RestrictPaperResize As Boolean

'################################################################
'Process double click on a virtual paper
'Mainly: opens object properties on double click
'################################################################

Public Sub PaperDoubleClick()
Dim AP As Long, AF As Long, AL As Long, ASG As Long
Dim locX As Double, locY As Double

locX = DragS.OX
locY = DragS.OY

If DrawingState = dsSelect Then
    If DragS.State = dscScroll Then
        DragS.State = dscNormalState
    End If
End If

If DragS.State = dscCreateStaticGraphic Then
    If DragS.TypeOfStaticGraphic = sgPolygon Then
        If DragS.NumberOfPoints >= 3 Then
            BasePoint(DragS.Points(1)).ShowName = Not DragS.ShouldHideFirstPointName
            DragS.ShouldHideFirstPointName = False
            
            DragS.NumberOfPoints = DragS.NumberOfPoints + 1
            ReDim Preserve DragS.Points(1 To DragS.NumberOfPoints)
            DragS.Points(DragS.NumberOfPoints) = DragS.Points(1)
            
            AddStaticGraphic sgPolygon, DragS.Points
            i_CheckMainbarButton GetString(ResStaticObjectBase + 2 * sgPolygon), False
            
            CancelOperation
            Exit Sub
        End If
    End If
End If

If DrawingState = dsPolygon And DragS.State <> dscNormalState Then
    If DragS.NumberOfPoints >= 3 Then
        BasePoint(DragS.Points(1)).ShowName = Not DragS.ShouldHideFirstPointName
        DragS.ShouldHideFirstPointName = False
        
        DragS.NumberOfPoints = DragS.NumberOfPoints + 1
        ReDim Preserve DragS.Points(1 To DragS.NumberOfPoints)
        DragS.Points(DragS.NumberOfPoints) = DragS.Points(1)
        
        AddStaticGraphic sgPolygon, DragS.Points
        
        CancelOperation
        Exit Sub
    End If
End If

'If not dscNormalState and Selection then exit sub
If DragS.State <> dscNormalState Or DrawingState <> dsSelect Then Exit Sub

'Get objects from cursor to open its properties window
AP = GetPointFromCursor(locX, locY)
AF = GetFigureByPoint(locX, locY)
AL = GetLabelFromPoint(locX, locY)
ASG = GetSGFromPoint(locX, locY)

If IsPoint(AP) Then
    ActivePoint = AP
    frmPointProps.Show
    Exit Sub
End If

If IsFigure(AF) Then
    ActiveFigure = AF
    If Figures(AF).FigureType = dsMeasureDistance Or Figures(AF).FigureType = dsMeasureAngle Then
        MenuCommand ResMnuMeasurementProperties
    Else
        frmFigureProps.Show
    End If
    Exit Sub
End If

If IsLabel(AL) Then
    ActiveLabel = AL
    frmLabelProps.Show
    Exit Sub
End If

If ASG > 0 Then
    ActiveStatic = ASG
    frmStaticProps.Show
    Exit Sub
End If

End Sub

'################################################################
'Process key press on a virtual paper
'Mainly: scroll, Ctrl+A, Ctrl+T, Enter, Escape
'################################################################

Public Sub PaperKeyDown(KeyCode As Integer, Shift As Integer)
'===================================================
Dim CurTool As Long
'===================================================

'Debug section
'#If conDebug = 1 Then
'#End If

'===================================================

If DragS.State <> dscDemo Then
    Select Case KeyCode
        Case vbKeyUp
            ScrollUp Shift And vbCtrlMask
            
        Case vbKeyDown
            ScrollDown Shift And vbCtrlMask
        
        Case vbKeyLeft
            ScrollLeft Shift And vbCtrlMask
        
        Case vbKeyRight
            ScrollRight Shift And vbCtrlMask
        
        Case vbKeyPageUp
            ScrollUp True
        
        Case vbKeyPageDown
            ScrollDown True
        
        Case vbKeyHome
            ScrollLeft True
        
        Case vbKeyEnd
            ScrollRight True
        
        Case vbKeyAdd
            ZoomIn
            
        Case vbKeySubtract
            ZoomOut
            
        Case vbKeyH
            If Shift = 0 Then ScrollHome
    End Select
End If

'===================================================

If DragS.State = dscDemo Then
    If KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight Or KeyCode = vbKeyPageDown Then DemoNextStep
    If KeyCode = vbKeyBack Or KeyCode = vbKeyUp Or KeyCode = vbKeyLeft Or KeyCode = vbKeyPageUp Then DemoPreviousStep
    If KeyCode = vbKeyHome Then DemoFirstStep
    If KeyCode = vbKeyEnd Then DemoLastStep
    Exit Sub
End If

'===================================================

If DragS.State = dscSelectObjects Then
    If TempObjectSelection.PointCountMax = 0 Then
        If KeyCode = vbKeyReturn Then ObjectSelectionComplete
    End If
    Exit Sub
End If

'===================================================

If DragS.State = dscMacroStateGivens Then
    If KeyCode = vbKeyReturn Then MacroCreateResultsInit
    Exit Sub
End If

'===================================================

If DragS.State = dscMacroStateResults Then
    If KeyCode = vbKeyReturn Then MacroSaveInit
    Exit Sub
End If

'===================================================

Select Case KeyCode
    Case vbKeyA To vbKeyZ
        If Shift = 0 Then ShortcutKeyClick KeyCode
        
        'toggle ToolSelectOnce
        If KeyCode = vbKeyT And DragS.State = dscNormalState Then
            If (Shift And vbCtrlMask) Then
                setToolSelectOnce = Not setToolSelectOnce
                MsgBox GetString(ResSetSelectToolOnce) & " - " & Format(setToolSelectOnce, "on/off"), vbInformation
            End If
        End If
        
        'toggle AutoShowPointName
        If KeyCode = vbKeyA And DragS.State = dscNormalState Then
            If Shift And vbCtrlMask Then
                setAutoShowPointName = Not setAutoShowPointName
                MsgBox GetString(ResAutoShowPointName) & " - " & Format(setAutoShowPointName, "on/off"), vbInformation
            End If
        End If
    
    Case vbKeySpace
        '===================================================
        '====================TEST AREA======================
        '===================================================
        
        Dim T As Double, Z As Long
        
        T = timeGetTime
'        For Z = 1 To 1000
'            ShowAll
'            UpdateWindow Paper.hWnd
'        Next
'
'        MsgBox timeGetTime - t

'
        'For Z = 1 To 10000
        '    P = GetPerpPoint2(3, 4, 1, 0, 0, 1)
        'Next
        'MsgBox timeGetTime - T
        'BitBlt GetDC(Paper.hWnd), 0, 0, Paper.ScaleWidth - 1, Paper.ScaleHeight - 1, Paper.hDC, 0, 0, SRCCOPY
        
        '===================================================
        '===================================================
        
    Case vbKeyReturn
        If DragS.State = dscNormalState Then AddLabel
        
    Case vbKeyTab
        If DragS.State = dscNormalState Then
            CurTool = TransposeInv(DrawingState, MenuTransposition)
            If Shift = 0 Then
                CurTool = CurTool + 1
                If CurTool > MenuTransposition.Count Then CurTool = 1
            ElseIf Shift = 1 Then
                CurTool = CurTool - 1
                If CurTool < 1 Then CurTool = MenuTransposition.Count
            End If
            i_SelectTool Transpose(CurTool, MenuTransposition)
            ImitateMouseMove
        End If
        
    Case vbKeyControl, vbKeyShift, vbKeyMenu
        If WasKeyDown = False And DragS.State <> dscDraggingState And DragS.State <> dscScroll And DragS.State <> dscMovingState Then
            ImitateMouseMove
            WasKeyDown = True
        End If
End Select
End Sub

'################################################################
'Process key release on a virtual paper
'Imitate mouse move to update cursor and corresponding information
'################################################################

Public Sub PaperKeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyControl, vbKeyShift, vbKeyMenu
        If WasKeyDown And DragS.State <> dscDraggingState And DragS.State <> dscScroll And DragS.State <> dscMovingState Then
            ImitateMouseMove
            WasKeyDown = False
        End If
End Select
End Sub

'################################################################
'Process virtual paper mousedown event
'################################################################

Public Sub PaperMouseDown(Button As Integer, Shift As Integer, OX As Single, OY As Single)
'======================================================

On Local Error Resume Next

'======================================================

Dim P As OnePoint, lpPoint As POINTAPI
Dim P1 As Long, P2 As Long, P3 As Long, P4 As Long
Dim R As TwoPoints, Rad As Double, bFlagCount As Long
Dim XC As Double, YC As Double, Radius As Double

Dim objList As ObjectList
Dim AP As Long, AF As Long, AL As Long, ASG As Long, ALoc As Long, AB As Long, ASGs() As Long
Dim APoints() As Long, AFigures() As Long, MaxZOrder As Long

Dim pAction As Action
Dim X As Double, Y As Double, Z As Long, SGC As Long

'======================================================

X = OX
Y = OY
ToLogical X, Y

DragS.OX = X
DragS.OY = Y

'======================================================
'                           DEMO
'======================================================

If DragS.State = dscDemo Then
    If Button = 1 Then
        DemoNextStep
    ElseIf Button = 2 Then
        DemoPreviousStep
    End If
    Exit Sub
End If

'======================================================

If Shift = 1 And (DragS.State = dscDraggingState Or DragS.State = dscMovingState Or DragS.State = dscNormalState) Then
    X = Round(X)
    Y = Round(Y)
End If

GetObjectsFromPoint objList, X, Y
AP = ObjectListGetUpperPoint(objList)
AF = ObjectListGetUpperFigure(objList)
ASG = ObjectListGetUpperSG(objList)
AB = ObjectListGetUpperButton(objList)
AL = ObjectListGetUpperLabel(objList)
ALoc = ObjectListGetUpperLocus(objList)

'=========================================================
'                                       Select objects
'=========================================================

If DragS.State = dscSelectObjects Then
    If IsPoint(AP) Then
        ObjectListAddRemove TempObjectSelection, gotPoint, AP
        If TempObjectSelection.Type = ostCalcPoints Then
            If TempObjectSelection.PointCountMax > 0 Then
                If TempObjectSelection.PointCount >= TempObjectSelection.PointCountMax Then
                    ObjectSelectionComplete
                    Exit Sub
                End If
            End If
        End If
    ElseIf IsFigure(AF) Then
        If TempObjectSelection.Type = ostShowHideObjects Then
            ObjectListAddRemove TempObjectSelection, gotFigure, AF
        ElseIf TempObjectSelection.Type = ostCalcPoints And TempObjectSelection.PointCountMax = 2 And (Figures(AF).FigureType = dsSegment Or Figures(AF).FigureType = dsMeasureDistance) Then
            If TempObjectSelection.PointCount = 1 Then ObjectListDelete TempObjectSelection, gotPoint, 1
            ObjectListAdd TempObjectSelection, gotPoint, Figures(AF).Points(0)
            ObjectListAdd TempObjectSelection, gotPoint, Figures(AF).Points(1)
            ObjectSelectionComplete
            Exit Sub
        ElseIf TempObjectSelection.Type = ostCalcPoints And TempObjectSelection.PointCountMax = 3 And Figures(AF).FigureType = dsMeasureAngle Then
            Do While TempObjectSelection.PointCount > 0
                ObjectListDelete TempObjectSelection, gotPoint, 1
            Loop
            ObjectListAdd TempObjectSelection, gotPoint, Figures(AF).Points(0)
            ObjectListAdd TempObjectSelection, gotPoint, Figures(AF).Points(1)
            ObjectListAdd TempObjectSelection, gotPoint, Figures(AF).Points(2)
            ObjectSelectionComplete
            Exit Sub
        End If
    ElseIf IsSG(ASG) Then
        If TempObjectSelection.Type = ostShowHideObjects Then
            ObjectListAddRemove TempObjectSelection, gotSG, ASG
        End If
    ElseIf IsLabel(AL) Then
        If TempObjectSelection.Type = ostShowHideObjects Then
            ObjectListAddRemove TempObjectSelection, gotLabel, AL
        End If
    ElseIf IsLocus(ALoc) Then
        If TempObjectSelection.Type = ostShowHideObjects Then
            ObjectListAddRemove TempObjectSelection, gotLocus, ALoc
        End If
    ElseIf IsButton(AB) Then
        If TempObjectSelection.Type = ostShowHideObjects Then
            If TempObjectSelection.SubType = oscButton Then
                If (AB < ActiveButton) Or (ActiveButton = 0) Then ObjectListAddRemove TempObjectSelection, gotButton, AB
            Else
                ObjectListAddRemove TempObjectSelection, gotButton, AB
            End If
        End If
    Else
    End If
    PaperCls
    ShowSelectedAll TempObjectSelection
    Exit Sub
End If

'=========================================================
'                                       Move point label
'=========================================================

If Not IsPoint(AP) And Not IsLabel(AL) And Not IsButton(AB) And Shift = 0 And Button = 1 And DragS.State = dscNormalState And DrawingState = dsSelect Then
    Z = GetPointLabelFromCursor(X, Y)
    If IsPoint(Z) Then
        DragS.State = dscPointLabelDrag
        DragS.Number(1) = Z
        DragS.Number(2) = Maximum(BasePoint(Z).LabelWidth, BasePoint(Z).LabelHeight) + BasePoint(Z).PhysicalWidth + setCursorSensitivity
        DragS.Number(3) = BasePoint(DragS.Number(1)).PhysicalWidth  '+ BasePoint(DragS.Number(1)).LabelHeight \ 2
        DragS.OX = OX - BasePoint(DragS.Number(1)).LabelOffsetX
        DragS.OY = OY - BasePoint(DragS.Number(1)).LabelOffsetY + DragS.Number(3)
        DragS.NumOfMouseDowns = 1
        DragS.NumOfMouseUps = 0
        
        pAction.Type = actMovePointLabel
        pAction.pPoint = Z
        ReDim pAction.AuxPoints(1 To 1)
        pAction.AuxPoints(1).X = BasePoint(Z).LabelOffsetX
        pAction.AuxPoints(1).Y = BasePoint(Z).LabelOffsetY
        RecordAction pAction
        
        Exit Sub
    End If
End If

'=========================================================
'                                   Move measurement
'=========================================================

If IsFigure(AF) And Not IsPoint(AP) And Not IsLabel(AL) And Not IsButton(AB) And Shift = 0 And Button = 1 And DragS.State = dscNormalState And DrawingState = dsSelect Then
    If Figures(AF).FigureType = dsMeasureDistance Or Figures(AF).FigureType = dsMeasureAngle Then
        If Figures(AF).FigureType = dsMeasureDistance Or (Figures(AF).FigureType = dsMeasureAngle And IsPointInMeasureText(AF, X, Y)) Then
            DragS.State = dscMeasureDrag
            DragS.Number(1) = AF
            'DragS.Number(2) = BasePoint(Z).LabelWidth * 2 + BasePoint(Z).PhysicalWidth
            'DragS.Number(3) = BasePoint(DragS.Number(1)).PhysicalWidth  '+ BasePoint(DragS.Number(1)).LabelHeight \ 2
            
            DragS.X = OX - Figures(DragS.Number(1)).AuxPoints(6).X
            DragS.Y = OY - Figures(DragS.Number(1)).AuxPoints(6).Y
            
            DragS.OX = X
            DragS.OY = Y
            
            DragS.NumOfMouseDowns = 1
            DragS.NumOfMouseUps = 0
    '        pAction.Type = actMovePointLabel '// ????? Add new action type to the Undo buffer procedures
    '        pAction.pPoint = Z
    '        ReDim pAction.AuxPoints(1 To 1)
    '        pAction.AuxPoints(1).X = BasePoint(Z).LabelOffsetX
    '        pAction.AuxPoints(1).Y = BasePoint(Z).LabelOffsetY
    '        RecordAction pAction
            Exit Sub
        End If
    End If
End If

'=========================================================
'                                           Scroll paper
'=========================================================

If (Not IsPoint(AP)) And (Not IsLabel(AL)) And (Not IsButton(AB)) And (DragS.State = dscNormalState Or DragS.State = dscMovingState) And (Not (IsFigure(AF) And (DrawingState = dsPoint))) And (Button = vbLeftButton) And (CBool((Shift And vbCtrlMask) = vbCtrlMask) Or (DrawingState = dsSelect)) Then
    DragS.OX = OX
    DragS.OY = OY
    DragS.State = dscScroll
    Exit Sub
End If

'=========================================================
'Now process object right clicks and menu popups...
'=========================================================

If (Button <> 1 Or Shift <> 0) And DragS.State = dscNormalState Then
    If Button = 2 And Shift = 0 Then
        '================================================
        If IsButton(AB) Then 'begin dragging a button to a new location
        
            DragS.X = X
            DragS.Y = Y
            DragS.NumOfMouseDowns = 1
            DragS.NumOfMouseUps = 0
            DragS.WhatDoIDrag = gotButton
            DragS.State = dscDraggingState
            
            DragS.Number(1) = AB
            Exit Sub
        End If
        
        '================================================
        
        If IsPoint(AP) Then
            APoints = GetPointsFromCursor(X, Y)
            AFigures = GetFiguresByPoint(X, Y, AP)
            i_PointRightClicked AP, APoints, AFigures
            Exit Sub
        End If
        
        '================================================
        
        If IsFigure(AF) Then
            AFigures = GetFiguresByPoint(X, Y)
            i_FigureRightClicked AF, AFigures
            Exit Sub
        End If
        
        '================================================
        
        If IsLabel(AL) Then
            i_LabelRightClicked AL
            Exit Sub
        End If
        
        '================================================
        
        If ASG > 0 Then
            i_SGRightClicked ASG
            Exit Sub
        End If
    End If
    
    '================================================
    
    If Button = 2 And Not IsPoint(AP) And Not IsFigure(AF) And Not IsLabel(AL) Then
        If DrawingState = dsSelect Then Exit Sub
        i_SelectTool dsSelect
        i_SetMousePointer curStateArrow
        i_ShowStatus GetString(ResSelect)
        Exit Sub
    End If
    
    If Button <> 1 Then CancelOperation: Exit Sub
End If

'=========================================================
'                                   Run macro
'=========================================================

If DragS.State = dscMacroStateRun Then
    If DragS.MacroCurrentObject > DragS.MacroObjectCount Then Exit Sub
    
    If DragS.MacroObjectType(DragS.MacroCurrentObject) = dsPoint Then
        
        If IsPoint(AP) Then
            DragS.MacroObjects(DragS.MacroCurrentObject) = AP
            ShowSelectedPoint Paper.hDC, AP, True
        Else
            AddBasePoint X, Y, , , , , False
            DragS.MacroObjects(DragS.MacroCurrentObject) = PointCount
            ShowSelectedPoint Paper.hDC, PointCount, True
        End If
        
        DragS.MacroCurrentObject = DragS.MacroCurrentObject + 1
        
    Else
        
        If IsFigure(AF) Then
            If IsThereATypeMatch(Figures(AF).FigureType, DragS.MacroObjectType(DragS.MacroCurrentObject)) Then
                DragS.MacroObjects(DragS.MacroCurrentObject) = AF
                ShowSelectedFigure Paper.hDC, AF, True
                DragS.MacroCurrentObject = DragS.MacroCurrentObject + 1
            End If
        End If
        
    End If
    
    Exit Sub
End If

'=========================================================
'                                       Select macro givens
'=========================================================

If DragS.State = dscMacroStateGivens Then
    'APoints = GetPointsFromCursor(X, Y)
    'AFigures = GetFiguresByPoint(X, Y, , False)
    'i_MacroGivensClick AP, AF, APoints, AFigures
    i_MacroGivensClick objList
    Exit Sub
End If

'=========================================================
'                                       Select macro results
'=========================================================

If DragS.State = dscMacroStateResults Then
    i_MacroResultsClick objList
    Exit Sub
End If

'=========================================================
'                                       Create static graphic
'=========================================================

If DragS.State = dscCreateStaticGraphic Then
    Select Case DragS.TypeOfStaticGraphic
        Case sgBezier
        Case sgPolygon
        Case sgVector
    End Select
    Exit Sub
End If

'======================================================
'======================================================
'======================================================

With DragS
'    If .NumOfMouseUps = 0 And .State = dscNormalState Then
'        .OldAutoRedraw = Paper.AutoRedraw
'    End If
    If .NumOfMouseUps > .NumOfMouseDowns Then
        .State = dscNormalState
    End If
    
    Select Case DragS.State
        Case dscNormalState
            .NumOfMouseDowns = 1
            .NumOfMouseUps = 0
        Case dscMovingState
            .NumOfMouseDowns = .NumOfMouseDowns + 1
        Case dscErrorState
            .NumOfMouseDowns = 1
            .NumOfMouseUps = 0
    End Select
    
    .State = dscDraggingState
    .Button = Button
    .Shift = Shift
    .OX = X
    .OY = Y
End With

'======================================================
'   Disable all menus
'======================================================

i_StartOperation

'======================================================

Select Case DrawingState
'##############################
'SELECT
Case dsSelect
    If IsButton(AB) Then
        If Button = 1 And Shift = 0 Then
            DragS.Number(1) = AB
            DragS.State = dscPushingState
            DragS.WhatDoIDrag = gotButton
            Buttons(AB).Pushed = True
            ShowButton Paper.hDC, AB, True, True, True
            'ShowAll
            Exit Sub
        End If
    End If
    
    If IsPoint(AP) Then
        If BasePoint(AP).Type = dsPoint Or BasePoint(AP).Type = dsPointOnFigure Then
            DragS.Number(1) = AP
            DragS.WhatDoIDrag = gotPoint
            'Paper.AutoRedraw = True
            pAction.Type = actMovePoint
            ReDim pAction.AuxPoints(1 To 1)
            pAction.AuxPoints(1).X = BasePoint(AP).X
            pAction.AuxPoints(1).Y = BasePoint(AP).Y
            pAction.pPoint = AP
            ReDim pAction.AuxInfo(1 To 4)
            If BasePoint(AP).Type = dsPointOnFigure Then
                pAction.AuxInfo(4) = 1
                pAction.AuxInfo(3) = Figures(BasePoint(AP).ParentFigure).AuxInfo(1)
                pAction.pFigure = BasePoint(AP).ParentFigure
            End If
            If BasePoint(AP).Locus <> 0 Then
                If Locuses(BasePoint(AP).Locus).Enabled Then
                    pAction.AuxInfo(1) = Locuses(BasePoint(AP).Locus).LocusPointCount
                    pAction.AuxInfo(2) = 1
                End If
            End If
            RecordAction pAction
            
        Else ' cannot drag a point
            DragS.State = dscErrorState
            i_ShowStatus GetString(ResMsgCantDragThisPoint) & "."
            i_SetMousePointer curStateNo
            Exit Sub
        End If
        
        If BasePoint(AP).Type = dsPointOnFigure Then ManualDragFlag = AP
        i_ShowStatus GetString(ResShiftSnapToGrid) & "."
        Exit Sub
    End If ' IsPoint(AP)
    
    If IsLabel(AL) Then
        If Not TextLabels(AL).Fixed Then
            DragS.Number(1) = AL
            DragS.WhatDoIDrag = gotLabel
            DragS.State = dscDraggingState
            DragS.X = DragS.OX
            DragS.Y = DragS.OY
            Exit Sub
        End If
    End If

'##############################
'POINT
Case dsPoint
'    If IsPoint(AP) Then
'        If BasePoint(AP).Type = dsPoint Or BasePoint(AP).Type = dsPointOnFigure Then DragS.Number(1) = AP Else DragS.State = dscErrorState: i_SetMousePointer curStateNo
'        If BasePoint(AP).Type = dsPointOnFigure Then ManualDragFlag = AP
'        i_ShowStatus
'    End If

'##############################
'SEGMENT
Case dsSegment
    If Button = 1 Then
        If DragS.NumOfMouseUps = 0 Then
            If Not IsPoint(AP) Then
                'Paper.AutoRedraw = True
                AddBasePoint X, Y
                'Paper.AutoRedraw = False
                DragS.Number(1) = PointCount
            Else
                DragS.Number(1) = AP
            End If
        End If
    End If

'##############################
'RAY
Case dsRay
    If Button = 1 Then
        If DragS.NumOfMouseUps = 0 Then
            If Not IsPoint(AP) Then
                'Paper.AutoRedraw = True
                AddBasePoint X, Y
                'Paper.AutoRedraw = False
                DragS.Number(1) = PointCount
            Else
                DragS.Number(1) = AP
            End If
        End If
    End If

'##############################
'LINE_2POINTS

Case dsLine_2Points
    If Button = 1 Then
        If DragS.NumOfMouseUps = 0 Then
            If Not IsPoint(AP) Then
                'Paper.AutoRedraw = True
                AddBasePoint X, Y
                'Paper.AutoRedraw = False
                DragS.Number(1) = PointCount
            Else
                DragS.Number(1) = AP
            End If
        End If
    End If

'##############################
'BISECTOR
Case dsBisector
    If Button = 1 Then
        If DragS.NumOfMouseUps = 0 Then
            If Not IsPoint(AP) Then
                AddBasePoint X, Y
                DragS.Number(1) = PointCount
                ShowSelectedPoint Paper.hDC, PointCount
                Paper.Refresh
            Else
                DragS.Number(1) = AP
                ShowSelectedPoint Paper.hDC, AP
                Paper.Refresh
            End If
        End If
    End If

'##############################
'LINE_PARALLEL

Case dsLine_PointAndParallelLine
    If Button = 1 And DragS.NumOfMouseUps = 0 Then
        If Not IsLine(AF) Then
            If IsPoint(AP) Then
                DragS.Number(1) = -1
                DragS.Number(2) = AP
                ShowSelectedPoint Paper.hDC, AP
                Exit Sub
            End If
            DragS.NumOfMouseDowns = 0
            DragS.NumOfMouseUps = 0
            DragS.State = dscErrorState
            i_SetMousePointer curStateNo
            Exit Sub
        End If
        DragS.Number(1) = AF
        R = GetLineCoordinates(AF)
        R = GetLineFromSegment(X, Y, X + (R.P1.X - R.P2.X), Y + (R.P1.Y - R.P2.Y))
        DrawLine R.P1.X, R.P1.Y, R.P2.X, R.P2.Y, , vbInvert
   End If
       
'################################
'LINE_PERPENDICULAR
Case dsLine_PointAndPerpendicularLine
     If Button = 1 And DragS.NumOfMouseUps = 0 Then
        If Not IsLine(AF) Then
           If IsPoint(AP) Then
               DragS.Number(1) = -1
               DragS.Number(2) = AP
               ShowSelectedPoint Paper.hDC, AP
               Paper.Refresh
               Exit Sub
           End If
           DragS.NumOfMouseDowns = 0
           DragS.NumOfMouseUps = 0
           DragS.State = dscErrorState
           i_SetMousePointer curStateNo
           Exit Sub
        End If
        DragS.Number(1) = AF
        R = GetLineCoordinates(AF)
        R = GetPerpendicularLine(X, Y, R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
        DrawLine R.P1.X, R.P1.Y, R.P2.X, R.P2.Y, , vbInvert
        Paper.Refresh
    
    End If
'###############################
'CIRCLE
Case dsCircle_CenterAndCircumPoint
    If Button = 1 Then
        If DragS.NumOfMouseUps = 0 Then
            If Not IsPoint(AP) Then
                'Paper.AutoRedraw = True
                AddBasePoint X, Y
                'Paper.AutoRedraw = False
                DragS.Number(1) = PointCount
            Else
                DragS.Number(1) = AP
            End If
            R.P1.X = BasePoint(DragS.Number(1)).X
            R.P1.Y = BasePoint(DragS.Number(1)).Y
            R.P2.X = DragS.OX
            R.P2.Y = DragS.OY
            Rad = Distance(R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
            Radius = Rad
            ToPhysicalLength Radius
            XC = R.P1.X
            YC = R.P1.Y
            ToPhysical XC, YC
            DrawCircle Paper.hDC, XC, YC, Radius, , , , vbInvert
            Paper.Refresh
        End If
    End If

'##############################
'CIRCLE BY RADIUS

Case dsCircle_CenterAndTwoPoints
    If Button = 1 Then
        If DragS.NumOfMouseUps = 0 Then
            If Not IsPoint(AP) Then
                'Paper.AutoRedraw = True
                AddBasePoint X, Y
                'Paper.AutoRedraw = False
                DragS.Number(1) = PointCount
            Else
                DragS.Number(1) = AP
            End If
            DrawLine BasePoint(DragS.Number(1)).X, BasePoint(DragS.Number(1)).Y, X, Y, , vbInvert
       End If
    End If

'##############################
'ARC

Case dsCircle_ArcCenterAndRadiusAndTwoPoints
    If Button = 1 Then
        If DragS.NumOfMouseUps = 0 Then
            If Not IsPoint(AP) Then
                'Paper.AutoRedraw = True
                AddBasePoint X, Y
                'Paper.AutoRedraw = False
                DragS.Number(1) = PointCount
            Else
                DragS.Number(1) = AP
            End If
            DrawLine BasePoint(DragS.Number(1)).X, BasePoint(DragS.Number(1)).Y, X, Y, , vbInvert
        End If
    End If
    
'##############################
'MIDDLE POINT

Case dsMiddlePoint
    If Button = 1 Then
        If DragS.NumOfMouseUps = 0 Then
            If Not IsPoint(AP) Then
                If IsFigure(AF) Then
                    If Figures(AF).FigureType = dsSegment Then
                        DragS.Number(1) = AF
                        DragS.Number(2) = -1
                        Exit Sub
                    End If
                End If
                'Paper.AutoRedraw = True
                AddBasePoint X, Y
                'Paper.AutoRedraw = False
                DragS.Number(1) = PointCount
            Else
                DragS.Number(1) = AP
            End If
        End If
    End If
    
'##############################
'SIMM POINT

Case dsSimmPoint
    If Button = 1 Then
        If DragS.NumOfMouseUps = 0 Then
            If Not IsPoint(AP) Then
                'Paper.AutoRedraw = True
                AddBasePoint X, Y
                'Paper.AutoRedraw = False
                DragS.Number(1) = PointCount
            Else
                DragS.Number(1) = AP
            End If
        End If
    End If

'##############################
'SIMM POINT BY LINE

Case dsSimmPointByLine
    If Button = 1 Then
        If DragS.NumOfMouseUps = 0 Then
            If Not IsPoint(AP) Then
                If IsLine(AF) Then
                    DragS.Number(2) = AF
                    ShowSelectedFigure Paper.hDC, AF
                    Paper.Refresh
                    Exit Sub
                End If
                'Paper.AutoRedraw = True
                AddBasePoint X, Y
                'Paper.AutoRedraw = False
                DragS.Number(1) = PointCount
            Else
                DragS.Number(1) = AP
            End If
            ShowSelectedPoint Paper.hDC, DragS.Number(1)
            Paper.Refresh
        End If
    End If

'##############################
'INVERTED POINT

Case dsInvert
    If Button = 1 Then
        If DragS.NumOfMouseUps = 0 Then
            If Not IsPoint(AP) Then
                If IsCircle(AF) Then
                    DragS.Number(2) = AF
                    ShowSelectedFigure Paper.hDC, AF
                    Paper.Refresh
                    Exit Sub
                End If
                'Paper.AutoRedraw = True
                AddBasePoint X, Y
                'Paper.AutoRedraw = False
                DragS.Number(1) = PointCount
            Else
                DragS.Number(1) = AP
            End If
            ShowSelectedPoint Paper.hDC, DragS.Number(1)
            Paper.Refresh
        End If
    End If
    
'##############################
'INTERSECT
Case dsIntersect
    If Button = 1 And DragS.NumOfMouseUps = 0 Then
        If Not IsLine(AF) And Not IsCircle(AF) Then
            DragS.NumOfMouseDowns = 0
            DragS.NumOfMouseUps = 0
            DragS.State = dscErrorState
            i_SetMousePointer curStateNo
            Exit Sub
        Else
            Dim AFS() As Long
            
            AFS = GetFiguresByPoint(X, Y, , False)
            If LBound(AFS) + 1 = UBound(AFS) Then
                AddIntersectionPoints AFS(LBound(AFS)), AFS(UBound(AFS))
                ShowAll
                DragS.ShouldComplete = True
            Else
                DragS.Number(1) = AF
            End If
        End If
   End If

'###############################
'POINT ON FIGURE
Case dsPointOnFigure
    If Button = 1 And DragS.NumOfMouseUps = 0 Then
        If IsLine(AF) Or IsCircle(AF) Then
            If Not IsPoint(AP) Then
                'Paper.AutoRedraw = True
                AddPointOnFigure AF, X, Y
                'Paper.AutoRedraw = False
            End If
        End If
        If LocusCount > 0 Then
            ALoc = GetLocusFromPoint(X, Y)
            If ALoc <> 0 Then
                If Locuses(ALoc).Dynamic Then
                    AF = -1
                    For Z = 0 To FigureCount - 1
                        If Figures(Z).FigureType = dsDynamicLocus Then
                            If Figures(Z).Points(0) = Locuses(ALoc).ParentPoint Then AF = Z
                        End If
                    Next
                    'Paper.AutoRedraw = True
                    If IsFigure(AF) Then AddPointOnFigure AF, X, Y
                    'Paper.AutoRedraw = False
                End If
            End If
        End If
    End If
    
'###############################
'MEASURE DISTANCE
Case dsMeasureDistance
    If Button = 1 Then
        If DragS.NumOfMouseUps = 0 Then
            If Not IsPoint(AP) Then
                If IsFigure(AF) Then
                    If Figures(AF).FigureType = dsSegment Then
                        AddMeasureDistance Figures(AF).Points(0), Figures(AF).Points(1)
                        'If Paper.AutoRedraw <> DragS.OldAutoRedraw Then Paper.AutoRedraw = DragS.OldAutoRedraw
                        DragS.ShouldComplete = True
                        ShowAll
                        Exit Sub
                    End If
                End If
                DragS.NumOfMouseDowns = 0
                DragS.NumOfMouseUps = 0
                DragS.State = dscErrorState
                i_SetMousePointer curStateNo
                Exit Sub
            Else
                DragS.Number(1) = AP
                ShowSelectedPoint Paper.hDC, AP
                Paper.Refresh
            End If
        End If
    End If


'###############################
'MEASURE ANGLE
Case dsMeasureAngle
    If Button = 1 Then
        If DragS.NumOfMouseUps = 0 Then
            If Not IsPoint(AP) Then
                DragS.NumOfMouseDowns = 0
                DragS.NumOfMouseUps = 0
                DragS.State = dscErrorState
                i_SetMousePointer curStateNo
                Exit Sub
            Else
                DragS.Number(1) = AP
                ShowSelectedPoint Paper.hDC, AP
                Paper.Refresh
            End If
        End If
    End If

'##############################
'DYNAMIC LOCUS
Case dsDynamicLocus
    If Button = 1 Then
        If DragS.NumOfMouseUps = 0 Then
            If Not IsPoint(AP) Then
                DragS.NumOfMouseDowns = 0
                DragS.NumOfMouseUps = 0
                DragS.State = dscErrorState
                i_SetMousePointer curStateNo
                Exit Sub
            Else
                If BasePoint(AP).Type = dsPoint Then
                    DragS.NumOfMouseDowns = 0
                    DragS.NumOfMouseUps = 0
                    DragS.State = dscErrorState
                    i_SetMousePointer curStateNo
                    Exit Sub
                End If
                DragS.Number(1) = AP
                ShowSelectedPoint Paper.hDC, AP
                Paper.Refresh
            End If
        End If
    End If
    
'=========================================================
'                                           MEASURE AREA
'=========================================================
    
Case dsMeasureArea
    If Button = 1 And DragS.NumOfMouseUps = 0 Then
        If IsSG(ASG) And Not IsPoint(AP) Then
            If StaticGraphics(ASG).Type = sgPolygon Then
                AddMeasureArea StaticGraphics(ASG).Points
                DragS.ShouldComplete = True
                ShowAll
                Exit Sub
            End If
        End If
        If IsCircle(AF) And Not IsPoint(AP) Then
            Dim P5(1 To 2) As Long
            Select Case Figures(AF).FigureType
            Case dsCircle_CenterAndCircumPoint
                P5(1) = Figures(AF).Points(0)
                P5(2) = Figures(AF).Points(1)
                AddMeasureArea P5
            Case dsCircle_CenterAndTwoPoints
                P5(1) = Figures(AF).Points(1)
                P5(2) = Figures(AF).Points(2)
                AddMeasureArea P5
            Case dsCircle_ArcCenterAndRadiusAndTwoPoints
                P5(1) = Figures(AF).Points(1)
                P5(2) = Figures(AF).Points(2)
                AddMeasureArea P5
            Case dsAnCircle
                AddTextLabel GetString(ResArea) & " = " & Format(PI * GetCircleRadius(AF) ^ 2, setFormatDistance), GetCircleCenter(AF).X, GetCircleCenter(AF).Y
            End Select
            DragS.ShouldComplete = True
            ShowAll
            Exit Sub
        End If
    End If
    
'====================================================
'          POLYGON
'====================================================

Case dsPolygon
    
End Select
End Sub

'################################################################
'Process virtual paper MOUSEMOVE event; called from Canvas_MouseMove
'################################################################

Public Sub PaperMouseMove(Button As Integer, Shift As Integer, OX As Single, OY As Single)
'======================================================

On Local Error Resume Next

'======================================================

Dim P As OnePoint
Dim R As TwoPoints
Dim Z As Long
Dim R2 As TwoPoints
Dim XC As Double, YC As Double, Radius As Double
Dim B As Boolean, Ang1 As Double, Ang2 As Double
Dim tCount As Long

Dim TC As CursorState
Dim X As Double, Y As Double, tX As Double, tY As Double
Dim Rad As Double

Dim objList As ObjectList
Dim AP As Long, AF As Long, ASG As Long, ALoc As Long, AL As Long, AB As Long, APL As Long
Dim APIsAPoint As Boolean, AFIsAFigure As Boolean
Dim AFS() As Long

Dim X1 As Double, X2 As Double, Y1 As Double, Y2 As Double
Dim bPointAlreadySelected As Boolean

'=========================================================

If DragS.State = dscErrorState Then Exit Sub

'=========================================================
'                                               DEMO
'======================================================

If DragS.State = dscDemo Then
    Exit Sub
End If

'=========================================================
'                                   POINT LABEL DRAG
'======================================================

If DragS.State = dscPointLabelDrag Then
    'ShowPoint Paper.hDC, DragS.Number(1), True
    PaperCls
    tX = OX - DragS.OX
    tY = OY - DragS.OY
    If Sqr(tX ^ 2 + tY ^ 2) > DragS.Number(2) Then
        NormalizeVector tX, tY
        tX = tX * DragS.Number(2)
        tY = tY * DragS.Number(2)
    End If
    BasePoint(DragS.Number(1)).LabelOffsetX = tX
    BasePoint(DragS.Number(1)).LabelOffsetY = tY + DragS.Number(3)
    ShowAll
    Exit Sub
End If

'=========================================================
'                                   MEASURE DRAG
'======================================================

If DragS.State = dscMeasureDrag Then
    PaperCls
    tX = OX - DragS.X
    tY = OY - DragS.Y
    
    Rad = AngleMarkDist
    ToPhysicalLength Rad
    Rad = Rad * 2
    If Sqr(tX ^ 2 + tY ^ 2) > Rad And Figures(DragS.Number(1)).FigureType = dsMeasureAngle Then
        NormalizeVector tX, tY
        tX = tX * Rad
        tY = tY * Rad
    End If
    
    If Figures(DragS.Number(1)).FigureType = dsMeasureDistance Then
        tX = tX + Figures(DragS.Number(1)).AuxPoints(3).X
        tY = tY + Figures(DragS.Number(1)).AuxPoints(3).Y
        
        X1 = BasePoint(Figures(DragS.Number(1)).Points(0)).X
        Y1 = BasePoint(Figures(DragS.Number(1)).Points(0)).Y
        X2 = BasePoint(Figures(DragS.Number(1)).Points(1)).X
        Y2 = BasePoint(Figures(DragS.Number(1)).Points(1)).Y
        ToPhysical X1, Y1
        ToPhysical X2, Y2
        
        LinkPointToSegment tX, tY, X1, Y1, X2, Y2, Rad
        
        tX = tX - Figures(DragS.Number(1)).AuxPoints(3).X
        tY = tY - Figures(DragS.Number(1)).AuxPoints(3).Y
    End If
    
    Figures(DragS.Number(1)).AuxPoints(6).X = tX
    Figures(DragS.Number(1)).AuxPoints(6).Y = tY
    ShowAll
    Exit Sub
End If

'=========================================================

X = OX
Y = OY
ToLogical X, Y

'=========================================================
'                           SCROLL
'=========================================================

If DragS.State = dscScroll Then
    tX = OX - DragS.OX
    tY = OY - DragS.OY
    DragS.OX = OX
    DragS.OY = OY
    ScrollMouse tX, tY
    Exit Sub
End If

'=========================================================

If Shift = 1 And (DragS.State = dscDraggingState Or DragS.State = dscMovingState) Then
    X = Round(X)
    Y = Round(Y)
End If

i_MouseMoved X, Y 'notify concrete environment of a paper mouse move

'=========================================================

'Dim D As Long
'D = timeGetTime
GetObjectsFromPoint objList, X, Y
'FormMain.Caption = timeGetTime - D & " - " & objList.SGCount

AP = ObjectListGetUpperPoint(objList)
AF = ObjectListGetUpperFigure(objList)
ASG = ObjectListGetUpperSG(objList)
AB = ObjectListGetUpperButton(objList)
AL = ObjectListGetUpperLabel(objList)
ALoc = ObjectListGetUpperLocus(objList)
APIsAPoint = IsPoint(AP)
AFIsAFigure = IsFigure(AF)

'If StaticGraphicCount > 0 Then
'    Dim IsIn As Boolean
'    Dim N As Integer
'    Dim i As Integer
'
'    IsIn = False
'
'    With StaticGraphics(1)
'        N = .NumberOfPoints
'        .ObjectPixels(0).X = .ObjectPixels(N).X
'        .ObjectPixels(0).Y = .ObjectPixels(N).Y
'        For i = 0 To N - 1
'            If OY > .ObjectPixels(i).Y Eqv OY <= .ObjectPixels(i + 1).Y Then
'                If (OX - .ObjectPixels(i).X < (OY - .ObjectPixels(i).Y) * (.ObjectPixels(i + 1).X - .ObjectPixels(i).X) / (.ObjectPixels(i + 1).Y - .ObjectPixels(i).Y)) Then IsIn = Not IsIn
'            End If
'        Next
'    End With
'    FormMain.Caption = IsIn
'
'End If

'IsIn:=False;
'x[0]:=x[n];
'y[0]:=y[n];
'for i:=0 to n-1 do
'  begin
'    if not((y>y[i])xor(y<=y[i+1]))
'      then
'        begin
'          if (x-x[i]<(y-y[i])*(x[i+1]-x[i])/(y[i+1]-y[i]))
'            then
'              IsIn:=not(IsIn)
'        end;
'  end;


'=========================================================
'                                   SELECT OBJECTS
'======================================================

If DragS.State = dscSelectObjects Then
    If Button = 0 Then
        If IsPoint(AP) Then
            i_ShowStatus GetString(ResSelectObjects) & ": " & BasePoint(AP).Name
            If ObjectListFind(TempObjectSelection, gotPoint, AP) = 0 Then i_SetMousePointer curStateAdd Else i_SetMousePointer curStateRemove
        ElseIf IsFigure(AF) Then
            If TempObjectSelection.Type = ostShowHideObjects Then
                i_ShowStatus GetString(ResSelectObjects) & ": " & Figures(AF).Name
                If ObjectListFind(TempObjectSelection, gotFigure, AF) = 0 Then i_SetMousePointer curStateAdd Else i_SetMousePointer curStateRemove
            ElseIf TempObjectSelection.Type = ostCalcPoints Then
                If TempObjectSelection.PointCountMax = 2 And (Figures(AF).FigureType = dsSegment Or Figures(AF).FigureType = dsMeasureDistance) Then
                    i_ShowStatus GetString(ResSelectObjects) & ": " & BasePoint(Figures(AF).Points(0)).Name & "; " & BasePoint(Figures(AF).Points(1)).Name
                    i_SetMousePointer curStateAdd
                ElseIf TempObjectSelection.PointCountMax = 3 And Figures(AF).FigureType = dsMeasureAngle Then
                    i_ShowStatus GetString(ResSelectObjects) & ": " & BasePoint(Figures(AF).Points(0)).Name & "; " & BasePoint(Figures(AF).Points(1)).Name & "; " & BasePoint(Figures(AF).Points(2)).Name
                    i_SetMousePointer curStateAdd
                Else
                    i_ShowStatus GetString(ResLocate) & " " & GetString(ResPoint)
                    i_SetMousePointer curStateArrow
                End If
            Else
                i_ShowStatus GetString(ResSelectObjects)
                i_SetMousePointer curStateArrow
            End If
        ElseIf IsSG(ASG) Then
            If TempObjectSelection.Type = ostShowHideObjects Then
                i_ShowStatus GetString(ResSelectObjects) & ": " & GetString(ResStaticObjectBase + 2 * StaticGraphics(ASG).Type) & ASG
                If ObjectListFind(TempObjectSelection, gotSG, ASG) = 0 Then i_SetMousePointer curStateAdd Else i_SetMousePointer curStateRemove
            ElseIf TempObjectSelection.Type = ostCalcPoints Then
                i_ShowStatus GetString(ResLocate) & " " & GetString(ResPoint)
                i_SetMousePointer curStateArrow
            Else
                i_ShowStatus GetString(ResSelectObjects)
                i_SetMousePointer curStateArrow
            End If
        ElseIf IsLabel(AL) Then
            If TempObjectSelection.Type = ostShowHideObjects Then
                i_ShowStatus GetString(ResSelectObjects) & ": " & GetString(ResLabel) & AL
                If ObjectListFind(TempObjectSelection, gotLabel, AL) = 0 Then i_SetMousePointer curStateAdd Else i_SetMousePointer curStateRemove
            ElseIf TempObjectSelection.Type = ostCalcPoints Then
                i_ShowStatus GetString(ResLocate) & " " & GetString(ResPoint)
                i_SetMousePointer curStateArrow
            Else
                i_ShowStatus GetString(ResSelectObjects)
                i_SetMousePointer curStateArrow
            End If
        ElseIf IsLocus(ALoc) Then
            If TempObjectSelection.Type = ostShowHideObjects Then
                i_ShowStatus GetString(ResSelectObjects) & ": " & GetString(ResLocus) & ALoc
                If ObjectListFind(TempObjectSelection, gotLocus, ALoc) = 0 Then i_SetMousePointer curStateAdd Else i_SetMousePointer curStateRemove
            ElseIf TempObjectSelection.Type = ostCalcPoints Then
                i_ShowStatus GetString(ResLocate) & " " & GetString(ResPoint)
                i_SetMousePointer curStateArrow
            Else
                i_ShowStatus GetString(ResSelectObjects)
                i_SetMousePointer curStateArrow
            End If
        ElseIf IsButton(AB) Then
            If TempObjectSelection.Type = ostShowHideObjects Then
                i_ShowStatus GetString(ResSelectObjects) & ": " & GetString(ResButton) & AB
                If ObjectListFind(TempObjectSelection, gotButton, AB) = 0 Then i_SetMousePointer curStateAdd Else i_SetMousePointer curStateRemove
            ElseIf TempObjectSelection.Type = ostCalcPoints Then
                i_ShowStatus GetString(ResLocate) & " " & GetString(ResPoint)
                i_SetMousePointer curStateArrow
            Else
                i_ShowStatus GetString(ResSelectObjects)
                i_SetMousePointer curStateArrow
            End If
        Else
            Select Case TempObjectSelection.Type
            Case ostShowHideObjects
                i_ShowStatus GetString(ResSelectObjects)
            Case ostCalcPoints
                i_ShowStatus GetString(ResLocate) & " " & GetString(ResPoint)
            End Select
            i_SetMousePointer curStateArrow
        End If
    End If
    Exit Sub
End If

'=========================================================
'                                   MACRO RUN
'======================================================

If DragS.State = dscMacroStateRun Then
    If DragS.MacroCurrentObject > DragS.MacroObjectCount Then Exit Sub
    If DragS.MacroObjectType(DragS.MacroCurrentObject) = dsPoint Then
        If APIsAPoint Then
            i_SetMousePointer curStateAdd
            i_ShowStatus DragS.MacroCurrentObject & ". " & DragS.MacroObjectDescription(DragS.MacroCurrentObject) & " (" & BasePoint(AP).Name & ")"
            Exit Sub
        End If
    Else
        If AFIsAFigure Then
            If IsThereATypeMatch(Figures(AF).FigureType, DragS.MacroObjectType(DragS.MacroCurrentObject)) Then
                i_SetMousePointer curStateAdd
                'i_ShowStatus GetString(ResLocateObject) & " " & Figures(AF).Name
                i_ShowStatus DragS.MacroCurrentObject & ". " & DragS.MacroObjectDescription(DragS.MacroCurrentObject) & " (" & Figures(AF).Name & ")"
                Exit Sub
            End If
        End If
    End If
    'i_ShowStatus GetString(ResLocateObject) & " " & LCase(GetString(ResFigureBase + 2 * DragS.MacroObjectType(DragS.MacroCurrentObject)))
    i_ShowStatus DragS.MacroCurrentObject & ". " & DragS.MacroObjectDescription(DragS.MacroCurrentObject)
    i_SetMousePointer curStateCross
    Exit Sub
End If

'=========================================================
'                   MACRO GIVENS
'=========================================================

If DragS.State = dscMacroStateGivens Then
    
    If objList.PointCount + objList.FigureCount > 1 And Not (objList.FigureCount > 0 And objList.PointCount = 1) Then
        tCount = objList.PointCount
        
        For Z = 1 To objList.FigureCount
            If IsVisual(objList.Figures(Z)) Then tCount = tCount + 1
        Next
        
        If tCount > 1 Then
            i_SetMousePointer curStateQuestion
            i_ShowStatus GetString(ResChoiceAmbiguity)
            Exit Sub
        End If
    End If
    
    If APIsAPoint Then
        If GivenPoints(AP) = 0 Then
            i_SetMousePointer curStateAdd
            i_ShowStatus GetString(ResFigurePoint) & " (" & BasePoint(AP).Name & ")"
            Exit Sub
        Else
            i_SetMousePointer curStateRemove
            i_ShowStatus GetString(ResFigurePoint) & " (" & BasePoint(AP).Name & ")"
            Exit Sub
        End If
    End If
    
    If AFIsAFigure Then
        If Figures(AF).FigureType <> dsMeasureDistance And Figures(AF).FigureType <> dsMeasureAngle Then
            If GivenFigures(AF) = 0 Then
                i_SetMousePointer curStateAdd
                i_ShowStatus GetString(ResFigureBase + Figures(AF).FigureType * 2) & " (" & Figures(AF).Name & ")"
                Exit Sub
            Else
                i_SetMousePointer curStateRemove
                i_ShowStatus GetString(ResFigureBase + Figures(AF).FigureType * 2) & " (" & Figures(AF).Name & ")"
                Exit Sub
            End If
        End If
    End If
    
    i_SetMousePointer curStateCross
    i_ShowStatus GetString(ResSelectGivens) & "."
    Exit Sub
End If

'=========================================================
'                               MACRO RESULTS
'=========================================================

If DragS.State = dscMacroStateResults Then
    
    If objList.FigureCount + objList.PointCount + objList.SGCount > 1 Then
        tCount = 0
        
        If objList.PointCount <> 1 Then
            For Z = 1 To objList.FigureCount
                If ResultFigures(objList.Figures(Z)) >= 0 Then tCount = tCount + 1
            Next
        End If
        
        For Z = 1 To objList.PointCount
            If ResultPoints(objList.Points(Z)) >= 0 Then tCount = tCount + 1
        Next
        
        For Z = 1 To objList.SGCount
            If ResultSGs(objList.SGs(Z)) >= 0 Then tCount = tCount + 1
        Next
        
        If tCount > 1 Then
            i_SetMousePointer curStateQuestion
            i_ShowStatus GetString(ResChoiceAmbiguity)
            Exit Sub
        End If
    End If
    
    If IsPoint(AP) Then
        If ResultPoints(AP) >= 0 Then
            If ResultFigures(BasePoint(AP).ParentFigure) >= 0 Then
                If ResultFigures(BasePoint(AP).ParentFigure) = 0 Then
                    i_SetMousePointer curStateAdd
                Else
                    i_SetMousePointer curStateRemove
                End If
                i_ShowStatus GetString(ResFigureBase + BasePoint(AP).Type * 2) & " (" & BasePoint(AP).Name & ")"
                Exit Sub
            End If
        End If
    End If
    
    If IsFigure(AF) Then
        If ResultFigures(AF) >= 0 Then
            If ResultFigures(AF) = 0 Then
                i_SetMousePointer curStateAdd
            Else
                i_SetMousePointer curStateRemove
            End If
            i_ShowStatus GetString(ResFigureBase + Figures(AF).FigureType * 2) & " (" & Figures(AF).Name & ")"
            Exit Sub
        End If
    End If
    
    If IsSG(ASG) Then
        If ResultSGs(ASG) >= 0 Then
            If ResultSGs(ASG) = 0 Then
                i_SetMousePointer curStateAdd
            Else
                i_SetMousePointer curStateRemove
            End If
            i_ShowStatus GetString(ResStaticObjectBase + 2 * StaticGraphics(ASG).Type)
            Exit Sub
        End If
    End If
    
    If IsLocus(ALoc) Then
        If ResultLoci(ALoc) >= 0 Then
            If ResultLoci(ALoc) = 0 Then
                i_SetMousePointer curStateAdd
            Else
                i_SetMousePointer curStateRemove
            End If
            i_ShowStatus GetString(ResFigureBase + dsDynamicLocus * 2)
            Exit Sub
        End If
    End If
    
    i_SetMousePointer curStateCross
    i_ShowStatus GetString(ResSelectResults)
    Exit Sub
End If

'=========================================================
'                   CREATE STATIC GRAPHIC
'=========================================================

If DragS.State = dscCreateStaticGraphic Then
    If APIsAPoint Then
        bPointAlreadySelected = False
        For Z = 2 To DragS.NumberOfPoints
            If DragS.Points(Z) = AP Then bPointAlreadySelected = True
        Next
        
        If bPointAlreadySelected Then
            TC = curStateCross
            If DragS.TypeOfStaticGraphic = sgPolygon And DragS.NumberOfPoints > 1 Then
                i_ShowStatus GetString(ResLocate) & " " & GetString(ResPoint) & ". " & Replace(GetString(ResClickToClosePolygon), "%1", BasePoint(DragS.Points(1)).Name) & " " & GetString(ResPressEsc)
            Else
                i_ShowStatus GetString(ResLocate) & " " & GetString(ResPoint) & ". " & GetString(ResPressEsc)
            End If
        Else
            TC = curStateAdd
            If DragS.TypeOfStaticGraphic = sgPolygon And DragS.NumberOfPoints > 1 Then
                If AP = DragS.Points(1) Then
                    i_ShowStatus Replace(GetString(ResClickToClosePolygon), "%1", BasePoint(DragS.Points(1)).Name) & " " & GetString(ResPressEsc)
                Else
                    i_ShowStatus GetString(ResLocate) & " " & GetString(ResPoint) & " (" & BasePoint(AP).Name & "). " & Replace(GetString(ResClickToClosePolygon), "%1", BasePoint(DragS.Points(1)).Name) & " " & GetString(ResPressEsc)
                End If
            Else
                i_ShowStatus GetString(ResLocate) & " " & GetString(ResPoint) & " (" & BasePoint(AP).Name & "). " & GetString(ResPressEsc)
            End If
        End If
    Else
        TC = curStateCross
        If DragS.TypeOfStaticGraphic = sgPolygon And DragS.NumberOfPoints > 1 Then
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResPoint) & ". " & Replace(GetString(ResClickToClosePolygon), "%1", BasePoint(DragS.Points(1)).Name) & " " & GetString(ResPressEsc)
        Else
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResPoint) & ". " & GetString(ResPressEsc)
        End If
    End If
    i_SetMousePointer TC
    
    If DragS.TypeOfStaticGraphic = sgPolygon Then
        If DragS.NumberOfPoints = 1 Then
            DrawLine BasePoint(DragS.Points(1)).X, BasePoint(DragS.Points(1)).Y, DragS.OX, DragS.OY, 0, vbInvert, False
            DrawLine BasePoint(DragS.Points(1)).X, BasePoint(DragS.Points(1)).Y, X, Y, 0, 6, False
            ShowSelectedPoint Paper.hDC, DragS.Points(1)
            Paper.Refresh
        End If
        If DragS.NumberOfPoints > 1 Then
            'ToPhysical DragS.OX, DragS.OY
            'DragS.Pixels(DragS.NumberOfPoints + 1).X = DragS.OX
            'DragS.Pixels(DragS.NumberOfPoints + 1).Y = DragS.OY
            'DrawPolygon DragS.Pixels(), setdefcolFigureFill, 7, False
            PaperCls
            ShowAll , , False
            For Z = 1 To DragS.NumberOfPoints
                ShowSelectedPoint Paper.hDC, DragS.Points(Z)
            Next
            DragS.Pixels(DragS.NumberOfPoints + 1).X = OX
            DragS.Pixels(DragS.NumberOfPoints + 1).Y = OY
            DrawPolygon DragS.Pixels(), setdefcolFigure, setdefcolFigureFill, 9, True
        End If
        DragS.OX = X
        DragS.OY = Y
    End If
    
    Exit Sub
End If

'=========================================================
'               STATUSBAR AND CURSOR
'=========================================================

If Button = 0 And (DragS.State = dscNormalState Or DragS.State = dscMovingState) Then
    If APIsAPoint Then TC = curStateAdd Else TC = curStateCross
    
    Select Case DrawingState
        Case dsSelect 'SELECT
            APL = GetPointLabelFromCursor(X, Y)
            If IsButton(AB) Then
                TC = curStateArrow
            ElseIf APIsAPoint Then
                If BasePoint(AP).Type = dsPoint Or BasePoint(AP).Type = dsPointOnFigure Then TC = curStateDrag Else TC = curStateSelect
            ElseIf AFIsAFigure Or IsLabel(GetLabelFromPoint(X, Y)) Or IsSG(ASG) Then
                TC = curStateSelect
                If AFIsAFigure Then
                    If Figures(AF).FigureType = dsMeasureAngle Or Figures(AF).FigureType = dsMeasureDistance Then
                        If (Figures(AF).FigureType = dsMeasureAngle And IsPointInMeasureText(AF, X, Y)) Or Figures(AF).FigureType = dsMeasureDistance Then TC = curStateDrag
                    End If
                End If
            ElseIf IsPoint(APL) Then
                TC = curStateSelect
            Else
                TC = curStateArrow
            End If
        
        Case dsPoint 'POINT
            AFS = GetFiguresByPoint(X, Y, , False)
            If LBound(AFS) + 1 = UBound(AFS) Then
                TC = curStateSelect
            Else
                If AFIsAFigure Then
                    TC = curStateAdd
                Else
                    TC = curStateCross
                End If
            End If
        
        Case dsSegment, dsRay, dsLine_2Points, dsBisector, dsSimmPoint, dsCircle_CenterAndCircumPoint
            If DragS.NumOfMouseUps = 1 And AP = DragS.Number(1) Then TC = curStateCross
        
        Case dsMiddlePoint
            If DragS.NumOfMouseUps = 1 And AP = DragS.Number(1) Then TC = curStateCross
            If DragS.State = dscNormalState And IsFigure(AF) Then
                If Figures(AF).FigureType = dsSegment Then
                    TC = curStateAdd
                End If
            End If
        
        Case dsLine_PointAndParallelLine, dsLine_PointAndPerpendicularLine
            If DragS.NumOfMouseUps = 0 Then
                If IsLine(AF) Or IsPoint(AP) Then TC = curStateAdd Else TC = curStateCross
            End If
            If DragS.NumOfMouseUps > 0 And DragS.Number(1) = -1 Then
                If IsLine(AF) Then TC = curStateAdd Else TC = curStateCross
            End If
        
        Case dsIntersect
            AFS = GetFiguresByPoint(X, Y, , False)
            If LBound(AFS) + 1 = UBound(AFS) Then
                TC = curStateSelect
            Else
                If AFIsAFigure Then TC = curStateAdd Else TC = curStateCross
            End If
        
        Case dsPointOnFigure
            ALoc = GetLocusFromPoint(X, Y)
            If ALoc > 0 Then If Not Locuses(ALoc).Dynamic Then ALoc = 0
            If AFIsAFigure Or ALoc > 0 Then TC = curStateAdd Else TC = curStateCross
        
        Case dsCircle_CenterAndTwoPoints, dsCircle_ArcCenterAndRadiusAndTwoPoints
            If DragS.NumOfMouseUps = 1 And AP = DragS.Number(1) Then TC = curStateCross
        
        Case dsSimmPointByLine
            If DragS.NumOfMouseUps = 1 Then
                If DragS.Number(1) > 0 Then
                    If IsLine(AF) Then TC = curStateAdd Else TC = curStateCross
                Else
                    If IsPoint(AP) Then TC = curStateAdd Else TC = curStateCross
                End If
            End If
            If DragS.NumOfMouseUps = 0 Then
                If IsLine(AF) Or IsPoint(AP) Then TC = curStateAdd Else TC = curStateCross
            End If
        
        Case dsInvert
            If DragS.NumOfMouseUps = 1 Then
                If DragS.Number(1) > 0 Then
                    If IsCircle(AF) Then TC = curStateAdd Else TC = curStateCross
                Else
                    If IsPoint(AP) Then TC = curStateAdd Else TC = curStateCross
                End If
            End If
            If DragS.NumOfMouseUps = 0 Then
                If IsCircle(AF) Or IsPoint(AP) Then TC = curStateAdd Else TC = curStateCross
            End If
        
        Case dsMeasureDistance, dsMeasureAngle
            If IsFigure(AF) Then
                If Figures(AF).FigureType = dsSegment Then TC = curStateAdd
            End If
        
        Case dsMeasureArea
            If IsSG(ASG) And DragS.NumOfMouseUps = 0 Then
                If StaticGraphics(ASG).Type = sgPolygon Then
                    TC = curStateSelect
                End If
            End If
            If IsCircle(AF) And DragS.NumOfMouseUps = 0 Then
                TC = curStateSelect
            End If
            If IsPoint(AP) Then TC = curStateAdd
        
        Case dsDynamicLocus
            TC = curStateCross
            If IsPoint(AP) Then
                If BasePoint(AP).Type <> dsPoint Then
                    TC = curStateAdd
                End If
            End If
        
        Case dsPolygon
            
    End Select
    
    If DrawingState <> dsPolygon Then i_SetMousePointer TC
End If

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

Select Case DrawingState
'############################
'SELECT
Case dsSelect
    If DragS.State = dscNormalState Then
        AL = GetLabelFromPoint(X, Y)
        AB = GetButtonFromPoint(X, Y)
        APL = GetPointLabelFromCursor(X, Y)
        If IsButton(AB) Then
             i_ShowStatus IIf(Buttons(AB).Fixed, "", GetString(ResMoveButton) & ". ") & GetString(ResRightClickForMenu)
        ElseIf APIsAPoint Then
            If BasePoint(AP).Type = dsPoint Or BasePoint(AP).Type = dsPointOnFigure Then
                i_ShowStatus GetString(ResDragPoint) & " (" & BasePoint(AP).Name & "). " & GetString(ResRightClickForMenu) & " " & GetString(ResDoubleClickProps)
            Else
                i_ShowStatus GetString(ResMsgCantDragThisPoint) & " (" & BasePoint(AP).Name & "). " & GetString(ResRightClickForMenu) & " " & GetString(ResDoubleClickProps)
            End If
        ElseIf IsLabel(AL) Then
            i_ShowStatus IIf(TextLabels(AL).Fixed, "", GetString(ResMoveLabel) & ". ") & GetString(ResRightClickForMenu) & " " & GetString(ResDoubleClickProps)
        ElseIf IsFigure(AF) Then
            If Figures(AF).FigureType = dsMeasureAngle Or Figures(AF).FigureType = dsMeasureDistance Then
                If IsPointInMeasureText(AF, X, Y) Then
                    i_ShowStatus GetString(ResMoveMeasurement) & " " & GetString(ResRightClickForMenu)
                Else
                    i_ShowStatus GetString(ResRightClickForMenu)
                End If
            Else
                i_ShowStatus Figures(AF).Name & ". " & GetString(ResRightClickForMenu) & " " & GetString(ResDoubleClickProps)
            End If
        ElseIf IsPoint(APL) Then
            i_ShowStatus GetString(ResMovePointName)
        ElseIf ASG <> 0 Then
            i_ShowStatus GetString(ResRightClickForMenu) & " " & GetString(ResDoubleClickProps)
        Else
            i_ShowStatus GetString(ResSelect) & "."
        End If
    End If
    If DragS.State = dscDraggingState And DragS.Number(1) <> 0 Then
        '==========================================
        '                   Drag a button
        '==========================================
        If DragS.WhatDoIDrag = gotButton Then
            If Not Buttons(DragS.Number(1)).Fixed Then
                PaperCls
                MoveButton DragS.Number(1), Buttons(DragS.Number(1)).LogicalPosition.P1.X + X - DragS.OX, Buttons(DragS.Number(1)).LogicalPosition.P1.Y + Y - DragS.OY
                'ShowButton Paper.hDC, DragS.Number(1)
                ShowAll
                DragS.OX = X
                DragS.OY = Y
            End If
        '==========================================
        '                   Drag a point
        '==========================================
        ElseIf DragS.WhatDoIDrag = gotPoint Then
            If BasePoint(DragS.Number(1)).Type = dsPoint Then
                tX = X
                tY = Y
            ElseIf BasePoint(DragS.Number(1)).Type = dsPointOnFigure Then
                P = GetPointOnFigure(Figures(BasePoint(DragS.Number(1)).ParentFigure).Parents(0), X, Y)
                tX = P.X
                tY = P.Y
                If P.X <> EmptyVar And P.Y <> EmptyVar Then
                    If BasePoint(DragS.Number(1)).Visible = False Then BasePoint(DragS.Number(1)).Visible = True
                    RecalcSemiDependentInfo BasePoint(DragS.Number(1)).ParentFigure, X, Y
                End If
            End If
            
'            Static t0 As LARGE_INTEGER, t1 As Variant, t2 As Variant
'            t2 = t1
'            QueryPerformanceCounter t0
'            t1 = LargeInteger(t0)
'            Debug.Print CDec(t1 - t2)
            
            MovePoint DragS.Number(1), tX, tY
            
            If Not nGradientPaper And setWallpaper = "" Then PaperCls
            'HideAll
            RecalcAllAuxInfo
            
            'If WECount > 0 Then ValueTable1.UpdateExpressions '?????WEWEWE
            If LabelCount > 0 Then UpdateLabels
            ShowAll
            
            DragS.OX = X
            DragS.OY = Y
        
        '==========================================
        '                   Drag a label
        '==========================================
        ElseIf DragS.WhatDoIDrag = gotLabel Then
            PaperCls
            MoveLabel DragS.Number(1), TextLabels(DragS.Number(1)).LogicalPosition.P1.X + X - DragS.OX, TextLabels(DragS.Number(1)).LogicalPosition.P1.Y + Y - DragS.OY
            ShowAll
            DragS.OX = X
            DragS.OY = Y
        End If
    End If

'====================================================
'POINT
'====================================================

Case dsPoint
    AFS = GetFiguresByPoint(X, Y, , False)
    If LBound(AFS) + 1 = UBound(AFS) Then
        i_ShowStatus GetString(ResCreateIntersection) & ": " & Figures(AFS(LBound(AFS))).Name & ", " & Figures(AFS(UBound(AFS))).Name
    Else
        If AFIsAFigure Then
            i_ShowStatus GetString(ResCreatePointOnFigure) & " (" & Figures(AF).Name & ")"
        Else
            i_ShowStatus GetString(ResAddPoint) & " " & GetString(ResShiftSnapToGrid) & "."
        End If
    End If
    'If APIsAPoint Then i_ShowStatus Else i_ShowStatus GetString(ResAddPoint)

'====================================================
'SEGMENT
'====================================================

Case dsSegment
    If DragS.State = dscNormalState Then
        If IsPoint(AP) Then
            i_ShowStatus LTrim(GetString(ResLocate)) & " " & Trim(GetString(ResFirst)) & " " & Trim(GetString(ResPoint)) & ". (" & BasePoint(AP).Name & ")"
        Else
            i_ShowStatus LTrim(GetString(ResLocate)) & " " & Trim(GetString(ResFirst)) & " " & Trim(GetString(ResPoint)) & "."
        End If
    End If
    'If DragS.State = dscDraggingState And DragS.NumOfMouseUps = 0 Then i_ShowStatus GetString(ResReleaseButtonTo) & " " & LCase(GetString(ResLocate) & " " & GetString(ResFirst) & " " & GetString(ResPoint) & ".")
    If DragS.State <> dscNormalState And DragS.NumOfMouseUps = 1 And DragS.Number(1) <> 0 Then
        DrawLine BasePoint(DragS.Number(1)).X, BasePoint(DragS.Number(1)).Y, DragS.OX, DragS.OY, , vbInvert, False
        DrawLine BasePoint(DragS.Number(1)).X, BasePoint(DragS.Number(1)).Y, X, Y, , vbInvert, False
        If APIsAPoint And AP <> DragS.Number(1) Then
            i_ShowStatus GetString(ResCreateSegment) & " (" & BasePoint(DragS.Number(1)).Name & BasePoint(AP).Name & "). " & GetString(ResPressEsc)
        Else
            i_ShowStatus LTrim(GetString(ResLocate)) & " " & Trim(GetString(ResSecond)) & " " & GetString(ResPoint) & ". " & GetString(ResPressEsc)
        End If
        ShowPoint Paper.hDC, DragS.Number(1), , True
    End If
    DragS.OX = X
    DragS.OY = Y

'====================================================
'RAY
'====================================================

Case dsRay
    If DragS.State = dscNormalState Then
        If IsPoint(AP) Then
            i_ShowStatus LTrim(GetString(ResLocate)) & " " & Trim(GetString(ResFirst)) & " " & GetString(ResPoint) & ". (" & BasePoint(AP).Name & ")"
        Else
            i_ShowStatus LTrim(GetString(ResLocate)) & " " & Trim(GetString(ResFirst)) & " " & GetString(ResPoint) & "."
        End If
    End If
    'If DragS.State = dscDraggingState And DragS.NumOfMouseUps = 0 Then i_ShowStatus GetString(ResReleaseButtonTo) & " " & LCase(GetString(ResLocate) & " " & GetString(ResFirst) & " " & GetString(ResPoint) & ".")
    If DragS.State <> dscNormalState And DragS.NumOfMouseUps = 1 And DragS.Number(1) <> 0 Then
        If APIsAPoint And AP <> DragS.Number(1) Then
            i_ShowStatus GetString(ResCreateRay) & " (" & BasePoint(DragS.Number(1)).Name & BasePoint(AP).Name & "). " & GetString(ResPressEsc)
        Else
            i_ShowStatus LTrim(GetString(ResLocate)) & " " & Trim(GetString(ResSecond)) & " " & GetString(ResPoint) & ". " & GetString(ResPressEsc)
        End If
        R.P1.X = BasePoint(DragS.Number(1)).X
        R.P1.Y = BasePoint(DragS.Number(1)).Y
        R.P2.X = DragS.OX
        R.P2.Y = DragS.OY
        R2 = GetLineFromSegment(R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
        R2.P2 = R.P1
        DrawLine R2.P1.X, R2.P1.Y, R2.P2.X, R2.P2.Y, , vbInvert, False
        R.P2.X = X
        R.P2.Y = Y
        R2 = GetLineFromSegment(R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
        R2.P2 = R.P1
        DrawLine R2.P1.X, R2.P1.Y, R2.P2.X, R2.P2.Y, , vbInvert, False
        ShowPoint Paper.hDC, DragS.Number(1), , True
    End If
    DragS.OX = X
    DragS.OY = Y
    
'====================================================
'LINE_2POINTS
'====================================================

Case dsLine_2Points
    If DragS.State = dscNormalState Then
        If IsPoint(AP) Then
            i_ShowStatus Trim(GetString(ResLocate)) & " " & Trim(GetString(ResFirst)) & " " & GetString(ResPoint) & ". (" & BasePoint(AP).Name & ")"
        Else
            i_ShowStatus Trim(GetString(ResLocate)) & " " & Trim(GetString(ResFirst)) & " " & GetString(ResPoint) & "."
        End If
    End If
    If DragS.State <> dscNormalState And DragS.NumOfMouseUps = 1 And DragS.Number(1) <> 0 Then
        If APIsAPoint And AP <> DragS.Number(1) Then
            i_ShowStatus GetString(ResCreateLine) & " (" & BasePoint(DragS.Number(1)).Name & BasePoint(AP).Name & "). " & GetString(ResPressEsc)
        Else
            i_ShowStatus Trim(GetString(ResLocate)) & " " & Trim(GetString(ResSecond)) & " " & GetString(ResPoint) & ". " & GetString(ResPressEsc)
        End If
        R.P1.X = BasePoint(DragS.Number(1)).X
        R.P1.Y = BasePoint(DragS.Number(1)).Y
        R.P2.X = DragS.OX
        R.P2.Y = DragS.OY
        R2 = GetLineFromSegment(R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
        DrawLine R2.P1.X, R2.P1.Y, R2.P2.X, R2.P2.Y, , vbInvert, False
        R.P2.X = X
        R.P2.Y = Y
        R2 = GetLineFromSegment(R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
        DrawLine R2.P1.X, R2.P1.Y, R2.P2.X, R2.P2.Y, , vbInvert, False
        ShowPoint Paper.hDC, DragS.Number(1), , True
    End If
    DragS.OX = X
    DragS.OY = Y

'====================================================
'BISECTOR
'====================================================

Case dsBisector
    If DragS.State = dscNormalState Then
        If IsPoint(AP) Then
            i_ShowStatus Trim(GetString(ResLocate)) & " " & Trim(GetString(ResFirst)) & " " & GetString(ResPoint) & ". (" & BasePoint(AP).Name & ")"
        Else
            i_ShowStatus Trim(GetString(ResLocate)) & " " & Trim(GetString(ResFirst)) & " " & GetString(ResPoint) & "."
        End If
    Else
        If DragS.NumOfMouseUps = 1 Then
            If IsPoint(AP) Then
                i_ShowStatus GetString(ResLocateAngleVertex) & " (" & BasePoint(AP).Name & ") " & GetString(ResPressEsc)
            Else
                i_ShowStatus GetString(ResLocateAngleVertex) & " " & GetString(ResPressEsc)
            End If
        ElseIf DragS.NumOfMouseUps = 2 Then
            If IsPoint(AP) Then
                i_ShowStatus GetString(ResLocate) & " " & GetString(ResSecond) & " " & GetString(ResPoint) & ". (" & BasePoint(AP).Name & ") " & GetString(ResPressEsc)
            Else
                i_ShowStatus GetString(ResLocate) & " " & GetString(ResSecond) & " " & GetString(ResPoint) & ". " & GetString(ResPressEsc)
            End If
            
            R.P1.X = BasePoint(DragS.Number(2)).X
            R.P1.Y = BasePoint(DragS.Number(2)).Y
            R.P2 = GetBisector(BasePoint(DragS.Number(1)).X, BasePoint(DragS.Number(1)).Y, BasePoint(DragS.Number(2)).X, BasePoint(DragS.Number(2)).Y, DragS.OX, DragS.OY)
            R2 = GetLineFromSegment(R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
            DrawLine R2.P1.X, R2.P1.Y, R2.P2.X, R2.P2.Y, , vbInvert, False
            R.P2 = GetBisector(BasePoint(DragS.Number(1)).X, BasePoint(DragS.Number(1)).Y, BasePoint(DragS.Number(2)).X, BasePoint(DragS.Number(2)).Y, X, Y)
            R2 = GetLineFromSegment(R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
            DrawLine R2.P1.X, R2.P1.Y, R2.P2.X, R2.P2.Y, , vbInvert, False
            ShowPoint Paper.hDC, DragS.Number(1), , True
        End If
    End If
    
    DragS.OX = X
    DragS.OY = Y

'====================================================
'LINE_PARALLEL
'====================================================

Case dsLine_PointAndParallelLine
    If DragS.State = dscNormalState Then
        If IsLine(AF) Then
            i_ShowStatus Trim(GetString(ResLocate)) & " " & GetString(ResLineRayOrSegment) & " (" & Figures(AF).Name & ")"
        ElseIf IsPoint(AP) Then
            i_ShowStatus Trim(GetString(ResLocate)) & " " & Trim(GetString(ResPoint)) & " (" & BasePoint(AP).Name & ")"
        Else
            i_ShowStatus Trim(GetString(ResLocate)) & " " & Trim(GetString(ResPoint)) & ", " & GetString(ResLineRayOrSegment)
        End If
    End If
    If DragS.State <> dscNormalState And IsLine(DragS.Number(1)) Then
        If APIsAPoint Then
            i_ShowStatus GetString(ResFasten) & " " & GetString(ResPressEsc)
        Else
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResPoint) & ". " & GetString(ResPressEsc)
        End If
        R = GetLineCoordinates(DragS.Number(1))
        R = GetLineFromSegment(DragS.OX, DragS.OY, DragS.OX + (R.P1.X - R.P2.X), DragS.OY + (R.P1.Y - R.P2.Y))
        DrawLine R.P1.X, R.P1.Y, R.P2.X, R.P2.Y, , vbInvert, False
        DragS.OX = X
        DragS.OY = Y
        R = GetLineCoordinates(DragS.Number(1))
        R = GetLineFromSegment(X, Y, X + (R.P1.X - R.P2.X), Y + (R.P1.Y - R.P2.Y))
        DrawLine R.P1.X, R.P1.Y, R.P2.X, R.P2.Y, , vbInvert
    End If
    If DragS.State <> dscNormalState And DragS.Number(1) = -1 Then
        If IsLine(AF) Then
            i_ShowStatus GetString(ResCreateParLine) & " " & GetString(ResPressEsc)
        Else
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResLineRayOrSegment) & " " & GetString(ResPressEsc)
        End If
    End If
    
'====================================================
'LINE_PERPENDICULAR
'====================================================

Case dsLine_PointAndPerpendicularLine
    If DragS.State = dscNormalState Then
        If IsLine(AF) Then
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResLineRayOrSegment) & " (" & Figures(AF).Name & ")"
        ElseIf IsPoint(AP) Then
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResPoint) & ". (" & BasePoint(AP).Name & ")"
        Else
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResPoint) & ", " & GetString(ResLineRayOrSegment)
        End If
    End If
    If DragS.State <> dscNormalState And IsLine(DragS.Number(1)) Then
        If APIsAPoint Then
            i_ShowStatus GetString(ResFasten) & " " & GetString(ResPressEsc)
        Else
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResPoint) & ". " & GetString(ResPressEsc)
        End If
        R = GetLineCoordinates(DragS.Number(1))
        R = GetPerpendicularLine(DragS.OX, DragS.OY, R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
        DrawLine R.P1.X, R.P1.Y, R.P2.X, R.P2.Y, , vbInvert, False
        DragS.OX = X
        DragS.OY = Y
        R = GetLineCoordinates(DragS.Number(1))
        R = GetPerpendicularLine(X, Y, R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
        DrawLine R.P1.X, R.P1.Y, R.P2.X, R.P2.Y, , vbInvert
    End If
    If DragS.State <> dscNormalState And DragS.Number(1) = -1 Then
        If IsLine(AF) Then
            i_ShowStatus GetString(ResCreatePerpLine) & " " & GetString(ResPressEsc)
        Else
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResLineRayOrSegment) & " " & GetString(ResPressEsc)
        End If
    End If
    
'====================================================
'CIRCLE
'====================================================

Case dsCircle_CenterAndCircumPoint
    If DragS.State = dscNormalState Then
        If IsPoint(AP) Then
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResCircleCenter) & ". (" & BasePoint(AP).Name & ")"
        Else
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResCircleCenter) & "."
        End If
    End If
    If DragS.State <> dscNormalState And DragS.NumOfMouseUps <= 1 And DragS.Number(1) <> 0 Then
        If APIsAPoint And AP <> DragS.Number(1) Then
            i_ShowStatus GetString(ResCreateCircle) & ". " & GetString(ResPressEsc)
        Else
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResRadiusEnding) & ". " & GetString(ResPressEsc)
        End If
        R.P1.X = BasePoint(DragS.Number(1)).X
        R.P1.Y = BasePoint(DragS.Number(1)).Y
        R.P2.X = DragS.OX
        R.P2.Y = DragS.OY
        Rad = Distance(R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
        Radius = Rad
        ToPhysicalLength Radius
        XC = R.P1.X
        YC = R.P1.Y
        ToPhysical XC, YC
        DrawCircle Paper.hDC, XC, YC, Radius, , , , vbInvert
        R.P1.X = BasePoint(DragS.Number(1)).X
        R.P1.Y = BasePoint(DragS.Number(1)).Y
        R.P2.X = X
        R.P2.Y = Y
        Rad = Distance(R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
        Radius = Rad
        ToPhysicalLength Radius
        XC = R.P1.X
        YC = R.P1.Y
        ToPhysical XC, YC
        DrawCircle Paper.hDC, XC, YC, Radius, , , , vbInvert
        Paper.Refresh
    End If
    DragS.OX = X
    DragS.OY = Y
        
'====================================================
'CIRCLE BY RADIUS
'====================================================

Case dsCircle_CenterAndTwoPoints
    If DragS.State = dscNormalState Then
        If IsPoint(AP) Then
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResRadiusStart) & ". (" & BasePoint(AP).Name & ")"
        Else
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResRadiusStart) & "."
        End If
    End If
    If DragS.State <> dscNormalState And DragS.NumOfMouseUps <= 1 And DragS.Number(1) <> 0 Then
        DrawLine BasePoint(DragS.Number(1)).X, BasePoint(DragS.Number(1)).Y, DragS.OX, DragS.OY, , vbInvert, False
        DrawLine BasePoint(DragS.Number(1)).X, BasePoint(DragS.Number(1)).Y, X, Y, , vbInvert, False
        ShowPoint Paper.hDC, DragS.Number(1), , True
        If IsPoint(AP) Then
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResRadiusEnding) & " (" & BasePoint(AP).Name & "). " & GetString(ResPressEsc)
        Else
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResRadiusEnding) & ". " & GetString(ResPressEsc)
        End If
    End If
    If DragS.State <> dscNormalState And DragS.NumOfMouseUps > 1 And DragS.Number(2) <> 0 Then
        If APIsAPoint Then
            i_ShowStatus GetString(ResCreateCircle) & ". " & GetString(ResPressEsc)
        Else
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResCircleCenter) & "." & " " & GetString(ResPressEsc)
        End If
        R.P1.X = BasePoint(DragS.Number(1)).X
        R.P1.Y = BasePoint(DragS.Number(1)).Y
        R.P2.X = BasePoint(DragS.Number(2)).X
        R.P2.Y = BasePoint(DragS.Number(2)).Y
        Rad = Distance(R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
        Radius = Rad
        ToPhysicalLength Radius
        XC = DragS.OX
        YC = DragS.OY
        ToPhysical XC, YC
        DrawCircle Paper.hDC, XC, YC, Radius, , , , vbInvert
        R.P1.X = BasePoint(DragS.Number(1)).X
        R.P1.Y = BasePoint(DragS.Number(1)).Y
        R.P2.X = BasePoint(DragS.Number(2)).X
        R.P2.Y = BasePoint(DragS.Number(2)).Y
        Rad = Distance(R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
        Radius = Rad
        ToPhysicalLength Radius
        XC = X
        YC = Y
        ToPhysical XC, YC
        DrawCircle Paper.hDC, XC, YC, Radius, , , , vbInvert
        Paper.Refresh
    End If
    DragS.OX = X
    DragS.OY = Y
        
'====================================================
'ARC
'====================================================

Case dsCircle_ArcCenterAndRadiusAndTwoPoints
    If DragS.State = dscNormalState Then
        If IsPoint(AP) Then
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResRadiusStart) & ". (" & BasePoint(AP).Name & ")"
        Else
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResRadiusStart) & "."
        End If
    End If
    If DragS.State <> dscNormalState Then
        Select Case DragS.NumOfMouseUps
            Case 0, 1
                DrawLine BasePoint(DragS.Number(1)).X, BasePoint(DragS.Number(1)).Y, DragS.OX, DragS.OY, , vbInvert, False
                DrawLine BasePoint(DragS.Number(1)).X, BasePoint(DragS.Number(1)).Y, X, Y, , vbInvert, False
                ShowPoint Paper.hDC, DragS.Number(1), , True
                i_ShowStatus GetString(ResLocate) & " " & GetString(ResRadiusEnding) & "." & " " & GetString(ResPressEsc)
            Case 2
                i_ShowStatus GetString(ResLocate) & " " & GetString(ResCircleCenter) & "." & " " & GetString(ResPressEsc)
                R.P1.X = BasePoint(DragS.Number(1)).X
                R.P1.Y = BasePoint(DragS.Number(1)).Y
                R.P2.X = BasePoint(DragS.Number(2)).X
                R.P2.Y = BasePoint(DragS.Number(2)).Y
                Rad = Distance(R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
                Radius = Rad
                ToPhysicalLength Radius
                XC = DragS.OX
                YC = DragS.OY
                ToPhysical XC, YC
                DrawCircle Paper.hDC, XC, YC, Radius, , , , vbInvert
                XC = X
                YC = Y
                ToPhysical XC, YC
                DrawCircle Paper.hDC, XC, YC, Radius, , , , vbInvert
                Paper.Refresh
            Case 3
                i_ShowStatus GetString(ResLocate) & " " & GetString(ResArcStart) & ". " & GetString(ResPressEsc)
                R.P1.X = BasePoint(DragS.Number(1)).X
                R.P1.Y = BasePoint(DragS.Number(1)).Y
                R.P2.X = BasePoint(DragS.Number(2)).X
                R.P2.Y = BasePoint(DragS.Number(2)).Y
                Rad = Distance(R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
                R2.P1.X = BasePoint(DragS.Number(3)).X
                R2.P1.Y = BasePoint(DragS.Number(3)).Y
                Ang1 = GetAngle(R2.P1.X, R2.P1.Y, DragS.OX, DragS.OY)
                DrawLine R2.P1.X, R2.P1.Y, R2.P1.X + Rad * Cos(Ang1), R2.P1.Y - Rad * Sin(Ang1), , vbInvert, False
                Ang1 = GetAngle(R2.P1.X, R2.P1.Y, X, Y)
                DrawLine R2.P1.X, R2.P1.Y, R2.P1.X + Rad * Cos(Ang1), R2.P1.Y - Rad * Sin(Ang1), , vbInvert, False
                ShowPoint Paper.hDC, DragS.Number(3), , True
            Case 4
                If APIsAPoint Then
                    i_ShowStatus GetString(ResCreateArc) & ". " & GetString(ResPressEsc)
                Else
                    i_ShowStatus GetString(ResLocate) & " " & GetString(ResArcEnd) & ". " & GetString(ResPressEsc)
                End If
                R.P1.X = BasePoint(DragS.Number(1)).X
                R.P1.Y = BasePoint(DragS.Number(1)).Y
                R.P2.X = BasePoint(DragS.Number(2)).X
                R.P2.Y = BasePoint(DragS.Number(2)).Y
                Rad = Distance(R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
                R2.P1.X = BasePoint(DragS.Number(3)).X
                R2.P1.Y = BasePoint(DragS.Number(3)).Y
                R2.P2.X = BasePoint(DragS.Number(4)).X
                R2.P2.Y = BasePoint(DragS.Number(4)).Y
                Ang1 = GetAngle(R2.P1.X, R2.P1.Y, R2.P2.X, R2.P2.Y)
                Ang2 = GetAngle(R2.P1.X, R2.P1.Y, DragS.OX, DragS.OY)
                Radius = Rad
                ToPhysicalLength Radius
                XC = R2.P1.X
                YC = R2.P1.Y
                ToPhysical XC, YC
                DrawLine R2.P1.X, R2.P1.Y, R2.P1.X + Rad * Cos(Ang2), R2.P1.Y - Rad * Sin(Ang2), , vbInvert, False
                DrawCircle Paper.hDC, XC, YC, Radius, Ang1, Ang2, , vbInvert
                Ang2 = GetAngle(R2.P1.X, R2.P1.Y, X, Y)
                DrawLine R2.P1.X, R2.P1.Y, R2.P1.X + Rad * Cos(Ang2), R2.P1.Y - Rad * Sin(Ang2), , vbInvert, False
                ShowPoint Paper.hDC, DragS.Number(3)
                DrawCircle Paper.hDC, XC, YC, Radius, Ang1, Ang2, , vbInvert
                Paper.Refresh
        End Select
        DragS.OX = X
        DragS.OY = Y
    End If
        
'====================================================
'MIDDLE POINT
'====================================================

Case dsMiddlePoint
    If DragS.State = dscNormalState Then
        If IsPoint(AP) Then
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResFirst) & " " & GetString(ResPoint) & ". (" & BasePoint(AP).Name & ")"
        Else
            If IsFigure(AF) Then
                If Figures(AF).FigureType = dsSegment Then
                    i_ShowStatus GetString(ResCreateMiddlePoint) & " (" & BasePoint(Figures(AF).Points(0)).Name & BasePoint(Figures(AF).Points(1)).Name & ")."
                    DragS.OX = X
                    DragS.OY = Y
                    Exit Sub
                End If
            End If
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResFirst) & " " & GetString(ResPoint) & "."
        End If
    End If
    If DragS.State <> dscNormalState And DragS.NumOfMouseUps = 1 And DragS.Number(1) <> 0 And DragS.Number(2) <> -1 Then
        DrawPoint (BasePoint(DragS.Number(1)).X + DragS.OX) / 2, (BasePoint(DragS.Number(1)).Y + DragS.OY) / 2, , vbInvert
        DrawPoint (BasePoint(DragS.Number(1)).X + X) / 2, (BasePoint(DragS.Number(1)).Y + Y) / 2, , vbInvert
        Paper.Refresh
        If APIsAPoint And AP <> DragS.Number(1) Then
            i_ShowStatus GetString(ResCreateMiddlePoint) & " (" & BasePoint(DragS.Number(1)).Name & BasePoint(AP).Name & "). " & GetString(ResPressEsc)
        Else
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResSecond) & " " & GetString(ResPoint) & "." & " " & GetString(ResPressEsc)
        End If
    End If
    DragS.OX = X
    DragS.OY = Y
        
'====================================================
'SIMM POINT
'====================================================

Case dsSimmPoint
    If DragS.State = dscNormalState Then
        If IsPoint(AP) Then
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResFirst) & " " & GetString(ResPoint) & ". (" & BasePoint(AP).Name & ")"
        Else
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResFirst) & " " & GetString(ResPoint) & "."
        End If
    End If
    If DragS.State <> dscNormalState And DragS.NumOfMouseUps = 1 And DragS.Number(1) <> 0 Then
        If APIsAPoint And AP <> DragS.Number(1) Then
            i_ShowStatus GetString(ResCreateSymmetricPoint) & ". " & GetString(ResPressEsc)
        Else
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResSymmCenter) & "." & " " & GetString(ResPressEsc)
        End If
        DrawPoint 2 * DragS.OX - BasePoint(DragS.Number(1)).X, 2 * DragS.OY - BasePoint(DragS.Number(1)).Y, , vbInvert
        DrawPoint 2 * X - BasePoint(DragS.Number(1)).X, 2 * Y - BasePoint(DragS.Number(1)).Y, , vbInvert
        Paper.Refresh
    End If
    DragS.OX = X
    DragS.OY = Y
    
'====================================================
'SIMM POINT BY LINE
'====================================================

Case dsSimmPointByLine
    If DragS.State = dscNormalState Then
        If IsPoint(AP) Then
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResPoint) & ". (" & BasePoint(AP).Name & ")"
        ElseIf IsLine(AF) Then
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResLineRayOrSegment) & " (" & Figures(AF).Name & ")"
        Else
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResPoint) & ", " & GetString(ResLineRayOrSegment)
        End If
    End If
    If DragS.State = dscMovingState Then
        If DragS.Number(1) = 0 Then
            If IsPoint(AP) Then
                i_ShowStatus GetString(ResCreateSymmetricPoint) & ". " & GetString(ResPressEsc)
            Else
                i_ShowStatus GetString(ResLocate) & " " & GetString(ResPoint) & ". " & GetString(ResPressEsc)
            End If
        Else
            If IsLine(AF) Then
                i_ShowStatus GetString(ResCreateSymmetricPoint) & ". " & GetString(ResPressEsc)
            Else
                i_ShowStatus GetString(ResLocate) & " " & GetString(ResLineRayOrSegment) & " " & GetString(ResPressEsc)
            End If
        End If
    End If

'====================================================
'INVERTED POINT
'====================================================

Case dsInvert
    If DragS.State = dscNormalState Then
        If IsPoint(AP) Then
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResPoint) & ". (" & BasePoint(AP).Name & ")"
        ElseIf IsCircle(AF) Then
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResCircle) & ". (" & Figures(AF).Name & ")"
        Else
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResPoint) & ", " & GetString(ResCircle) & "."
        End If
    End If
    If DragS.State = dscMovingState Then
        If DragS.Number(1) = 0 Then
            If IsPoint(AP) Then
                i_ShowStatus GetString(ResCreateInvertedPoint) & ". " & GetString(ResPressEsc)
            Else
                i_ShowStatus GetString(ResLocate) & " " & GetString(ResPoint) & ". " & GetString(ResPressEsc)
            End If
        Else
            If IsCircle(AF) Then
                i_ShowStatus GetString(ResCreateInvertedPoint) & ". " & GetString(ResPressEsc)
            Else
                i_ShowStatus GetString(ResLocate) & " " & GetString(ResCircle) & ". " & GetString(ResPressEsc)
            End If
        End If
    End If

'====================================================
'INTERSECT
'====================================================

Case dsIntersect
    If DragS.State = dscNormalState Then
        AFS = GetFiguresByPoint(X, Y, , False)
        If LBound(AFS) + 1 = UBound(AFS) Then
            i_ShowStatus GetString(ResCreateIntersection) & ": " & Figures(AFS(LBound(AFS))).Name & ", " & Figures(AFS(UBound(AFS))).Name
        Else
            If IsFigure(AF) Then
                i_ShowStatus GetString(ResLocate) & " " & GetString(ResFirst) & " " & GetString(ResFigure) & " - " & Figures(AF).Name & "."
            Else
                i_ShowStatus GetString(ResLocate) & " " & GetString(ResFirst) & " " & GetString(ResFigure) & "."
            End If
        End If
    End If
    If DragS.State = dscMovingState Then
        If IsFigure(AF) And AF <> DragS.Number(1) Then
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResSecond) & " " & GetString(ResFigure) & " - " & Figures(AF).Name & ". " & GetString(ResPressEsc)
        Else
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResSecond) & " " & GetString(ResFigure) & ". " & GetString(ResPressEsc)
        End If
    End If

'====================================================
'POINT ON FIGURE
'====================================================

Case dsPointOnFigure
    If DragS.State = dscNormalState Then
        'ALoc = GetLocusFromPoint(X, Y)
        If ALoc > 0 Then If Not Locuses(ALoc).Dynamic Then ALoc = 0
        If AFIsAFigure And ALoc > 0 Then
            i_ShowStatus GetString(ResCreatePointOnFigure)
        Else
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResFigure) & "."
        End If
    End If

'====================================================
'MEASURE DISTANCE
'====================================================

Case dsMeasureDistance
    If DragS.State = dscNormalState Then
        If IsFigure(AF) Then
            If Figures(AF).FigureType = dsSegment Then
                i_ShowStatus GetString(ResMeasureDistance)
                Exit Sub
            End If
        End If
        If IsPoint(AP) Then
            i_ShowStatus Trim(GetString(ResLocate)) & " " & Trim(GetString(ResFirst)) & " " & GetString(ResPoint) & ". (" & BasePoint(AP).Name & ")"
        Else
            i_ShowStatus Trim(GetString(ResLocate)) & " " & Trim(GetString(ResFirst)) & " " & GetString(ResPoint) & "."
        End If
    Else
        If IsPoint(AP) Then
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResSecond) & " " & GetString(ResPoint) & ". (" & BasePoint(AP).Name & ") " & GetString(ResPressEsc)
        Else
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResSecond) & " " & GetString(ResPoint) & ". " & GetString(ResPressEsc)
        End If
    End If

'====================================================
'MEASURE ANGLE
'====================================================

Case dsMeasureAngle
    If DragS.State = dscNormalState Then
        If IsPoint(AP) Then
            i_ShowStatus Trim(GetString(ResLocate)) & " " & Trim(GetString(ResFirst)) & " " & GetString(ResPoint) & ". (" & BasePoint(AP).Name & ")"
        Else
            i_ShowStatus Trim(GetString(ResLocate)) & " " & Trim(GetString(ResFirst)) & " " & GetString(ResPoint) & "."
        End If
    Else
        If DragS.NumOfMouseUps < 2 Then
            If IsPoint(AP) Then
                i_ShowStatus GetString(ResLocateAngleVertex) & " (" & BasePoint(AP).Name & ") " & GetString(ResPressEsc)
            Else
                i_ShowStatus GetString(ResLocateAngleVertex) & " " & GetString(ResPressEsc)
            End If
        Else
            If IsPoint(AP) Then
                i_ShowStatus GetString(ResLocate) & " " & GetString(ResSecond) & " " & GetString(ResPoint) & ". (" & BasePoint(AP).Name & ") " & GetString(ResPressEsc)
            Else
                i_ShowStatus GetString(ResLocate) & " " & GetString(ResSecond) & " " & GetString(ResPoint) & ". " & GetString(ResPressEsc)
            End If
        End If
    End If

'====================================================
'DYNAMIC LOCUS
'====================================================

Case dsDynamicLocus
    If Not IsPoint(AP) Then
        i_ShowStatus GetString(ResLocate) & " " & GetString(ResPoint) & "."
    Else
        If BasePoint(AP).Type <> dsPoint Then
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResPoint) & " (" & BasePoint(AP).Name & ")"
        Else
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResPoint) & "."
        End If
    End If

'=========================================================
'                                           MEASURE AREA
'=========================================================

Case dsMeasureArea
    If APIsAPoint Then
        If DragS.NumberOfPoints > 1 Then
            If AP = DragS.Points(1) Then
                i_ShowStatus Replace(GetString(ResClickToClosePolygon), "%1", BasePoint(DragS.Points(1)).Name) & " " & GetString(ResPressEsc)
            Else
                i_ShowStatus GetString(ResLocate) & " " & GetString(ResPoint) & " (" & BasePoint(AP).Name & "). " & Replace(GetString(ResClickToClosePolygon), "%1", BasePoint(DragS.Points(1)).Name) & " " & GetString(ResPressEsc)
            End If
        Else
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResPoint) & " (" & BasePoint(AP).Name & "). " & GetString(ResPressEsc)
        End If
    Else
        If DragS.NumberOfPoints > 1 Then
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResPoint) & ". " & Replace(GetString(ResClickToClosePolygon), "%1", BasePoint(DragS.Points(1)).Name) & " " & GetString(ResPressEsc)
        Else
            If IsSG(ASG) And DragS.NumOfMouseUps = 0 Then
                If StaticGraphics(ASG).Type = sgPolygon Then
                    Dim T As String
                    
                    With StaticGraphics(ASG)
                        T = "("
                        For Z = LBound(.Points) To UBound(.Points) - 1
                            T = T & BasePoint(.Points(Z)).Name & IIf(Z = UBound(.Points) - 1, "", ",")
                        Next
                        T = T & ")."
                    End With
                    i_ShowStatus GetString(ResArea) & T
                    Exit Sub
                End If
            End If
            If IsCircle(AF) And DragS.NumOfMouseUps = 0 Then
                i_ShowStatus GetString(ResArea) & ": " & Figures(AF).Name & "."
                Exit Sub
            End If
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResPoint) & ". " & GetString(ResPressEsc)
        End If
    End If
    Exit Sub

'====================================================
'                               POLYGON
'====================================================

Case dsPolygon
    If APIsAPoint Then
        bPointAlreadySelected = False
        For Z = 2 To DragS.NumberOfPoints
            If DragS.Points(Z) = AP Then bPointAlreadySelected = True
        Next
        
        If bPointAlreadySelected Then
            TC = curStateCross
            If DragS.NumberOfPoints > 1 Then
                i_ShowStatus GetString(ResLocate) & " " & GetString(ResPoint) & ". " & Replace(GetString(ResClickToClosePolygon), "%1", BasePoint(DragS.Points(1)).Name) & " " & GetString(ResPressEsc)
            Else
                i_ShowStatus GetString(ResLocate) & " " & GetString(ResPoint) & ". " & GetString(ResPressEsc)
            End If
        
        Else
        
            TC = curStateAdd
            If DragS.NumberOfPoints > 1 Then
                If AP = DragS.Points(1) Then
                    i_ShowStatus Replace(GetString(ResClickToClosePolygon), "%1", BasePoint(DragS.Points(1)).Name) & " " & GetString(ResPressEsc)
                Else
                    i_ShowStatus GetString(ResLocate) & " " & GetString(ResPoint) & " (" & BasePoint(AP).Name & "). " & Replace(GetString(ResClickToClosePolygon), "%1", BasePoint(DragS.Points(1)).Name) & " " & GetString(ResPressEsc)
                End If
            Else
                i_ShowStatus GetString(ResLocate) & " " & GetString(ResPoint) & " (" & BasePoint(AP).Name & "). " & GetString(ResPressEsc)
            End If
            
        End If
    
    Else 'Ap is not a point
        TC = curStateCross
        If DragS.NumberOfPoints > 1 Then
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResPoint) & ". " & Replace(GetString(ResClickToClosePolygon), "%1", BasePoint(DragS.Points(1)).Name) & " " & GetString(ResPressEsc)
        Else
            i_ShowStatus GetString(ResLocate) & " " & GetString(ResPoint) & ". " & GetString(ResPressEsc)
        End If
    End If
    
    i_SetMousePointer TC
    
    If DragS.NumberOfPoints = 1 Then
        DrawLine BasePoint(DragS.Points(1)).X, BasePoint(DragS.Points(1)).Y, DragS.OX, DragS.OY, 0, vbInvert, False
        DrawLine BasePoint(DragS.Points(1)).X, BasePoint(DragS.Points(1)).Y, X, Y, 0, 6, False
        ShowSelectedPoint Paper.hDC, DragS.Points(1)
        Paper.Refresh
    End If
    If DragS.NumberOfPoints > 1 Then
        PaperCls
        ShowAll , , False
        For Z = 1 To DragS.NumberOfPoints
            ShowSelectedPoint Paper.hDC, DragS.Points(Z)
        Next
        DragS.Pixels(DragS.NumberOfPoints + 1).X = OX
        DragS.Pixels(DragS.NumberOfPoints + 1).Y = OY
        DrawPolygon DragS.Pixels(), setdefcolFigure, setdefcolFigureFill, 9, True
    End If
    DragS.OX = X
    DragS.OY = Y
End Select

End Sub

'################################################################
'Process virtual paper mouseup event
'################################################################

Public Sub PaperMouseUp(Button As Integer, Shift As Integer, OX As Single, OY As Single)
'======================================================

On Local Error Resume Next

'======================================================

Dim EndOperation As Boolean
Dim pAction As Action

Dim AP As Long, AF As Long, Z As Long, AB As Long
Dim AFS() As Long
Dim bPointAlreadySelected As Boolean

Dim P As OnePoint
Dim R As TwoPoints
Dim R2 As TwoPoints
Dim XC As Double, YC As Double, Radius As Double
Dim Ang1 As Double, Ang2 As Double, Rad As Double
Dim X As Double, Y As Double

'======================================================

X = OX
Y = OY
ToLogical X, Y

'======================================================
'                                                   DEMO
'======================================================

If DragS.State = dscDemo Then
    Exit Sub
End If

'======================================================
'                                   POINT LABEL DRAG
'======================================================

If DragS.State = dscPointLabelDrag Then
    DragS.State = dscNormalState
    DragS.Number(1) = 0
    DragS.Number(2) = 0
    DragS.Number(3) = 0
    Exit Sub
End If

'======================================================
'                                   MEASURE DRAG
'======================================================

If DragS.State = dscMeasureDrag Then
    DragS.State = dscNormalState
    DragS.Number(1) = 0
    DragS.Number(2) = 0
    DragS.Number(3) = 0
    Exit Sub
End If

'======================================================
'                               SELECT OBJECTS
'======================================================

If DragS.State = dscSelectObjects Then Exit Sub

'======================================================
'                                   SCROLL
'======================================================

If DragS.State = dscScroll Then
    i_Scrolled True, True
    DragS.State = dscNormalState
    DragS.OX = X
    DragS.OY = Y
    Exit Sub
End If

'======================================================
'                       MACRO RUN
'======================================================

If DragS.State = dscMacroStateRun Then
    If DragS.MacroCurrentObject > DragS.MacroObjectCount Then
        RunMacro DragS.Number(1), DragS.MacroObjects
        i_ExitMacroRunMode
    End If
    Exit Sub
End If

'======================================================
'               MACRO GIVENS AND RESULTS NOT PARTICIPATING
'======================================================

If DragS.State = dscMacroStateGivens Or DragS.State = dscMacroStateResults Then Exit Sub

'======================================================
'                                   STATIC GRAPHIC
'======================================================
If DragS.State = dscCreateStaticGraphic Then
    AP = GetPointFromCursor(X, Y)
    
    DragS.NumberOfPoints = DragS.NumberOfPoints + 1
    ReDim Preserve DragS.Points(1 To DragS.NumberOfPoints)
    
    If IsPoint(AP) Then
        bPointAlreadySelected = False
        For Z = 2 To DragS.NumberOfPoints - 1
            If DragS.Points(Z) = AP Then bPointAlreadySelected = True
        Next
        
        If Not bPointAlreadySelected Then
            DragS.Points(DragS.NumberOfPoints) = AP
        Else
            AddBasePoint X, Y
            DragS.Points(DragS.NumberOfPoints) = PointCount
        End If
    Else
        AddBasePoint X, Y
        DragS.Points(DragS.NumberOfPoints) = PointCount
        If DragS.NumberOfPoints = 2 Then If DragS.TypeOfStaticGraphic = sgVector Then BasePoint(PointCount).Hide = True
    End If
    
    ShowSelectedPoint Paper.hDC, DragS.Points(DragS.NumberOfPoints), True
    
    Select Case DragS.TypeOfStaticGraphic
        Case sgBezier
            If DragS.NumberOfPoints = 4 Then
                AddStaticGraphic sgBezier, DragS.Points
                EndOperation = True
                GoTo EndOperation
            End If
        
        Case sgPolygon
            ReDim Preserve DragS.Pixels(1 To DragS.NumberOfPoints + 1)
            DragS.Pixels(DragS.NumberOfPoints).X = BasePoint(DragS.Points(DragS.NumberOfPoints)).PhysicalX
            DragS.Pixels(DragS.NumberOfPoints).Y = BasePoint(DragS.Points(DragS.NumberOfPoints)).PhysicalY
            If DragS.NumberOfPoints = 1 Then
                DragS.ShouldHideFirstPointName = Not BasePoint(DragS.Points(1)).ShowName
                BasePoint(DragS.Points(1)).ShowName = True
                DrawLine BasePoint(DragS.Points(1)).X, BasePoint(DragS.Points(1)).Y, X, Y, 0, 6, False
                ShowSelectedPoint Paper.hDC, DragS.Points(DragS.NumberOfPoints), True
            End If
            If DragS.NumberOfPoints = 2 Then
                DrawLine BasePoint(DragS.Points(1)).X, BasePoint(DragS.Points(1)).Y, DragS.OX, DragS.OY, 0, 6, False
                ShowSelectedPoint Paper.hDC, DragS.Points(DragS.NumberOfPoints)
                'DragS.Pixels(DragS.NumberOfPoints + 1).X = OX
                'DragS.Pixels(DragS.NumberOfPoints + 1).Y = OY
                'DrawPolygon DragS.Pixels(), setdefcolFigure, setdefcolFigureFill, 7, True
            End If
            
            If DragS.NumberOfPoints > 3 And ((DragS.Points(1) = DragS.Points(DragS.NumberOfPoints))) Then
                'ToPhysical DragS.OX, DragS.OY
                'DragS.Pixels(DragS.NumberOfPoints + 1).X = DragS.OX
                'DragS.Pixels(DragS.NumberOfPoints + 1).Y = DragS.OY
                'DrawPolygon DragS.Pixels(), setdefcolFigure, setdefcolFigureFill, 7, False
                
                BasePoint(DragS.Points(1)).ShowName = Not DragS.ShouldHideFirstPointName
                DragS.ShouldHideFirstPointName = False
                AddStaticGraphic sgPolygon, DragS.Points
                'i_CheckMainbarButton GetString(ResStaticObjectBase + 2 * sgPolygon), False
                
                EndOperation = True
                GoTo EndOperation
            End If
            
            If DragS.NumberOfPoints > 2 And Button = 2 Then
                BasePoint(DragS.Points(1)).ShowName = Not DragS.ShouldHideFirstPointName
                DragS.ShouldHideFirstPointName = False
                AddStaticGraphic sgPolygon, DragS.Points
                'i_CheckMainbarButton GetString(ResStaticObjectBase + 2 * sgPolygon), False
                
                EndOperation = True
                GoTo EndOperation
            End If
            
        Case sgVector
            If DragS.NumberOfPoints = 2 Then
                If DragS.Points(1) <> DragS.Points(2) Then AddStaticGraphic sgVector, DragS.Points
                EndOperation = True
                GoTo EndOperation
            End If
    End Select
    DragS.OX = X
    DragS.OY = Y
    Exit Sub
End If

'====================================================
'                                           BUTTON
'====================================================
If DrawingState = dsSelect And DragS.Number(1) <> 0 And DragS.WhatDoIDrag = gotButton And DragS.State = dscDraggingState And Button = 2 Then
    If DragS.OX <> DragS.X Or DragS.OY <> DragS.Y Then
        pAction.Type = actMoveButton
        pAction.pButton = DragS.Number(1)
        ReDim pAction.AuxPoints(1 To 1)
        pAction.AuxPoints(1).X = DragS.X - DragS.OX + Buttons(DragS.Number(1)).LogicalPosition.P1.X
        pAction.AuxPoints(1).Y = DragS.Y - DragS.OY + Buttons(DragS.Number(1)).LogicalPosition.P1.Y
        RecordAction pAction
    End If

    If DragS.OX = DragS.X And DragS.OY = DragS.Y Then
        AB = GetButtonFromPoint(X, Y)
        If AB = DragS.Number(1) Then
            i_ButtonRightClicked AB
        End If
    End If
    
    EndOperation = True
End If

'====================================================

If DragS.NumOfMouseUps >= DragS.NumOfMouseDowns Or DragS.State = dscErrorState Then EndOperation = True: GoTo EndOperation
If Button <> 1 And DragS.State <> dscNormalState And DrawingState <> dsPolygon Then EndOperation = True: GoTo EndOperation
If Button <> 1 And DrawingState <> dsPolygon Then Exit Sub '?????

'====================================================

If Shift = 1 Then
    X = Round(X)
    Y = Round(Y)
End If

AF = GetFigureByPoint(X, Y)
AP = GetPointFromCursor(X, Y)

'====================================================
'                           DRAWING STATE
'====================================================

Select Case DrawingState
'###############################
'SELECT
Case dsSelect
    If DragS.Number(1) <> 0 Then
        ManualDragFlag = 0
        If DragS.WhatDoIDrag = gotPoint Then
            'ValueTable1.UpdateExpressions?????WEWEWE
            RenderHighQualityDynamicLoci
        ElseIf DragS.WhatDoIDrag = gotLabel Then
            If DragS.X <> DragS.OX Or DragS.Y <> DragS.OY Then
                pAction.Type = actMoveLabel
                pAction.pLabel = DragS.Number(1)
                ReDim pAction.AuxPoints(1 To 1)
                pAction.AuxPoints(1).X = DragS.X - DragS.OX + TextLabels(DragS.Number(1)).LogicalPosition.P1.X
                pAction.AuxPoints(1).Y = DragS.Y - DragS.OY + TextLabels(DragS.Number(1)).LogicalPosition.P1.Y
                RecordAction pAction
            End If
        ElseIf DragS.WhatDoIDrag = gotButton Then
            'ShowButton Paper.hDC, DragS.Number(1), True, True
            Buttons(DragS.Number(1)).Pushed = False
            PaperCls
            ShowAll
            If GetButtonFromPoint(X, Y) = DragS.Number(1) Then ButtonPushed DragS.Number(1)
        End If
    End If
    EndOperation = True

'###############################
'POINT
Case dsPoint
    If DragS.Number(1) = 0 And Button = 1 Then 'And Not IsPoint(AP)
        AFS = GetFiguresByPoint(X, Y, , False)
        If LBound(AFS) + 1 = UBound(AFS) Then
            AddIntersectionPoints AFS(LBound(AFS)), AFS(UBound(AFS))
            ShowAll
        Else
            If IsFigure(AF) Then
                AddPointOnFigure AF, X, Y
            Else
                If Shift = 1 Then X = Round(X): Y = Round(Y)
                AddBasePoint X, Y
            End If
        End If
    End If
    EndOperation = True

'###############################
'SEGMENT
Case dsSegment
        If DragS.NumOfMouseUps > 0 And DragS.Number(1) <> 0 Then
            If AP = 0 Then
                AddBasePoint X, Y
                DragS.Number(2) = PointCount
            Else
                DragS.Number(2) = AP
            End If
            If DragS.Number(1) <> DragS.Number(2) Then AddSegment DragS.Number(1), DragS.Number(2)
            DrawLine BasePoint(DragS.Number(1)).X, BasePoint(DragS.Number(1)).Y, DragS.OX, DragS.OY, , vbInvert, False
            'If Paper.AutoRedraw <> DragS.OldAutoRedraw Then Paper.AutoRedraw = DragS.OldAutoRedraw
            ShowAll
            EndOperation = True
        Else
            DrawLine BasePoint(DragS.Number(1)).X, BasePoint(DragS.Number(1)).Y, DragS.OX, DragS.OY, , vbInvert
        End If
        
'###############################
'RAY

Case dsRay
        If DragS.NumOfMouseUps > 0 And DragS.Number(1) <> 0 Then
            If AP = 0 Then
                AddBasePoint X, Y
                DragS.Number(2) = PointCount
            Else
                DragS.Number(2) = AP
            End If
            R.P1.X = BasePoint(DragS.Number(1)).X
            R.P1.Y = BasePoint(DragS.Number(1)).Y
            R.P2.X = DragS.OX
            R.P2.Y = DragS.OY
            R2 = GetLineFromSegment(R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
            R2.P2 = R.P1
            DrawLine R2.P1.X, R2.P1.Y, R2.P2.X, R2.P2.Y, , vbInvert, False
            If DragS.Number(1) <> DragS.Number(2) Then AddRay DragS.Number(1), DragS.Number(2)
            'If Paper.AutoRedraw <> DragS.OldAutoRedraw Then Paper.AutoRedraw = DragS.OldAutoRedraw
            ShowAll
            EndOperation = True
        Else
            R.P1.X = BasePoint(DragS.Number(1)).X
            R.P1.Y = BasePoint(DragS.Number(1)).Y
            R.P2.X = DragS.OX
            R.P2.Y = DragS.OY
            R2 = GetLineFromSegment(R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
            R2.P2 = R.P1
            DrawLine R2.P1.X, R2.P1.Y, R2.P2.X, R2.P2.Y, , vbInvert
        End If
        
'###############################
'LINE_2POINTS

Case dsLine_2Points
    If DragS.NumOfMouseUps > 0 And DragS.Number(1) <> 0 Then
        If AP = 0 Then
            AddBasePoint X, Y
            DragS.Number(2) = PointCount
        Else
            DragS.Number(2) = AP
        End If
        R.P1.X = BasePoint(DragS.Number(1)).X
        R.P1.Y = BasePoint(DragS.Number(1)).Y
        R.P2.X = DragS.OX
        R.P2.Y = DragS.OY
        R = GetLineFromSegment(R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
        DrawLine R.P1.X, R.P1.Y, R.P2.X, R.P2.Y, , vbInvert, False
        If DragS.Number(1) <> DragS.Number(2) Then AddLine DragS.Number(1), DragS.Number(2)
        'If Paper.AutoRedraw <> DragS.OldAutoRedraw Then Paper.AutoRedraw = DragS.OldAutoRedraw
        ShowAll
        EndOperation = True
    Else
        R.P1.X = BasePoint(DragS.Number(1)).X
        R.P1.Y = BasePoint(DragS.Number(1)).Y
        R.P2.X = DragS.OX
        R.P2.Y = DragS.OY
        R = GetLineFromSegment(R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
        DrawLine R.P1.X, R.P1.Y, R.P2.X, R.P2.Y, , vbInvert
    End If
        
'###############################
'BISECTOR
Case dsBisector
    Select Case DragS.NumOfMouseUps
        Case 1
            If Not IsPoint(AP) Then
                AddBasePoint X, Y
                AP = PointCount
            End If
            
            DragS.Number(2) = AP
            ShowSelectedPoint Paper.hDC, AP
            
            R.P1.X = BasePoint(DragS.Number(2)).X
            R.P1.Y = BasePoint(DragS.Number(2)).Y
            R.P2 = GetBisector(BasePoint(DragS.Number(1)).X, BasePoint(DragS.Number(1)).Y, BasePoint(DragS.Number(2)).X, BasePoint(DragS.Number(2)).Y, X, Y)
            R2 = GetLineFromSegment(R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
            DrawLine R2.P1.X, R2.P1.Y, R2.P2.X, R2.P2.Y, , vbInvert
            
            Paper.Refresh
        Case 2
            If Not IsPoint(AP) Then
                AddBasePoint X, Y
                AP = PointCount
            End If
            
            DragS.Number(3) = AP
            ShowSelectedPoint Paper.hDC, AP
            
            R.P1.X = BasePoint(DragS.Number(2)).X
            R.P1.Y = BasePoint(DragS.Number(2)).Y
            R.P2 = GetBisector(BasePoint(DragS.Number(1)).X, BasePoint(DragS.Number(1)).Y, BasePoint(DragS.Number(2)).X, BasePoint(DragS.Number(2)).Y, DragS.OX, DragS.OY)
            R2 = GetLineFromSegment(R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
            DrawLine R2.P1.X, R2.P1.Y, R2.P2.X, R2.P2.Y, , vbInvert, False
            
            If DragS.Number(1) <> DragS.Number(2) And DragS.Number(2) <> DragS.Number(3) And DragS.Number(3) <> DragS.Number(1) Then AddBisector DragS.Number(1), DragS.Number(2), DragS.Number(3)
            ShowAll
            EndOperation = True
    End Select
    
        
'###############################
'LINE_PARALLEL

Case dsLine_PointAndParallelLine
    If DragS.NumOfMouseUps > 0 And IsLine(DragS.Number(1)) Then
        If AP = 0 Then
            AddBasePoint X, Y
            DragS.Number(2) = PointCount
        Else
            DragS.Number(2) = AP
        End If
        
        R = GetLineCoordinates(DragS.Number(1))
        R = GetLineFromSegment(DragS.OX, DragS.OY, DragS.OX + (R.P1.X - R.P2.X), DragS.OY + (R.P1.Y - R.P2.Y))
        DrawLine R.P1.X, R.P1.Y, R.P2.X, R.P2.Y, , vbInvert, False
        
        AddParallelLine DragS.Number(2), DragS.Number(1)
        'If Paper.AutoRedraw <> DragS.OldAutoRedraw Then Paper.AutoRedraw = DragS.OldAutoRedraw
        ShowAll
        EndOperation = True
    End If
    If DragS.NumOfMouseUps > 0 And DragS.Number(1) = -1 And IsPoint(DragS.Number(2)) Then
        If IsLine(AF) Then
            DragS.Number(1) = AF
            AddParallelLine DragS.Number(2), DragS.Number(1)
        End If
        'If Paper.AutoRedraw <> DragS.OldAutoRedraw Then Paper.AutoRedraw = DragS.OldAutoRedraw
        PaperCls
        ShowAll
        EndOperation = True
    End If
        
'###############################
'LINE_PERPENDICULAR

Case dsLine_PointAndPerpendicularLine
    If DragS.NumOfMouseUps > 0 And IsLine(DragS.Number(1)) Then
        If Not IsPoint(AP) Then
            AddBasePoint X, Y
            DragS.Number(2) = PointCount
        Else
            DragS.Number(2) = AP
        End If
        
        R = GetLineCoordinates(DragS.Number(1))
        R = GetPerpendicularLine(DragS.OX, DragS.OY, R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
        DrawLine R.P1.X, R.P1.Y, R.P2.X, R.P2.Y, , vbInvert, False
        
        AddPerpendicularLine DragS.Number(2), DragS.Number(1)
        'If Paper.AutoRedraw <> DragS.OldAutoRedraw Then Paper.AutoRedraw = DragS.OldAutoRedraw
        ShowAll
        EndOperation = True
    End If
    If DragS.NumOfMouseUps > 0 And DragS.Number(1) = -1 And IsPoint(DragS.Number(2)) Then
        If IsLine(AF) Then
            DragS.Number(1) = AF
            AddPerpendicularLine DragS.Number(2), DragS.Number(1)
        End If
        'If Paper.AutoRedraw <> DragS.OldAutoRedraw Then Paper.AutoRedraw = DragS.OldAutoRedraw
        PaperCls
        ShowAll
        EndOperation = True
    End If
        
'###############################
'CIRCLE

Case dsCircle_CenterAndCircumPoint
    If DragS.NumOfMouseUps > 0 And DragS.Number(1) <> 0 Then
        If Not IsPoint(AP) Then
            AddBasePoint X, Y
            DragS.Number(2) = PointCount
        Else
            DragS.Number(2) = AP
        End If
        If DragS.Number(1) <> DragS.Number(2) Then AddCircle DragS.Number(1), DragS.Number(2)
        R.P1.X = BasePoint(DragS.Number(1)).X
        R.P1.Y = BasePoint(DragS.Number(1)).Y
        R.P2.X = DragS.OX
        R.P2.Y = DragS.OY
        Rad = Distance(R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
        Radius = Rad
        ToPhysicalLength Radius
        XC = R.P1.X
        YC = R.P1.Y
        ToPhysical XC, YC
        DrawCircle Paper.hDC, XC, YC, Radius, , , , vbInvert
        
        'If Paper.AutoRedraw <> DragS.OldAutoRedraw Then Paper.AutoRedraw = DragS.OldAutoRedraw
        ShowAll
        EndOperation = True
    End If
        
'###############################
'CIRCLE BY RADIUS

Case dsCircle_CenterAndTwoPoints
    If DragS.NumOfMouseUps = 1 And DragS.Number(1) <> 0 Then
        DrawLine BasePoint(DragS.Number(1)).X, BasePoint(DragS.Number(1)).Y, DragS.OX, DragS.OY, , vbInvert, False
        ShowPoint Paper.hDC, DragS.Number(1)
        If Not IsPoint(AP) Then
            'Paper.AutoRedraw = True
            AddBasePoint X, Y
            'Paper.AutoRedraw = False
            DragS.Number(2) = PointCount
        Else
            DragS.Number(2) = AP
        End If
        
        If DragS.Number(1) = DragS.Number(2) Then
            EndOperation = True
            GoTo EndOperation
        End If
        
        R.P1.X = BasePoint(DragS.Number(1)).X
        R.P1.Y = BasePoint(DragS.Number(1)).Y
        R.P2.X = BasePoint(DragS.Number(2)).X
        R.P2.Y = BasePoint(DragS.Number(2)).Y
        Rad = Distance(R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
        Radius = Rad
        ToPhysicalLength Radius
        XC = X
        YC = Y
        ToPhysical XC, YC
        DrawCircle Paper.hDC, XC, YC, Radius, , , , vbInvert
        Paper.Refresh
    End If
    If DragS.NumOfMouseUps = 2 And DragS.Number(1) <> 0 And DragS.Number(2) <> 0 Then
        R.P1.X = BasePoint(DragS.Number(1)).X
        R.P1.Y = BasePoint(DragS.Number(1)).Y
        R.P2.X = BasePoint(DragS.Number(2)).X
        R.P2.Y = BasePoint(DragS.Number(2)).Y
        Rad = Distance(R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
        Radius = Rad
        ToPhysicalLength Radius
        XC = DragS.OX
        YC = DragS.OY
        ToPhysical XC, YC
        DrawCircle Paper.hDC, XC, YC, Radius, , , , vbInvert
        If Not IsPoint(AP) Then
            'Paper.AutoRedraw = True
            AddBasePoint X, Y, , , , , , , False
            'Paper.AutoRedraw = False
            DragS.Number(3) = PointCount
        Else
            DragS.Number(3) = AP
        End If
        AddCircleByRadius DragS.Number(1), DragS.Number(2), DragS.Number(3)
        
        'If Paper.AutoRedraw <> DragS.OldAutoRedraw Then Paper.AutoRedraw = DragS.OldAutoRedraw
        ShowAll
        EndOperation = True
    End If
        
'###############################
'ARC

Case dsCircle_ArcCenterAndRadiusAndTwoPoints
    If DragS.NumOfMouseUps = 1 And DragS.Number(1) <> 0 Then
        DrawLine BasePoint(DragS.Number(1)).X, BasePoint(DragS.Number(1)).Y, DragS.OX, DragS.OY, , vbInvert, False
        ShowPoint Paper.hDC, DragS.Number(1)
        If Not IsPoint(AP) Then
            'Paper.AutoRedraw = True
            AddBasePoint X, Y
            'Paper.AutoRedraw = False
            DragS.Number(2) = PointCount
        Else
            DragS.Number(2) = AP
        End If
        
        If DragS.Number(1) = DragS.Number(2) Then
            EndOperation = True
            GoTo EndOperation
        End If
        
        R.P1.X = BasePoint(DragS.Number(1)).X
        R.P1.Y = BasePoint(DragS.Number(1)).Y
        R.P2.X = BasePoint(DragS.Number(2)).X
        R.P2.Y = BasePoint(DragS.Number(2)).Y
        Rad = Distance(R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
        Radius = Rad
        ToPhysicalLength Radius
        XC = X
        YC = Y
        ToPhysical XC, YC
        DrawCircle Paper.hDC, XC, YC, Radius, , , , vbInvert
        Paper.Refresh
    End If
    If DragS.NumOfMouseUps = 2 And DragS.Number(1) <> 0 And DragS.Number(2) <> 0 Then
        R.P1.X = BasePoint(DragS.Number(1)).X
        R.P1.Y = BasePoint(DragS.Number(1)).Y
        R.P2.X = BasePoint(DragS.Number(2)).X
        R.P2.Y = BasePoint(DragS.Number(2)).Y
        Rad = Distance(R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
        Radius = Rad
        ToPhysicalLength Radius
        XC = DragS.OX
        YC = DragS.OY
        ToPhysical XC, YC
        DrawCircle Paper.hDC, XC, YC, Radius, , , , vbInvert
        If Not IsPoint(AP) Then
            'Paper.AutoRedraw = True
            AddBasePoint X, Y
            'Paper.AutoRedraw = False
            DragS.Number(3) = PointCount
        Else
            DragS.Number(3) = AP
        End If
        
        R2.P1.X = BasePoint(DragS.Number(3)).X
        R2.P1.Y = BasePoint(DragS.Number(3)).Y
        Ang1 = GetAngle(R2.P1.X, R2.P1.Y, X, Y)
        DrawLine R2.P1.X, R2.P1.Y, R2.P1.X + Rad * Cos(Ang1), R2.P1.Y - Rad * Sin(Ang1), , vbInvert, False
        ShowPoint Paper.hDC, DragS.Number(3), , True
    End If
    If DragS.NumOfMouseUps = 3 And DragS.Number(1) <> 0 And DragS.Number(2) <> 0 And DragS.Number(3) <> 0 Then
        R.P1.X = BasePoint(DragS.Number(1)).X
        R.P1.Y = BasePoint(DragS.Number(1)).Y
        R.P2.X = BasePoint(DragS.Number(2)).X
        R.P2.Y = BasePoint(DragS.Number(2)).Y
        Rad = Distance(R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
        R2.P1.X = BasePoint(DragS.Number(3)).X
        R2.P1.Y = BasePoint(DragS.Number(3)).Y
        Ang1 = GetAngle(R2.P1.X, R2.P1.Y, DragS.OX, DragS.OY)
        DrawLine R2.P1.X, R2.P1.Y, R2.P1.X + Rad * Cos(Ang1), R2.P1.Y - Rad * Sin(Ang1), , vbInvert, False
        ShowPoint Paper.hDC, DragS.Number(3), , True
        If Not IsPoint(AP) Then
            'Paper.AutoRedraw = True
            AddBasePoint X, Y
            'Paper.AutoRedraw = False
            DragS.Number(4) = PointCount
        Else
            DragS.Number(4) = AP
        End If
        
        If DragS.Number(4) = DragS.Number(3) Then
            EndOperation = True
            GoTo EndOperation
        End If
        
'        R.P1.X = BasePoint(DragS.Number(1)).X
'        R.P1.Y = BasePoint(DragS.Number(1)).Y
'        R.P2.X = BasePoint(DragS.Number(2)).X
'        R.P2.Y = BasePoint(DragS.Number(2)).Y
'        Rad = Distance(R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
'        R2.P1.X = BasePoint(DragS.Number(3)).X
'        R2.P1.Y = BasePoint(DragS.Number(3)).Y
'        R2.P2.X = BasePoint(DragS.Number(4)).X
'        R2.P2.Y = BasePoint(DragS.Number(4)).Y
'        Ang1 = GetAngle(R2.P1.X, R2.P1.Y, R2.P2.X, R2.P2.Y)
'        Ang2 = GetAngle(R2.P1.X, R2.P1.Y, X, Y)
'        DrawLine R2.P1.X, R2.P1.Y, R2.P1.X + Rad * Cos(Ang1), R2.P1.Y - Rad * Sin(Ang1), , vbInvert
                R.P1.X = BasePoint(DragS.Number(1)).X
                R.P1.Y = BasePoint(DragS.Number(1)).Y
                R.P2.X = BasePoint(DragS.Number(2)).X
                R.P2.Y = BasePoint(DragS.Number(2)).Y
                Rad = Distance(R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
                R2.P1.X = BasePoint(DragS.Number(3)).X
                R2.P1.Y = BasePoint(DragS.Number(3)).Y
                R2.P2.X = BasePoint(DragS.Number(4)).X
                R2.P2.Y = BasePoint(DragS.Number(4)).Y
                Ang1 = GetAngle(R2.P1.X, R2.P1.Y, R2.P2.X, R2.P2.Y)
                Ang2 = GetAngle(R2.P1.X, R2.P1.Y, X, Y)
                Radius = Rad
                ToPhysicalLength Radius
                XC = R2.P1.X
                YC = R2.P1.Y
                ToPhysical XC, YC
                DrawCircle Paper.hDC, XC, YC, Radius, Ang1, Ang2, , vbInvert
            Paper.Refresh

        'R.P1.X = BasePoint(DragS.Number(1)).X
        'R.P1.Y = BasePoint(DragS.Number(1)).Y
        'R.P2.X = BasePoint(DragS.Number(2)).X
        'R.P2.Y = BasePoint(DragS.Number(2)).Y
        'R2.P1.X = BasePoint(DragS.Number(3)).X
        'R2.P1.Y = BasePoint(DragS.Number(3)).Y
        'R2.P2.X = BasePoint(DragS.Number(4)).X
        'R2.P2.Y = BasePoint(DragS.Number(4)).Y
        'Rad = Distance(R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
        'Ang1 = GetAngle(R2.P1.X, R2.P1.Y, R2.P2.X, R2.P2.Y)
        'Ang2 = GetAngle(R2.P1.X, R2.P1.Y, X, Y) + 10 * ToRad
        'Radius = Rad
        'ToPhysicalLength Radius
        'XC = R2.P1.X
        'YC = R2.P1.Y
        'ToPhysical XC, YC
        'DrawCircle  Paper.hDC,XC, YC, Radius, Ang1, Ang2, , vbInvert
    End If
    If DragS.NumOfMouseUps = 4 And DragS.Number(1) <> 0 And DragS.Number(2) <> 0 And DragS.Number(3) <> 0 And DragS.Number(4) <> 0 Then
'        R.P1.X = BasePoint(DragS.Number(1)).X
'        R.P1.Y = BasePoint(DragS.Number(1)).Y
'        R.P2.X = BasePoint(DragS.Number(2)).X
'        R.P2.Y = BasePoint(DragS.Number(2)).Y
'        Rad = Distance(R.P1.X, R.P1.Y, R.P2.X, R.P2.Y)
'        R2.P1.X = BasePoint(DragS.Number(3)).X
'        R2.P1.Y = BasePoint(DragS.Number(3)).Y
'        R2.P2.X = BasePoint(DragS.Number(4)).X
'        R2.P2.Y = BasePoint(DragS.Number(4)).Y
'        Ang1 = GetAngle(R2.P1.X, R2.P1.Y, R2.P2.X, R2.P2.Y)
'        Ang2 = GetAngle(R2.P1.X, R2.P1.Y, DragS.OX, DragS.OY)
'        DrawLine R2.P1.X, R2.P1.Y, R2.P1.X + Rad * Cos(Ang1), R2.P1.Y - Rad * Sin(Ang1), , vbInvert
'        DrawLine R2.P1.X, R2.P1.Y, R2.P1.X + Rad * Cos(Ang2), R2.P1.Y - Rad * Sin(Ang2), , vbInvert
'        Radius = Rad
'        ToPhysicalLength Radius
'        XC = DragS.OX
'        YC = DragS.OY
'        ToPhysical XC, YC
'        DrawCircle Paper.hDC, XC, YC, Radius, Ang1, Ang2, , vbInvert
        If Not IsPoint(AP) Then
            'Paper.AutoRedraw = True
            AddBasePoint X, Y
            'Paper.AutoRedraw = False
            DragS.Number(5) = PointCount
        Else
            DragS.Number(5) = AP
        End If
        
        If (DragS.Number(3) <> DragS.Number(5)) Then AddArc DragS.Number(1), DragS.Number(2), DragS.Number(3), DragS.Number(4), DragS.Number(5)
        'If Paper.AutoRedraw <> DragS.OldAutoRedraw Then Paper.AutoRedraw = DragS.OldAutoRedraw
        ShowAll
        EndOperation = True
    End If
  
'###############################
'MIDDLE POINT
Case dsMiddlePoint
    If DragS.Number(2) = -1 Then
        AddMiddlePoint Figures(DragS.Number(1)).Points(0), Figures(DragS.Number(1)).Points(1)
        'If Paper.AutoRedraw <> DragS.OldAutoRedraw Then Paper.AutoRedraw = DragS.OldAutoRedraw
        ShowAll
        EndOperation = True
        GoTo EndOperation
    End If
    If DragS.NumOfMouseUps > 0 And DragS.Number(1) <> 0 Then
        If AP = 0 Then
            AddBasePoint X, Y
            DragS.Number(2) = PointCount
        Else
            DragS.Number(2) = AP
        End If
        DrawPoint (BasePoint(DragS.Number(1)).X + DragS.OX) / 2, (BasePoint(DragS.Number(1)).Y + DragS.OY) / 2, , vbInvert
        If DragS.Number(1) <> DragS.Number(2) Then AddMiddlePoint DragS.Number(1), DragS.Number(2)
        'If Paper.AutoRedraw <> DragS.OldAutoRedraw Then Paper.AutoRedraw = DragS.OldAutoRedraw
        ShowAll
        EndOperation = True
    Else
        ShowSelectedPoint Paper.hDC, DragS.Number(1)
        DrawPoint (BasePoint(DragS.Number(1)).X + DragS.OX) / 2, (BasePoint(DragS.Number(1)).Y + DragS.OY) / 2, , vbInvert
        Paper.Refresh
End If
        
'###############################
'SIMM POINT
Case dsSimmPoint
    If DragS.NumOfMouseUps > 0 And DragS.Number(1) <> 0 Then
        If AP = 0 Then
            AddBasePoint X, Y
            DragS.Number(2) = PointCount
        Else
            DragS.Number(2) = AP
        End If
        DrawPoint 2 * DragS.OX - BasePoint(DragS.Number(1)).X, 2 * DragS.OY - BasePoint(DragS.Number(1)).Y, , vbInvert
        If DragS.Number(1) <> DragS.Number(2) Then AddSimmPoint DragS.Number(1), DragS.Number(2)
        'If Paper.AutoRedraw <> DragS.OldAutoRedraw Then Paper.AutoRedraw = DragS.OldAutoRedraw
        ShowAll
        EndOperation = True
    Else
        ShowSelectedPoint Paper.hDC, DragS.Number(1)
        DrawPoint 2 * DragS.OX - BasePoint(DragS.Number(1)).X, 2 * DragS.OY - BasePoint(DragS.Number(1)).Y, , vbInvert
        Paper.Refresh
    End If
    
'##############################
'SIMM POINT BY LINE
Case dsSimmPointByLine
    If DragS.NumOfMouseUps > 0 And IsPoint(DragS.Number(1)) Then
        If Not IsLine(AF) Then i_SetMousePointer curStateNo: PaperCls: ShowAll: EndOperation = True: GoTo EndOperation
        DragS.Number(2) = AF
        AddSimmPointByLine DragS.Number(1), DragS.Number(2)
        'If Paper.AutoRedraw <> DragS.OldAutoRedraw Then Paper.AutoRedraw = DragS.OldAutoRedraw
        ShowAll
        EndOperation = True
        GoTo EndOperation
    End If
    If DragS.NumOfMouseUps > 0 And DragS.Number(1) = 0 And IsLine(DragS.Number(2)) Then
        If IsPoint(AP) Then
            DragS.Number(1) = AP
            AddSimmPointByLine DragS.Number(1), DragS.Number(2)
        End If
        'If Paper.AutoRedraw <> DragS.OldAutoRedraw Then Paper.AutoRedraw = DragS.OldAutoRedraw
        PaperCls
        ShowAll
        EndOperation = True
    End If

'##############################
'INVERTED POINT
Case dsInvert
    If DragS.NumOfMouseUps > 0 And DragS.Number(1) <> 0 Then
        If Not IsCircle(AF) Then i_SetMousePointer curStateNo: PaperCls: ShowAll: EndOperation = True: GoTo EndOperation
        DragS.Number(2) = AF
        AddInvertedPoint DragS.Number(1), DragS.Number(2)
        'If Paper.AutoRedraw <> DragS.OldAutoRedraw Then Paper.AutoRedraw = DragS.OldAutoRedraw
        ShowAll
        EndOperation = True
    End If
    If DragS.NumOfMouseUps > 0 And DragS.Number(1) = 0 And IsCircle(DragS.Number(2)) Then
        If IsPoint(AP) Then
            DragS.Number(1) = AP
            AddInvertedPoint DragS.Number(1), DragS.Number(2)
        End If
        'If Paper.AutoRedraw <> DragS.OldAutoRedraw Then Paper.AutoRedraw = DragS.OldAutoRedraw
        PaperCls
        ShowAll
        EndOperation = True
    End If

'##############################
'INTERSECT
Case dsIntersect
    If DragS.ShouldComplete Then
        EndOperation = True
        GoTo EndOperation
    End If
    If DragS.NumOfMouseUps > 0 And IsFigure(DragS.Number(1)) Then
        If Not IsLine(AF) And Not IsCircle(AF) Or DragS.Number(1) = AF Then i_SetMousePointer curStateNo: EndOperation = True: GoTo EndOperation
        DragS.Number(2) = AF
        
        ShowSelectedFigure Paper.hDC, DragS.Number(1)
        
        'If Paper.AutoRedraw <> DragS.OldAutoRedraw Then Paper.AutoRedraw = DragS.OldAutoRedraw
        AddIntersectionPoints DragS.Number(1), DragS.Number(2)
        ShowAll
        EndOperation = True
    Else
        If IsFigure(DragS.Number(1)) Then ShowSelectedFigure Paper.hDC, DragS.Number(1)
        Paper.Refresh
    End If

'###############################
'POINT ON FIGURE
Case dsPointOnFigure
    EndOperation = True
    
'###############################
'MEASURE DISTANCE
Case dsMeasureDistance
    If DragS.ShouldComplete Then
        EndOperation = True
        GoTo EndOperation:
    End If
    Select Case DragS.NumOfMouseUps
        Case 1
            If IsPoint(AP) Then
                DragS.Number(2) = AP
                ShowSelectedPoint Paper.hDC, AP
            Else
                i_SetMousePointer curStateNo: EndOperation = True: GoTo EndOperation
            End If
            If DragS.Number(1) <> DragS.Number(2) Then AddMeasureDistance DragS.Number(1), DragS.Number(2)
            'If Paper.AutoRedraw <> DragS.OldAutoRedraw Then Paper.AutoRedraw = DragS.OldAutoRedraw
            ShowAll
            EndOperation = True
    End Select

'###############################
'MEASURE ANGLE
Case dsMeasureAngle
    Select Case DragS.NumOfMouseUps
        Case 1
            If IsPoint(AP) Then
                DragS.Number(2) = AP
                ShowSelectedPoint Paper.hDC, AP
                Paper.Refresh
            Else
                i_SetMousePointer curStateNo: EndOperation = True: GoTo EndOperation
            End If
        Case 2
            If IsPoint(AP) Then
                DragS.Number(3) = AP
                ShowSelectedPoint Paper.hDC, AP
                Paper.Refresh
            Else
                i_SetMousePointer curStateNo: EndOperation = True: GoTo EndOperation
            End If
            If DragS.Number(1) <> DragS.Number(2) And DragS.Number(2) <> DragS.Number(3) And DragS.Number(3) <> DragS.Number(1) Then AddMeasureAngle DragS.Number(1), DragS.Number(2), DragS.Number(3)
            'If Paper.AutoRedraw <> DragS.OldAutoRedraw Then Paper.AutoRedraw = DragS.OldAutoRedraw
            ShowAll
            EndOperation = True
    End Select

'###############################
'DYNAMIC LOCUS
Case dsDynamicLocus
    If DragS.NumOfMouseUps > 0 Then
        If IsPoint(AP) Then
            If BasePoint(AP).Type <> dsPoint Then
                If BasePoint(DragS.Number(1)).Type <> dsPointOnFigure Or (BasePoint(DragS.Number(1)).Type = dsPointOnFigure And BasePoint(AP).Type = dsPointOnFigure) Then
                    If DragS.Number(1) <> AP Then AddDynamicLocus DragS.Number(1), AP
                Else
                    If DragS.Number(1) <> AP Then AddDynamicLocus AP, DragS.Number(1)
                End If
                'If Paper.AutoRedraw <> DragS.OldAutoRedraw Then Paper.AutoRedraw = DragS.OldAutoRedraw
                PaperCls
                ShowAll
                'DrawingState = -1
                'i_CheckMainbarButton GetString(ResFigureBase + 2 * dsDynamicLocus), False
                EndOperation = True
            Else
                i_SetMousePointer curStateNo: EndOperation = True: GoTo EndOperation
            End If
        Else
            i_SetMousePointer curStateNo: EndOperation = True: GoTo EndOperation
        End If
    End If
    
'=========================================================
'                                       MEASURE AREA
'=========================================================

Case dsMeasureArea
    If DragS.ShouldComplete Then
        EndOperation = True
        GoTo EndOperation:
    End If
    
    DragS.NumberOfPoints = DragS.NumberOfPoints + 1
    ReDim Preserve DragS.Points(1 To DragS.NumberOfPoints)
    If IsPoint(AP) Then
        DragS.Points(DragS.NumberOfPoints) = AP
    Else
        AddBasePoint X, Y
        DragS.Points(DragS.NumberOfPoints) = PointCount
    End If
    
    ShowSelectedPoint Paper.hDC, DragS.Points(DragS.NumberOfPoints), True
    
    If DragS.NumberOfPoints = 1 Then
        DragS.ShouldHideFirstPointName = Not BasePoint(DragS.Points(1)).ShowName
        BasePoint(DragS.Points(1)).ShowName = True
        ShowSelectedPoint Paper.hDC, DragS.Points(DragS.NumberOfPoints), True
    End If
    If DragS.NumberOfPoints > 3 And (DragS.Points(1) = DragS.Points(DragS.NumberOfPoints)) Then
        BasePoint(DragS.Points(1)).ShowName = Not DragS.ShouldHideFirstPointName
        DragS.ShouldHideFirstPointName = False
        
        AddMeasureArea DragS.Points
        
        EndOperation = True
        GoTo EndOperation
    End If
    
'=====================================================
'                                           POLYGON
'=====================================================

Case dsPolygon
    DragS.NumberOfPoints = DragS.NumberOfPoints + 1
    ReDim Preserve DragS.Points(1 To DragS.NumberOfPoints)
    
    If IsPoint(AP) Then
        bPointAlreadySelected = False
        For Z = 2 To DragS.NumberOfPoints - 1
            If DragS.Points(Z) = AP Then bPointAlreadySelected = True
        Next
        
        If Not bPointAlreadySelected Then
            DragS.Points(DragS.NumberOfPoints) = AP
        Else
            AddBasePoint X, Y
            DragS.Points(DragS.NumberOfPoints) = PointCount
        End If
    
    Else 'AP is not a point
    
        AddBasePoint X, Y
        DragS.Points(DragS.NumberOfPoints) = PointCount
    End If
    
    ShowSelectedPoint Paper.hDC, DragS.Points(DragS.NumberOfPoints), True
    
    ReDim Preserve DragS.Pixels(1 To DragS.NumberOfPoints + 1)
    DragS.Pixels(DragS.NumberOfPoints).X = BasePoint(DragS.Points(DragS.NumberOfPoints)).PhysicalX
    DragS.Pixels(DragS.NumberOfPoints).Y = BasePoint(DragS.Points(DragS.NumberOfPoints)).PhysicalY
    
    If DragS.NumberOfPoints = 1 Then
        DragS.ShouldHideFirstPointName = Not BasePoint(DragS.Points(1)).ShowName
        BasePoint(DragS.Points(1)).ShowName = True
        DrawLine BasePoint(DragS.Points(1)).X, BasePoint(DragS.Points(1)).Y, X, Y, 0, 6, False
        ShowSelectedPoint Paper.hDC, DragS.Points(DragS.NumberOfPoints), True
    End If
    
    If DragS.NumberOfPoints = 2 Then
        DrawLine BasePoint(DragS.Points(1)).X, BasePoint(DragS.Points(1)).Y, DragS.OX, DragS.OY, 0, 6, False
        ShowSelectedPoint Paper.hDC, DragS.Points(DragS.NumberOfPoints)
        'DragS.Pixels(DragS.NumberOfPoints + 1).X = OX
        'DragS.Pixels(DragS.NumberOfPoints + 1).Y = OY
        'DrawPolygon DragS.Pixels(), setdefcolFigure, setdefcolFigureFill, 7, True
    End If
    
    If (DragS.NumberOfPoints > 3 And ((DragS.Points(1) = DragS.Points(DragS.NumberOfPoints)))) Or (DragS.NumberOfPoints > 2 And Button = 2) Then
        'ToPhysical DragS.OX, DragS.OY
        'DragS.Pixels(DragS.NumberOfPoints + 1).X = DragS.OX
        'DragS.Pixels(DragS.NumberOfPoints + 1).Y = DragS.OY
        'DrawPolygon DragS.Pixels(), setdefcolFigure, setdefcolFigureFill, 7, False
        
        BasePoint(DragS.Points(1)).ShowName = Not DragS.ShouldHideFirstPointName
        DragS.ShouldHideFirstPointName = False
        AddStaticGraphic sgPolygon, DragS.Points
        'i_CheckMainbarButton GetString(ResStaticObjectBase + 2 * sgPolygon), False
        
        EndOperation = True
        GoTo EndOperation
    End If
    
'    If DragS.NumberOfPoints > 2 And Button = 2 Then
'        BasePoint(DragS.Points(1)).ShowName = Not DragS.ShouldHideFirstPointName
'        DragS.ShouldHideFirstPointName = False
'        AddStaticGraphic sgPolygon, DragS.Points
'        'i_CheckMainbarButton GetString(ResStaticObjectBase + 2 * sgPolygon), False
'
'        EndOperation = True
'        GoTo EndOperation
'    End If
            
    DragS.OX = X
    DragS.OY = Y
    
End Select

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

EndOperation:
With DragS
    If EndOperation Then
        'If DrawingState > dsMeasureAngle Then
        '    i_SelectTool dsSelect '????? why do I need this??
        'End If
        .NumOfMouseDowns = 0
        .NumOfMouseUps = 0
        .ShouldComplete = False
        For Z = 1 To MaxDragNumbers: .Number(Z) = 0:  Next
        .Button = 0
        .Shift = Shift
        .OX = X
        .OY = Y
        .State = dscNormalState
        
        'i_CheckMainbarButton GetString(ResFigureBase + 2 * dsDynamicLocus), False
        'i_CheckMainbarButton GetString(ResStaticObjectBase + 2 * sgPolygon), False
        'If Paper.AutoRedraw <> setNoFlicker Then Paper.AutoRedraw = setNoFlicker
        i_CancelOperation
        
        .NumberOfFigures = 0
        .NumberOfPoints = 0
        ReDim .Points(1 To 1)
        
        PaperCls
        ShowAll
        
        If setToolSelectOnce Then
            i_SelectTool dsSelect
            'If setGroupTools Then
            '    i_SelectTool dsSelect
                'Menus(2).Items(TransposeInv(DrawingState, MenuTransposition)).Checked = False
                'DrawingState = dsSelect
                'Menus(2).Items(TransposeInv(DrawingState, MenuTransposition)).Checked = True
                'MenuBar(2).Refresh
            'Else
            '    Menus(2).Items(DrawingState + 2).Checked = False
            '    DrawingState = dsSelect
            '    Menus(2).Items(DrawingState + 2).Checked = True
            '    MenuBar(2).Refresh
            'End If
            i_SetMousePointer curStateArrow
            i_ShowStatus GetString(ResSelect)
        End If

        'i_SetMousePointer curStateArrow
        Select Case DrawingState
            Case dsPoint
                i_SetMousePointer curStateCross
            Case Else
                i_SetMousePointer curStateArrow
        End Select
    
    Else
    
        .NumOfMouseUps = .NumOfMouseUps + 1
        .Button = 0
        .Shift = Shift
        .OX = X
        .OY = Y
        .State = dscMovingState
    End If
End With
End Sub

'################################################################
'Process virtual paper repaint; never used; maintained for compatibility
'################################################################

Public Sub PaperPaint()
If Not Paper.AutoRedraw And DragS.State = dscNormalState Then ShowAll
End Sub

'################################################################
' Occurs when the virtual paper is resized;
' Updates the content of the paper
' Considers current mode
'################################################################

Public Sub PaperResize()
Dim Z As Long

If RestrictPaperResize Then Exit Sub

RefreshCanvasBorders
PaperCls
RecalcAllAuxInfo
ShowProperAll
'If DragS.State = dscSelectObjects Then
'    ShowSelectedAll TempObjectSelection
'ElseIf DragS.State = dscDemo Then
'    'DrawLinearObjectList Paper.hDC, DemoSequence
'    DrawSituation
'ElseIf DragS.State = dscMacroStateGivens Then
'    PaperCls
'    ShowAllWithGivens
'ElseIf DragS.State = dscMacroStateResults Then
'    PaperCls
'    ShowAllWithResults
'ElseIf DragS.State = dscMacroStateRun Then
'    PaperCls
'    ShowAll
'Else
'    ShowAll
'End If
End Sub


'=========================================================
'Process menu selection
'=========================================================

Public Sub MenuCommand(ByVal St As Long)
Dim GraphicName As String
Dim pAction As Action
Dim Z As Long



Select Case St

'=========================================================

Case ResExit
    Unload FormMain
    
'=========================================================

Case ResNew
    If IsDirty Then
        Select Case MsgBox(GetString(ResSave) & " " & DrawingName & "?", vbYesNoCancel + vbQuestion, GetString(ResMsgConfirmation))
            Case vbYes
                MenuCommand ResSave
            Case vbNo
                'do nothing
            Case vbCancel
                Exit Sub
        End Select
    End If
    ClearAll
    ClearPrivileges
    DrawingName = GetString(ResUntitled)
    FormMain.Caption = GetString(ResCaption) + " - " + GetString(ResUntitled)
    If FormMain.Enabled And FormMain.Visible Then FormMain.SetFocus
    
'=========================================================
   
Case ResOpen
    If IsDirty Then
        Select Case MsgBox(GetString(ResSave) & " " & DrawingName & "?", vbYesNoCancel + vbQuestion, GetString(ResMsgConfirmation))
            Case vbYes
                MenuCommand ResSave
            Case vbNo
                'do nothing
            Case vbCancel
                Exit Sub
        End Select
    End If
    
    CD.Filter = GetString(ResDGFFiles) & "|*." & extFIG
    CD.Flags = &H1000 + &H4
    If Not IsValidPath(LastFigurePath) Then LastFigurePath = ProgramPath
    CD.InitDir = LastFigurePath
    CD.DialogTitle = GetString(ResOpen)
    CD.ShowOpen
    If CD.Cancelled = True Or Dir(CD.FileName) = "" Then Exit Sub
    
    DrawingName = CD.FileName
    LastFigurePath = AddDirSep(RetrieveDir(DrawingName))
    If Not IsValidPath(LastFigurePath) Then LastFigurePath = ProgramPath
    
    i_ShowStatus GetString(ResWorkingPleaseWait)
    FormMain.Enabled = False
    Screen.MousePointer = vbHourglass
    FormMain.Caption = GetString(ResCaption) + " - " + RetrieveName(DrawingName)
    DoEvents
    
    ClearPrivileges
    OpenFile DrawingName
    
    i_ShowStatus
    AddMRUItem DrawingName
    Screen.MousePointer = vbDefault
    FormMain.Enabled = True
    If FormMain.Enabled And FormMain.Visible Then FormMain.SetFocus
    
'=========================================================
    
Case ResSave
    If DrawingName = GetString(ResUntitled) Then MenuCommand ResSaveAs: Exit Sub
    
    If Not CanSave Then
        MsgBox GetString(ResBuyDG), vbInformation
        Exit Sub
    End If
    
    If privNoAlter Then
        If MsgBox(GetString(ResPrivNoAlter), vbYesNo + vbQuestion, RetrieveName(DrawingName)) = vbYes Then
            MenuCommand ResSaveAs
        End If
        Exit Sub
    End If
    
    FormMain.Enabled = False
    i_ShowStatus GetString(ResWorkingPleaseWait)
    Screen.MousePointer = vbHourglass
    DoEvents
    SaveFile DrawingName
    i_ShowStatus
    Screen.MousePointer = vbDefault
    FormMain.Enabled = True
    If FormMain.Enabled And FormMain.Visible Then FormMain.SetFocus
    
'=========================================================
    
Case ResSaveAs
    Dim CantOverwrite As Boolean
    
    If Not CanSave Then
        MsgBox GetString(ResBuyDG), vbInformation
        Exit Sub
    End If
    
    Do
        CD.FileName = DrawingName
        CD.Filter = GetString(ResDGFFiles) & "|*." & extFIG
        CD.Flags = &H4 + &H2
        If Not IsValidPath(LastFigurePath) Then LastFigurePath = ProgramPath
        CD.InitDir = LastFigurePath
        CD.DialogTitle = GetString(ResSaveAs)
        CD.ShowSave
        If CD.Cancelled Then Exit Sub
        
        CantOverwrite = Not CanOverwriteFile(CD.FileName)
        If CantOverwrite Then
            If MsgBox(GetString(ResPrivNoAlter), vbYesNo + vbQuestion, RetrieveName(CD.FileName)) = vbNo Then Exit Sub
        End If
    Loop Until Not CantOverwrite
    
    DrawingName = CD.FileName
    If Right(UCase(DrawingName), 4) <> "." & extFIG Then
        If InStr(RetrieveName(DrawingName), ".") = 0 Then
            DrawingName = DrawingName & "." & extFIG
        Else
            DrawingName = Left(DrawingName, InStr(DrawingName, ".")) + extFIG
        End If
    End If
    
    LastFigurePath = AddDirSep(RetrieveDir(DrawingName))
    If Not IsValidPath(LastFigurePath) Then LastFigurePath = ProgramPath
    FormMain.Caption = GetString(ResCaption) + " - " + RetrieveName(DrawingName)
    FormMain.Enabled = False
    i_ShowStatus GetString(ResWorkingPleaseWait)
    Screen.MousePointer = vbHourglass
    DoEvents
    
    privNoAlter = False
    
    SaveFile DrawingName
    
    i_ShowStatus
    AddMRUItem DrawingName
    Screen.MousePointer = vbDefault
    FormMain.Enabled = True
    If FormMain.Enabled And FormMain.Visible Then FormMain.SetFocus
    
'=========================================================
    
Case ResBMP
    GraphicName = RetrieveName(DrawingName)
    GraphicName = Left(GraphicName, InStr(GraphicName, "."))
    If GraphicName = "" Then GraphicName = "Untitled."
    CD.FileName = GraphicName & extBMP
    CD.DialogTitle = GetString(ResSaveAs)
    CD.Filter = GetString(ResBMP) & " (*." & extBMP & ")|*." & extBMP
    CD.Flags = &H4 + &H2
    If Not IsValidPath(LastPicturePath) Then LastPicturePath = ProgramPath
    CD.InitDir = LastPicturePath
    CD.ShowSave
    If CD.Cancelled Then Exit Sub
    GraphicName = CD.FileName
    If Right(UCase(GraphicName), 4) <> "." & extBMP Then
        If InStr(RetrieveName(GraphicName), ".") = 0 Then
            GraphicName = GraphicName & "." & extBMP
        Else
            GraphicName = Left(GraphicName, InStr(GraphicName, ".") - 1) & "." & extBMP
        End If
    End If
    LastPicturePath = AddDirSep(RetrieveDir(GraphicName))
    If Not IsValidPath(LastPicturePath) Then LastPicturePath = ProgramPath
    FormMain.Enabled = False
    i_ShowStatus GetString(ResWorkingPleaseWait)
    Screen.MousePointer = vbHourglass
    DoEvents
    SaveBMP GraphicName
    i_ShowStatus
    Screen.MousePointer = vbDefault
    FormMain.Enabled = True
    If FormMain.Enabled And FormMain.Visible Then FormMain.SetFocus
    
'=========================================================
    
Case ResWMF
    GraphicName = RetrieveName(DrawingName)
    GraphicName = Left(GraphicName, InStr(GraphicName, "."))
    If GraphicName = "" Then GraphicName = "Untitled."
    CD.FileName = GraphicName & extWMF
    CD.Filter = GetString(ResWMF) & " (*." & extWMF & ")|*." & extWMF
    CD.Flags = &H4 + &H2
    CD.DialogTitle = GetString(ResSaveAs)
    If Not IsValidPath(LastPicturePath) Then LastPicturePath = ProgramPath
    CD.InitDir = LastPicturePath
    CD.ShowSave
    If CD.Cancelled Then Exit Sub
    GraphicName = CD.FileName
    If Right(UCase(GraphicName), 4) <> "." & extWMF Then
        If InStr(RetrieveName(GraphicName), ".") = 0 Then
            GraphicName = GraphicName & "." & extWMF
        Else
            GraphicName = Left(GraphicName, InStr(GraphicName, ".") - 1) & "." & extWMF
        End If
    End If
    LastPicturePath = AddDirSep(RetrieveDir(GraphicName))
    If Not IsValidPath(LastPicturePath) Then LastPicturePath = ProgramPath
    FormMain.Enabled = False
    i_ShowStatus GetString(ResWorkingPleaseWait)
    Screen.MousePointer = vbHourglass
    DoEvents
    SaveWMF GraphicName
    i_ShowStatus
    Screen.MousePointer = vbDefault
    FormMain.Enabled = True
    If FormMain.Enabled And FormMain.Visible Then FormMain.SetFocus
    
'=========================================================
    
Case ResEMF
    GraphicName = RetrieveName(DrawingName)
    GraphicName = Left(GraphicName, InStr(GraphicName, "."))
    If GraphicName = "" Then GraphicName = "Untitled."
    CD.FileName = GraphicName & extEMF
    CD.Filter = GetString(ResEMF) & " (*." & extEMF & ")|*." & extEMF
    CD.Flags = &H4 + &H2
    CD.DialogTitle = GetString(ResSaveAs)
    If Not IsValidPath(LastPicturePath) Then LastPicturePath = ProgramPath
    CD.InitDir = LastPicturePath
    CD.ShowSave
    If CD.Cancelled Then Exit Sub
    GraphicName = CD.FileName
    If Right(UCase(GraphicName), 4) <> "." & extEMF Then
        If InStr(RetrieveName(GraphicName), ".") = 0 Then
            GraphicName = GraphicName & "." & extEMF
        Else
            GraphicName = Left(GraphicName, InStr(GraphicName, ".") - 1) & "." & extEMF
        End If
    End If
    LastPicturePath = AddDirSep(RetrieveDir(GraphicName))
    If Not IsValidPath(LastPicturePath) Then LastPicturePath = ProgramPath
    FormMain.Enabled = False
    i_ShowStatus GetString(ResWorkingPleaseWait)
    Screen.MousePointer = vbHourglass
    DoEvents
    SaveEMF GraphicName
    i_ShowStatus
    Screen.MousePointer = vbDefault
    FormMain.Enabled = True
    If FormMain.Enabled And FormMain.Visible Then FormMain.SetFocus
    
'=========================================================

Case ResJSPHTML
    GraphicName = RetrieveName(DrawingName)
    GraphicName = Left(GraphicName, InStr(GraphicName, "."))
    If GraphicName = "" Then GraphicName = "Untitled."
    CD.FileName = GraphicName & extHTM
    CD.Filter = GetString(ResJSPHTML) & " (*." & extHTM & "; *.HTML)|*." & extHTM & ";*.HTML"
    CD.DialogTitle = GetString(ResSaveAs)
    CD.Flags = &H4 + &H2
    If Not IsValidPath(LastHTMLPath) Then LastHTMLPath = ProgramPath
    CD.InitDir = LastHTMLPath
    CD.ShowSave
    If CD.Cancelled Then Exit Sub
    GraphicName = CD.FileName
    If Right(LCase(GraphicName), 3) <> "htm" And Right(LCase(GraphicName), 4) <> "html" Then
        If InStr(RetrieveName(GraphicName), ".") = 0 Then
            GraphicName = GraphicName & "." & extHTM
        Else
            GraphicName = Left(GraphicName, InStrRev(GraphicName, ".") - 1) & "." & extHTM
        End If
    End If
    LastHTMLPath = AddDirSep(RetrieveDir(GraphicName))
    If Not IsValidPath(LastHTMLPath) Then LastHTMLPath = ProgramPath
    FormMain.Enabled = False
    i_ShowStatus GetString(ResWorkingPleaseWait)
    Screen.MousePointer = vbHourglass
    DoEvents
    SaveJSPHTML GraphicName
    i_ShowStatus
    Screen.MousePointer = vbDefault
    FormMain.Enabled = True
    If FormMain.Enabled And FormMain.Visible Then FormMain.SetFocus
    
'=========================================================
    
Case ResWEWindow
    FormMain.Docked.Visible = Not FormMain.Docked.Visible
    FormMain.Form_Resize
    FormMain.mnuWE.Checked = Not FormMain.mnuWE.Checked
    
'=========================================================

Case ResCalcBase
    frmCalculator.Show
    
'=========================================================
    
Case ResOptions
    frmSettings.ShowSettings
    
'=========================================================
    
Case ResAbout
    ShowAbout
    
'=========================================================

Case ResActiveAxes
    AddActiveAxes
    
'=========================================================

Case ResMnuAnCircle
    frmAnCircle.Show
    
'=========================================================

Case ResMnuAnLine
    frmAnLine.Show vbModal
    
'=========================================================

Case ResMnuAnPoint
    frmAnPoint.Show
    
'=========================================================
    
Case ResStaticObjectBase + sgBezier * 2
    FormMain.EnableMenus mnsFigureCreate
    FormMain.SelectTool dsSelect
    DragS.State = dscCreateStaticGraphic
    DragS.TypeOfStaticGraphic = sgBezier
    
'=========================================================
    
Case ResClearAll
    RecordGenericAction ResUndoClearAll
    'pAction.Type = actClearAll
    'MakeStructureSnapshot pAction
    'RecordAction pAction
    ClearAll False
    
'=========================================================

Case ResDemo
    RunDemo
    
'=========================================================

Case ResDemoOptions
    Dim DemoList As LinearObjectList
    DemoList = GenerateDemoSequence
    If DemoList.Count = 0 Then
        MsgBox GetString(ResEmptyDemo), vbExclamation
        Exit Sub
    End If
    frmDemoProps.Show
    
'=========================================================

Case ResFigureBase + 2 * dsDynamicLocus
'    FormMain.EnableMenus mnsFigureCreate
'    DragS.State = dscNormalState
'    DrawingState = dsDynamicLocus
'    i_CheckMainbarButton GetString(ResFigureBase + 2 * dsDynamicLocus), True
    FormMain.SelectTool dsDynamicLocus
    
'=========================================================
    
Case ResFigureList
    'FormMain.Enabled = False
    'frmFigureList.Show
    i_ShowFigureList
    
'=========================================================
    
Case ResFileProps
    frmFileProps.Show

'=========================================================
    
Case ResHelpContents
    If App.HelpFile <> "" Then
        CD.HelpFile = App.HelpFile
        CD.HelpCommand = HELP_CONTENTS
        CD.ShowHelp
    Else
        MsgBox GetString(ResHelpFileNotFound) & " " & ProgramPath, vbExclamation
    End If

'=========================================================
    
Case ResInsertButton
    ActiveButton = 0
    frmButtonProps.Show

'=========================================================

Case ResInsertLabel
    ActiveLabel = 0
    AddLabel

'=========================================================
'                                   MACROS
'=========================================================
    
Case ResMnuMacroCreate
    MacroCreateInit
    
'=========================================================
    
Case ResMnuMacroLoad
    LoadMacroAs
    i_NeedToSetFocus
    
'=========================================================

Case ResMnuMacroOrganize
    MacroOrganizeShow
    
'=========================================================

Case ResMnuMacroSelectResults
    MacroCreateResultsInit
    'MacroCreateResults
    
'=========================================================
    
Case ResMnuMacroSave
    frmMacroSave.Show vbModal
    'MacroCreateSave

'=========================================================

Case ResShowAxes
    ToggleAxes Not nShowAxes
'=========================================================

Case ResShowGrid
    nShowGrid = Not nShowGrid
    FormMain.mnuShowGrid.Checked = nShowGrid
    'setShowGrid = nShowGrid
    'SaveSetting AppName, "Interface", "ShowGrid", Format(-CInt(setShowGrid))
    PaperCls
    ShowAll
    
'=========================================================

Case ResPointList
    FormMain.Enabled = False
    frmPointList.Show
    
'=========================================================
    
Case ResStaticObjectBase + 2 * sgPolygon
    FormMain.EnableMenus mnsFigureCreate
    'i_SelectTool dsSelect
    DragS.State = dscCreateStaticGraphic
    DragS.TypeOfStaticGraphic = sgPolygon
    'i_CheckMainbarButton GetString(ResStaticObjectBase + 2 * sgPolygon), True

'=========================================================

Case ResPrint
    PrintFile
    
'=========================================================
    
Case ResUndo
    If DragS.ShouldSkipUndo Then DragS.ShouldSkipUndo = False: Exit Sub
    Undo

'=========================================================
    
Case ResRedo
    Redo
    
'=========================================================
    
Case ResTipOfTheDay
    frmTips.Show


'=========================================================
'=========================================================
'=========================================================
    
Case ResMnuLabelProperties
    frmLabelProps.Show
    
'=========================================================

Case ResMnuDeleteLabel
    DeleteLabel ActiveLabel
    
'=========================================================

Case ResMnuRecalcLabel
    ParseLabel ActiveLabel
    GetLabelSize ActiveLabel
    PaperCls
    ShowAll

'=========================================================
    
Case ResMnuDeleteFigure
    DeleteFigure ActiveFigure
    
'=========================================================
    
Case ResMnuFigureProperties
    frmFigureProps.Show
    
'=========================================================
    
Case ResMnuDeletePoint
    If BasePoint(ActivePoint).Type = dsPoint Then
        DeletePoint ActivePoint
    Else
        DeleteFigure BasePoint(ActivePoint).ParentFigure
    End If
    
'=========================================================
    
Case ResMnuChoosePoint
    BringPointToFront ActivePoint
    
'=========================================================
    
Case ResMnuChooseFigure
    BringFigureToFront ActiveFigure
    
'=========================================================
    
Case ResMnuPointProperties
    frmPointProps.Show
    
'=========================================================
    
Case ResMnuSnapToFigure
    Basepoint2PointOnFigure ActivePoint, ActiveFigure
    
'=========================================================
    
Case ResMnuReleasePoint
    PointOnFigure2Basepoint ActivePoint

'=========================================================

Case ResMnuMeasurementProperties
    frmMeasurementProps.Show vbModal
    
'=========================================================

Case ResDeleteButton
    DeleteButton ActiveButton
    ActiveButton = 0
    
'=========================================================

Case ResFix + ResButton + 1
    pAction.Type = actChangeAttrButton
    ReDim pAction.sButton(1 To 1)
    pAction.sButton(1) = Buttons(ActiveButton)
    pAction.pButton = ActiveButton
    RecordAction pAction
    
    Buttons(ActiveButton).Fixed = Not Buttons(ActiveButton).Fixed
    ActiveButton = 0
    
'=========================================================

Case ResButtonProperties
    frmButtonProps.Show
    
'=========================================================

Case ResDeleteLocus
    ClearLocusMenu ActivePoint
    
'=========================================================

Case ResCreateLocus
    CreateLocusMenu ActivePoint
    
'=========================================================

Case ResFix + ResLabel + 1
    ReDim pAction.sLabel(1 To 1)
    pAction.sLabel(1) = TextLabels(ActiveLabel)
    pAction.pLabel = ActiveLabel
    pAction.Type = actChangeAttrLabel
    RecordAction pAction
    TextLabels(ActiveLabel).Fixed = Not TextLabels(ActiveLabel).Fixed


'=========================================================
    
Case ResHide + ResFigure + 1
    pAction.Type = actHideFigure
    pAction.pFigure = ActiveFigure
    RecordAction pAction
    Figures(ActiveFigure).Hide = True
    PaperCls
    ShowAll

'=========================================================
    
Case ResHide + ResPoint + 1
    HidePoint ActivePoint

'=========================================================
    
Case ResStaticObjectBase + 2 * sgVector
    FormMain.EnableMenus mnsFigureCreate
    i_SelectTool dsSelect
    DragS.State = dscCreateStaticGraphic
    DragS.TypeOfStaticGraphic = sgVector

'=========================================================

Case ResStaticObjectBase + 1
    DeleteStaticGraphic ActiveStatic
    ActiveStatic = 0
    PaperCls
    ShowAll

'=========================================================
    
Case ResStaticObjectBase + 3
    frmStaticProps.Show

'=========================================================
    
Case ResShowName
    pAction.Type = actChangeAttrPoint
    ReDim pAction.sPoint(1 To 1)
    pAction.sPoint(1) = BasePoint(ActivePoint)
    pAction.pPoint = ActivePoint
    ShowPoint Paper.hDC, ActivePoint, True
    BasePoint(ActivePoint).ShowName = Not BasePoint(ActivePoint).ShowName
    ShowAll
    RecordAction pAction

'=========================================================
    
'=========================================================
    
'=========================================================
    
'=========================================================
    
End Select
End Sub

'
''=========================================================
''Process menu selection
''=========================================================
'
'Public Sub MenuCommand(ByVal St As String)
'Dim GraphicName As String
'St = Replace(St, "&", "")
'
'Select Case St
'Case GetString(ResExit)
'    Unload FormMain
'
'Case GetString(ResNew)
'    If IsDirty Then
'        Select Case MsgBox(GetString(ResSave) & " " & DrawingName & "?", vbYesNoCancel, GetString(ResMsgConfirmation))
'            Case vbYes
'                MenuCommand GetString(ResSave)
'            Case vbNo
'                'do nothing
'            Case vbCancel
'                Exit Sub
'        End Select
'    End If
'    ClearAll
'    DrawingName = GetString(ResUntitled)
'    FormMain.Caption = GetString(ResCaption) + " - " + GetString(ResUntitled)
'    If FormMain.Enabled And FormMain.Visible Then FormMain.SetFocus
'
'Case GetString(ResOpen)
'    If IsDirty Then
'        Select Case MsgBox(GetString(ResSave) & " " & DrawingName & "?", vbYesNoCancel, GetString(ResMsgConfirmation))
'            Case vbYes
'                MenuCommand GetString(ResSave)
'            Case vbNo
'                'do nothing
'            Case vbCancel
'                Exit Sub
'        End Select
'    End If
'    CD.Filter = GetString(ResMnuFigures) & " (*." & extFIG & ")|*." & extFIG
'    CD.Flags = &H1000 + &H4
'    If Not IsValidPath(LastFigurePath) Then LastFigurePath = ProgramPath
'    CD.InitDir = LastFigurePath
'    CD.DialogTitle = GetString(ResOpen)
'    CD.ShowOpen
'    If CD.Cancelled = True Or Dir(CD.FileName) = "" Then Exit Sub
'    DrawingName = CD.FileName
'    LastFigurePath = AddDirSep(RetrieveDir(DrawingName))
'    If Not IsValidPath(LastFigurePath) Then LastFigurePath = ProgramPath
'    i_ShowStatus GetString(ResWorkingPleaseWait)
'    FormMain.Enabled = False
'    Screen.MousePointer = vbHourglass
'    FormMain.Caption = GetString(ResCaption) + " - " + RetrieveName(DrawingName)
'    DoEvents
'    OpenFile DrawingName
'    i_ShowStatus
'    AddMRUItem DrawingName
'    Screen.MousePointer = vbDefault
'    FormMain.Enabled = True
'    If FormMain.Enabled And FormMain.Visible Then FormMain.SetFocus
'
'
'Case GetString(ResSave)
'    If DrawingName = GetString(ResUntitled) Then MenuCommand GetString(ResSaveAs): Exit Sub
'    FormMain.Enabled = False
'    i_ShowStatus GetString(ResWorkingPleaseWait)
'    Screen.MousePointer = vbHourglass
'    DoEvents
'    SaveFile DrawingName
'    i_ShowStatus
'    Screen.MousePointer = vbDefault
'    FormMain.Enabled = True
'    If FormMain.Enabled And FormMain.Visible Then FormMain.SetFocus
'
'Case GetString(ResSaveAs)
'    CD.FileName = DrawingName
'    CD.Filter = GetString(ResMnuFigures) & " (*." & extFIG & ")|*." & extFIG
'    CD.Flags = &H4 + &H2
'    If Not IsValidPath(LastFigurePath) Then LastFigurePath = ProgramPath
'    CD.InitDir = LastFigurePath
'    CD.DialogTitle = GetString(ResSaveAs)
'    CD.ShowSave
'    If CD.Cancelled Then Exit Sub
'    DrawingName = CD.FileName
'    If Right(UCase(DrawingName), 4) <> "." & extFIG Then
'        If InStr(RetrieveName(DrawingName), ".") = 0 Then
'            DrawingName = DrawingName & "." & extFIG
'        Else
'            DrawingName = Left(DrawingName, InStr(DrawingName, ".")) + extFIG
'        End If
'    End If
'    LastFigurePath = AddDirSep(RetrieveDir(DrawingName))
'    If Not IsValidPath(LastFigurePath) Then LastFigurePath = ProgramPath
'    FormMain.Caption = GetString(ResCaption) + " - " + RetrieveName(DrawingName)
'    FormMain.Enabled = False
'    i_ShowStatus GetString(ResWorkingPleaseWait)
'    Screen.MousePointer = vbHourglass
'    DoEvents
'    SaveFile DrawingName
'    i_ShowStatus
'    AddMRUItem DrawingName
'    Screen.MousePointer = vbDefault
'    FormMain.Enabled = True
'    If FormMain.Enabled And FormMain.Visible Then FormMain.SetFocus
'
'Case GetString(ResBMP)
'    GraphicName = RetrieveName(DrawingName)
'    GraphicName = Left(GraphicName, InStr(GraphicName, "."))
'    If GraphicName = "" Then GraphicName = "Untitled."
'    CD.FileName = GraphicName & extBMP
'    CD.Filter = GetString(ResBMP) & " (*." & extBMP & ")|*." & extBMP
'    CD.Flags = &H4 + &H2
'    If Not IsValidPath(LastPicturePath) Then LastPicturePath = ProgramPath
'    CD.InitDir = LastPicturePath
'    CD.ShowSave
'    If CD.Cancelled Then Exit Sub
'    GraphicName = CD.FileName
'    If Right(UCase(GraphicName), 4) <> "." & extBMP Then
'        If InStr(RetrieveName(GraphicName), ".") = 0 Then
'            GraphicName = GraphicName & "." & extBMP
'        Else
'            GraphicName = Left(GraphicName, InStr(GraphicName, ".") - 1) & "." & extBMP
'        End If
'    End If
'    LastPicturePath = AddDirSep(RetrieveDir(GraphicName))
'    If Not IsValidPath(LastPicturePath) Then LastPicturePath = ProgramPath
'    FormMain.Enabled = False
'    i_ShowStatus GetString(ResWorkingPleaseWait)
'    Screen.MousePointer = vbHourglass
'    DoEvents
'    SaveBMP GraphicName
'    i_ShowStatus
'    Screen.MousePointer = vbDefault
'    FormMain.Enabled = True
'    If FormMain.Enabled And FormMain.Visible Then FormMain.SetFocus
'
'Case GetString(ResWMF)
'    GraphicName = RetrieveName(DrawingName)
'    GraphicName = Left(GraphicName, InStr(GraphicName, "."))
'    If GraphicName = "" Then GraphicName = "Untitled."
'    CD.FileName = GraphicName & extWMF
'    CD.Filter = GetString(ResWMF) & " (*." & extWMF & ")|*." & extWMF
'    CD.Flags = &H4 + &H2
'    If Not IsValidPath(LastPicturePath) Then LastPicturePath = ProgramPath
'    CD.InitDir = LastPicturePath
'    CD.ShowSave
'    If CD.Cancelled Then Exit Sub
'    GraphicName = CD.FileName
'    If Right(UCase(GraphicName), 4) <> "." & extWMF Then
'        If InStr(RetrieveName(GraphicName), ".") = 0 Then
'            GraphicName = GraphicName & "." & extWMF
'        Else
'            GraphicName = Left(GraphicName, InStr(GraphicName, ".") - 1) & "." & extWMF
'        End If
'    End If
'    LastPicturePath = AddDirSep(RetrieveDir(GraphicName))
'    If Not IsValidPath(LastPicturePath) Then LastPicturePath = ProgramPath
'    FormMain.Enabled = False
'    i_ShowStatus GetString(ResWorkingPleaseWait)
'    Screen.MousePointer = vbHourglass
'    DoEvents
'    SaveWMF GraphicName
'    i_ShowStatus
'    Screen.MousePointer = vbDefault
'    FormMain.Enabled = True
'    If FormMain.Enabled And FormMain.Visible Then FormMain.SetFocus
'
'Case GetString(ResEMF)
'    GraphicName = RetrieveName(DrawingName)
'    GraphicName = Left(GraphicName, InStr(GraphicName, "."))
'    If GraphicName = "" Then GraphicName = "Untitled."
'    CD.FileName = GraphicName & extEMF
'    CD.Filter = GetString(ResEMF) & " (*." & extEMF & ")|*." & extEMF
'    CD.Flags = &H4 + &H2
'    If Not IsValidPath(LastPicturePath) Then LastPicturePath = ProgramPath
'    CD.InitDir = LastPicturePath
'    CD.ShowSave
'    If CD.Cancelled Then Exit Sub
'    GraphicName = CD.FileName
'    If Right(UCase(GraphicName), 4) <> "." & extEMF Then
'        If InStr(RetrieveName(GraphicName), ".") = 0 Then
'            GraphicName = GraphicName & "." & extEMF
'        Else
'            GraphicName = Left(GraphicName, InStr(GraphicName, ".") - 1) & "." & extEMF
'        End If
'    End If
'    LastPicturePath = AddDirSep(RetrieveDir(GraphicName))
'    If Not IsValidPath(LastPicturePath) Then LastPicturePath = ProgramPath
'    FormMain.Enabled = False
'    i_ShowStatus GetString(ResWorkingPleaseWait)
'    Screen.MousePointer = vbHourglass
'    DoEvents
'    SaveEMF GraphicName
'    i_ShowStatus
'    Screen.MousePointer = vbDefault
'    FormMain.Enabled = True
'    If FormMain.Enabled And FormMain.Visible Then FormMain.SetFocus
'
'Case GetString(ResJSPHTML)
'    GraphicName = RetrieveName(DrawingName)
'    GraphicName = Left(GraphicName, InStr(GraphicName, "."))
'    If GraphicName = "" Then GraphicName = "Untitled."
'    CD.FileName = GraphicName & extHTM
'    CD.Filter = GetString(ResJSPHTML) & " (*." & extHTM & "; *.HTML)|*." & extHTM & ";*.HTML"
'    CD.Flags = &H4 + &H2
'    If Not IsValidPath(LastHTMLPath) Then LastHTMLPath = ProgramPath
'    CD.InitDir = LastHTMLPath
'    CD.ShowSave
'    If CD.Cancelled Then Exit Sub
'    GraphicName = CD.FileName
'    If Right(LCase(GraphicName), 3) <> "htm" And Right(LCase(GraphicName), 4) <> "html" Then
'        If InStr(RetrieveName(GraphicName), ".") = 0 Then
'            GraphicName = GraphicName & "." & extHTM
'        Else
'            GraphicName = Left(GraphicName, InStrRev(GraphicName, ".") - 1) & "." & extHTM
'        End If
'    End If
'    LastHTMLPath = AddDirSep(RetrieveDir(GraphicName))
'    If Not IsValidPath(LastHTMLPath) Then LastHTMLPath = ProgramPath
'    FormMain.Enabled = False
'    i_ShowStatus GetString(ResWorkingPleaseWait)
'    Screen.MousePointer = vbHourglass
'    DoEvents
'    SaveJSPHTML GraphicName
'    i_ShowStatus
'    Screen.MousePointer = vbDefault
'    FormMain.Enabled = True
'    If FormMain.Enabled And FormMain.Visible Then FormMain.SetFocus
'
'Case GetString(ResWEWindow)
'    FormMain.Docked.Visible = Not FormMain.Docked.Visible
'    FormMain.Form_Resize
'    FormMain.mnuWE.Checked = Not FormMain.mnuWE.Checked
'
'Case GetString(ResOptions)
'    frmSettings.Show
'
'Case GetString(ResMnuLabelProperties)
'    frmLabelProps.Show
'
'Case GetString(ResMnuDeleteLabel)
'    DeleteLabel ActiveLabel
'
'Case GetString(ResMnuDeleteFigure)
'    DeleteFigure ActiveFigure
'
'Case GetString(ResMnuFigureProperties)
'    frmFigureProps.Show
'
'Case GetString(ResMnuDeletePoint)
'    If BasePoint(ActivePoint).Type = dsPoint Then
'        DeletePoint ActivePoint
'    Else
'        DeleteFigure BasePoint(ActivePoint).ParentFigure
'    End If
'
'Case GetString(ResMnuChoosePoint)
'    BringPointToFront ActivePoint
'
'Case GetString(ResMnuChooseFigure)
'    BringFigureToFront ActiveFigure
'
'Case GetString(ResMnuPointProperties)
'    frmPointProps.Show
'
'Case GetString(ResMnuSnapToFigure)
'    Basepoint2PointOnFigure ActivePoint, ActiveFigure
'
'Case GetString(ResMnuReleasePoint)
'    PointOnFigure2Basepoint ActivePoint
'
'End Select
'End Sub

'===================================================
'Cancel current virtual paper operation
'===================================================

Public Sub CancelOperation()
Dim Z As Long

'===================================================

If DragS.State = dscNormalState Then Exit Sub

'===================================================

If DragS.State = dscDemo Then
    EndDemo
    Exit Sub
End If

'===================================================

If DragS.State = dscSelectObjects Then
    ObjectSelectionCancel
    Exit Sub
End If

'===================================================

If DragS.State = dscSelectObjects Then
    ObjectSelectionCancel
    Exit Sub
End If

'===================================================

If DragS.State = dscMacroStateGivens Or DragS.State = dscMacroStateResults Then
    i_CancelMacro
    Exit Sub
End If

'===================================================

If DragS.State = dscMacroStateRun Then
    i_ExitMacroRunMode
    Exit Sub
End If

'===================================================

'i_CheckMainbarButton GetString(ResFigureBase + 2 * dsDynamicLocus), False
'i_CheckMainbarButton GetString(ResStaticObjectBase + 2 * sgPolygon), False
With DragS
    'If DrawingState > dsMeasureAngle Then i_SelectTool dsSelect
    If .ShouldHideFirstPointName Then
        If IsPoint(DragS.Points(1)) Then
            BasePoint(DragS.Points(1)).ShowName = False
        End If
    End If
    .NumOfMouseDowns = 0
    .NumOfMouseUps = 0
    .ShouldComplete = False
    ReDim .Figures(1 To 1)
    .NumberOfFigures = 0
    .NumberOfPoints = 0
    ReDim .Points(1 To 1)
    
    For Z = 1 To MaxDragNumbers: .Number(Z) = 0: Next
    .Button = 0
    .Shift = 0
    .OX = 0
    .OY = 0
    .State = dscNormalState
    i_CancelOperation
End With
PaperCls
ShowAll
PaperMouseMove 0, 0, 0, 0
End Sub

'Public Sub RenamePointDecision(ByVal Choice As Long, Optional ByVal NewName As String)
'
'End Sub

Public Sub ShortcutKeyClick(ByVal KeyCode As Integer)
FormMain.MenuBar(2).ShortcutKeyClick KeyCode
End Sub

