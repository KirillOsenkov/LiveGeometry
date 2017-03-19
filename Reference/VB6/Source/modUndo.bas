Attribute VB_Name = "modUndo"
Option Explicit

'Warning! The ORDER of these enums is important!
Public Enum ActionType
    actAddPoint
    actAddFigure
    actAddLabel
    actAddLocus
    actAddWE
    actMovePoint
    actMovePointLabel
    actMoveLabel
    actChangeAttrPoint
    actChangeAttrFigure
    actChangeAttrLabel
    actChangeAttrLocus
    actChangeWE
    actRemovePoint
    actRemoveFigure
    actRemoveLabel
    actRemoveLocus
    actRemoveWE
    actPointZOrder
    actFigureZOrder
    actHidePoint
    actHideFigure
    actMacro
    actSnapPoint
    actReleasePoint
    actClearAll
    actDeleteMacro
    actDeleteObjects
    actCreateObjects
    actMovePointBack
    actRenamePoint
    actAddSG
    actChangeAttrSG
    actRemoveSG
    actAddButton
    actRemoveButton
    actMoveButton
    actChangeAttrButton
    actApplyButton
    actShowHideObjects
    actGroupAction
    actMoveMeasure
    actGenericAction ' action that needs *StructureSnapshot
    actGenericReaction
End Enum

Public Type Action
    Type As ActionType
    UndoType As ActionType
    
    UndoString As String
    RedoString As String
    
    EventTime As Date
    Group As Long
    CurrentTransform As TransformType
    
    pPoint As Long
    pFigure As Long
    pLabel As Long
    pLocus As Long
    pWE As Long
    pSG As Long
    pButton As Long
    sPoint() As BasePointType
    sFigure() As Figure
    sLabel() As TextLabel
    sLocus() As Locus
    sSG() As StaticGraphic
    sWE() As WatchExpression
    sButton() As Button
    
    AuxPoints() As OnePoint
    AuxInfo() As Double
    
    cObjectList As ObjectList
    
End Type

Public Sub Undo(Optional ByVal InAGroup As Boolean = False)
On Local Error GoTo EH:

Dim Z As Long
Dim pAct As Action, AType As ActionType, pAction As Action
Dim WasFromRedo As Boolean

If ActivityCount <= 0 Then Exit Sub

pAct = Activity(ActivityCount)
'If Not InAGroup Then ClearAllTags

WasFromRedo = FromRedo
FromRedo = True

With pAct 'Activity(ActivityCount)
    AType = .Type
    Select Case AType
        Case actGenericAction
            pAction.Type = actGenericReaction
            MakeStructureSnapshot pAction
            pAction.UndoString = pAct.UndoString
            pAction.RedoString = pAct.RedoString
            RecordAction pAction
            
            RestoreFromStructureSnapshot pAct
        Case actGenericReaction
            pAction.Type = actGenericAction
            MakeStructureSnapshot pAction
            pAction.UndoString = pAct.UndoString
            pAction.RedoString = pAct.RedoString
            RecordAction pAction
            
            RestoreFromStructureSnapshot pAct
        Case actAddFigure
            DeleteFigure .pFigure
        Case actAddLabel
            DeleteLabel .pLabel
        Case actAddButton
            DeleteButton .pButton
        Case actAddLocus
            If BasePoint(.pPoint).Locus <> 0 Then
                pAction.Type = actRemoveLocus
                ReDim pAction.sLocus(1 To 1)
                pAction.sLocus(1) = Locuses(BasePoint(.pPoint).Locus)
                pAction.pPoint = .pPoint
                RecordAction pAction
                DeleteLocus .pPoint
            End If
        Case actAddPoint
            DeletePoint .pPoint
        'Case actAddWE
            'RemoveWatchExpression .pWE
'        Case actApplyButton
'            pAction.Type = actApplyButton
'            MakeStructureSnapshot pAction
'            RecordAction pAction
'
'            RestoreFromStructureSnapshot pAct
        Case actChangeAttrFigure
            pAction.Type = actChangeAttrFigure
            ReDim pAction.sFigure(1 To 1)
            pAction.sFigure(1) = Figures(.pFigure)
            pAction.pFigure = .pFigure
            RecordAction pAction
            Figures(.pFigure) = .sFigure(1)
        Case actChangeAttrLabel
            pAction.Type = actChangeAttrLabel
            ReDim pAction.sLabel(1 To 1)
            pAction.sLabel(1) = TextLabels(.pLabel)
            pAction.pLabel = .pLabel
            RecordAction pAction
            TextLabels(.pLabel) = .sLabel(1)
        Case actChangeAttrButton
            pAction.Type = actChangeAttrButton
            ReDim pAction.sButton(1 To 1)
            pAction.sButton(1) = Buttons(.pButton)
            pAction.pButton = .pButton
            RecordAction pAction
            Buttons(.pButton) = .sButton(1)
        Case actChangeAttrLocus
            pAction.Type = actChangeAttrLocus
            ReDim pAction.sLocus(1 To 1)
            pAction.sLocus(1) = Locuses(.pLocus)
            pAction.pLocus = .pLocus
            RecordAction pAction
            Locuses(.pLocus) = .sLocus(1)
        Case actChangeAttrPoint
            pAction.Type = actChangeAttrPoint
            ReDim pAction.sPoint(1 To 1)
            pAction.sPoint(1) = BasePoint(.pPoint)
            pAction.pPoint = .pPoint
            RecordAction pAction
            BasePoint(.pPoint) = .sPoint(1)
        Case actChangeWE
            pAction.Type = actChangeWE
            ReDim pAction.sWE(1 To 1)
            pAction.sWE(1) = WatchExpressions(.pWE)
            pAction.pWE = .pWE
            RecordAction pAction
            WatchExpressions(.pWE) = .sWE(1)
        Case actFigureZOrder
            'Swap Figures(.pFigure).ZOrder, Figures(.AuxInfo(1)).ZOrder
            pAction.Type = actFigureZOrder
            ReDim pAction.AuxInfo(1 To 2)
            pAction.AuxInfo(1) = .AuxInfo(1)
            pAction.AuxInfo(2) = Figures(.AuxInfo(1)).ZOrder
            RecordAction pAction
            
            For Z = 0 To FigureCount - 1
                If Figures(Z).ZOrder >= .AuxInfo(2) Then Figures(Z).ZOrder = Figures(Z).ZOrder + 1
            Next Z
            Figures(.AuxInfo(1)).ZOrder = .AuxInfo(2)
'        Case actGroupAction
'            pAction.Type = actGroupAction
'            MakeStructureSnapshot pAction
'            RecordAction pAction
'
'            RestoreFromStructureSnapshot pAct
        Case actHidePoint
            pAction.Type = actChangeAttrPoint
            ReDim pAction.sPoint(1 To 1)
            pAction.sPoint(1) = BasePoint(.pPoint)
            pAction.pPoint = .pPoint
            RecordAction pAction
            BasePoint(.pPoint).Hide = False
            BasePoint(.pPoint).Enabled = True
        Case actHideFigure
            pAction.Type = actChangeAttrFigure
            ReDim pAction.sFigure(1 To 1)
            pAction.sFigure(1) = Figures(.pFigure)
            pAction.pFigure = .pFigure
            RecordAction pAction
            'Figures(.pFigure) = .sFigure(1)
            
            Figures(.pFigure).Hide = False
'        Case actMacro
'            pAction.Type = actDeleteMacro
'            MakeStructureSnapshot pAction
'            RecordAction pAction
'
'            RestoreFromStructureSnapshot pAct
'            For Z = 1 To .AuxInfo(2)
'                DeleteFigure FigureCount - 1, False, False
'            Next
'
'            For Z = 1 To .AuxInfo(3)
'                DeleteStaticGraphic StaticGraphicCount, False
'            Next
        Case actMoveLabel
            pAction.Type = actMoveLabel
            ReDim pAction.AuxPoints(1 To 1)
            pAction.pLabel = .pLabel
            pAction.AuxPoints(1).X = TextLabels(.pLabel).LogicalPosition.P1.X
            pAction.AuxPoints(1).Y = TextLabels(.pLabel).LogicalPosition.P1.Y
            RecordAction pAction
            
            MoveLabel .pLabel, .AuxPoints(1).X, .AuxPoints(1).Y
        Case actMoveButton
            pAction.Type = actMoveButton
            ReDim pAction.AuxPoints(1 To 1)
            pAction.pButton = .pButton
            pAction.AuxPoints(1).X = Buttons(.pButton).LogicalPosition.P1.X
            pAction.AuxPoints(1).Y = Buttons(.pButton).LogicalPosition.P1.Y
            RecordAction pAction
            
            MoveButton .pButton, .AuxPoints(1).X, .AuxPoints(1).Y
        Case actMovePoint
            pAction.Type = actMovePointBack
            ReDim pAction.AuxPoints(1 To 1)
            pAction.AuxPoints(1).X = BasePoint(.pPoint).X
            pAction.AuxPoints(1).Y = BasePoint(.pPoint).Y
            If BasePoint(.pPoint).Locus > 0 Then
                pAction.pLocus = BasePoint(.pPoint).Locus
                ReDim pAction.sLocus(1 To 1)
                pAction.sLocus(1) = Locuses(BasePoint(.pPoint).Locus)
            End If
            pAction.pPoint = .pPoint
            RecordAction pAction
            
            If .AuxInfo(2) = 1 Then
                Do While Locuses(BasePoint(.pPoint).Locus).LocusPointCount > .AuxInfo(1)
                    RemovePointFromLocus BasePoint(.pPoint).Locus
                Loop
            End If
            If .AuxInfo(4) = 1 Then
                Figures(.pFigure).AuxInfo(1) = .AuxInfo(3)
            Else
                MovePoint .pPoint, .AuxPoints(1).X, .AuxPoints(1).Y
                RemovePointFromLocus BasePoint(.pPoint).Locus
            End If
            RecalcAllAuxInfo
            'If WECount > 0 Then FormMain.ValueTable1.UpdateExpressions
            If LabelCount > 0 Then UpdateLabels
        Case actMovePointBack
            pAction.Type = actMovePoint
            ReDim pAction.AuxPoints(1 To 1)
            pAction.AuxPoints(1).X = BasePoint(.pPoint).X
            pAction.AuxPoints(1).Y = BasePoint(.pPoint).Y
            pAction.pPoint = .pPoint
            ReDim pAction.AuxInfo(1 To 4)
            If BasePoint(.pPoint).Type = dsPointOnFigure Then
                pAction.AuxInfo(4) = 1
                pAction.AuxInfo(3) = Figures(BasePoint(.pPoint).ParentFigure).AuxInfo(1)
                pAction.pFigure = BasePoint(.pPoint).ParentFigure
            End If
            If BasePoint(.pPoint).Locus <> 0 Then
                If Locuses(BasePoint(.pPoint).Locus).Enabled Then
                    pAction.AuxInfo(1) = Locuses(BasePoint(.pPoint).Locus).LocusPointCount
                    pAction.AuxInfo(2) = 1
                End If
            End If
            RecordAction pAction
            
            If .pLocus > 0 Then Locuses(.pLocus) = .sLocus(1)
            RemovePointFromLocus .pLocus
            MovePoint .pPoint, .AuxPoints(1).X, .AuxPoints(1).Y
            RecalcAllAuxInfo
            'If WECount > 0 Then FormMain.ValueTable1.UpdateExpressions
            If LabelCount > 0 Then UpdateLabels
        Case actMovePointLabel
            pAction.Type = actMovePointLabel
            pAction.pPoint = .pPoint
            ReDim pAction.AuxPoints(1 To 1)
            pAction.AuxPoints(1).X = BasePoint(.pPoint).LabelOffsetX
            pAction.AuxPoints(1).Y = BasePoint(.pPoint).LabelOffsetY
            RecordAction pAction
            
            BasePoint(.pPoint).LabelOffsetX = .AuxPoints(1).X
            BasePoint(.pPoint).LabelOffsetY = .AuxPoints(1).Y
        Case actPointZOrder
            pAction.Type = actPointZOrder
            ReDim pAction.AuxInfo(1 To 2)
            pAction.AuxInfo(1) = .AuxInfo(1)
            pAction.AuxInfo(2) = BasePoint(.AuxInfo(1)).ZOrder
            RecordAction pAction

            For Z = 1 To PointCount
                If BasePoint(Z).ZOrder >= .AuxInfo(2) Then BasePoint(Z).ZOrder = BasePoint(Z).ZOrder + 1
            Next Z
            BasePoint(.AuxInfo(1)).ZOrder = .AuxInfo(2)
        Case actReleasePoint
            Basepoint2PointOnFigure .pPoint, .pFigure
        Case actRemoveLabel
            pAction.Type = actAddLabel
            pAction.pLabel = .AuxInfo(1)
            RecordAction pAction
            
            RestoreFromStructureSnapshot pAct
        Case actRemoveButton
            pAction.Type = actAddButton
            pAction.pButton = .AuxInfo(1)
            RecordAction pAction
            
            RestoreFromStructureSnapshot pAct
        Case actRemoveLocus
            If .pLocus = 0 Then
                pAction.Type = actAddLocus
                AddLocus .pPoint
                pAction.pPoint = .pPoint
                RecordAction pAction
            Else
                pAction.Type = actChangeAttrLocus
                ReDim pAction.sLocus(1 To 1)
                pAction.sLocus(1) = Locuses(.pLocus)
                pAction.pLocus = .pLocus
                RecordAction pAction
                
                Locuses(.pLocus) = .sLocus(1)
            End If
'        Case actRemoveWE
'            pAction.Type = actCreateObjects
'            MakeStructureSnapshot pAction
'            RecordAction pAction
'
'            RestoreFromStructureSnapshot pAct
'        Case actRenamePoint
'            pAction.Type = actRenamePoint
'            MakeStructureSnapshot pAction
'            RecordAction pAction
'
'            RestoreFromStructureSnapshot pAct
        Case actSnapPoint
            PointOnFigure2Basepoint .pPoint
'        Case actRemoveFigure
'            pAction.Type = actCreateObjects
'            MakeStructureSnapshot pAction
'            RecordAction pAction
'
'            RestoreFromStructureSnapshot pAct
'        Case actRemovePoint
'            pAction.Type = actCreateObjects
'            MakeStructureSnapshot pAction
'            RecordAction pAction
'
'            RestoreFromStructureSnapshot pAct
'        Case actClearAll
'            pAction.Type = actCreateObjects
'            MakeStructureSnapshot pAction
'            RecordAction pAction
'
'            RestoreFromStructureSnapshot pAct
        
        Case actShowHideObjects
            pAction.Type = actShowHideObjects
            MakeStructureSnapshot pAction
            RecordAction pAction
            
            RestoreFromStructureSnapshot pAct
        
'        Case actDeleteMacro
'            pAction.Type = actCreateObjects
'            MakeStructureSnapshot pAction
'            RecordAction pAction
'
'            RestoreFromStructureSnapshot pAct
'        Case actDeleteObjects
'            pAction.Type = actCreateObjects
'            MakeStructureSnapshot pAction
'            RecordAction pAction
'
'            RestoreFromStructureSnapshot pAct
'        Case actCreateObjects
'            pAction.Type = actDeleteObjects
'            MakeStructureSnapshot pAction
'            RecordAction pAction
'
'            RestoreFromStructureSnapshot pAct
        Case actAddSG
            DeleteStaticGraphic .pSG
            For Z = 1 To .pFigure
                DeleteFigure .sFigure(1).Tag, False, False
            Next
        Case actChangeAttrSG
            pAction.Type = actChangeAttrSG
            pAction.pSG = .pSG
            ReDim pAction.sSG(1 To 1)
            pAction.sSG(1) = StaticGraphics(.pSG)
            RecordAction pAction
            
            StaticGraphics(.pSG) = .sSG(1)
        Case actRemoveSG
            pAction.Type = actAddSG
            pAction.pSG = .AuxInfo(1)
            RecordAction pAction
            
            RestoreFromStructureSnapshot pAct
        End Select
End With

'======================================================

Activity(ActivityCount).UndoType = Activity(ActivityCount - 1).Type
Activity(ActivityCount).UndoString = pAct.UndoString
Activity(ActivityCount).RedoString = pAct.RedoString
RecordUndoneAction Activity(ActivityCount)

ActivityCount = ActivityCount - 2
If ActivityCount > 0 Then ReDim Preserve Activity(1 To ActivityCount) Else ReDim Activity(1 To 1)
If ActivityCount > 0 And Not InAGroup Then
    'FormMain.mnuUndo.Caption = GetString(ResUndo) & IIf(setLanguage = langGerman, ":", "") & " " & GetString(ResUndoActionBase + 2 * Activity(ActivityCount).Type)
    FormMain.mnuUndo.Caption = Activity(ActivityCount).UndoString
Else
    FormMain.mnuUndo.Caption = GetString(ResUndo)
    FormMain.mnuUndo.Enabled = False
End If

If Not WasFromRedo Then
    'FormMain.mnuRedo.Caption = GetString(ResRedo) & IIf(setLanguage = langGerman, ":", "") & " " & GetString(ResUndoActionBase + 2 * pAct.Type)
    FormMain.mnuRedo.Caption = pAct.RedoString
End If

FromRedo = False

'If pAct.Group > 0 Then
'    For Z = 1 To pAct.Group
'        Undo True
'    Next
'End If

If Not InAGroup Then
    'ClearAllTags
    RecalcScrollAssociatedInfo
    RecalcAllAuxInfo
    PaperCls
    ShowAll
End If

EH:
End Sub

Public Sub Redo()
On Local Error GoTo EH:

Dim Z As Long
Dim pAct As Action, pAction As Action

If UndoneActivityCount <= 0 Then Exit Sub

FromRedo = True

pAct = UndoneActivity(UndoneActivityCount)
UndoneActivityCount = UndoneActivityCount - 1
If UndoneActivityCount > 0 Then ReDim Preserve UndoneActivity(1 To UndoneActivityCount) Else ReDim UndoneActivity(1 To 1)
RecordAction pAct

Undo

FromRedo = True

pAct = UndoneActivity(UndoneActivityCount)
UndoneActivityCount = UndoneActivityCount - 1
If UndoneActivityCount > 0 Then ReDim Preserve UndoneActivity(1 To UndoneActivityCount) Else ReDim UndoneActivity(1 To 1)
RecordAction pAct

If UndoneActivityCount > 0 Then
    FormMain.mnuRedo.Caption = UndoneActivity(UndoneActivityCount).RedoString
    'Debug.Print UndoneActivity(UndoneActivityCount).RedoString
    FormMain.mnuRedo.Enabled = True
Else
    FormMain.mnuRedo.Caption = GetString(ResRedo)
    FormMain.mnuRedo.Enabled = False
End If

FromRedo = False

EH:
End Sub

Public Sub MakeStructureSnapshot(pAction As Action)
pAction.sFigure = Figures
pAction.sPoint = BasePoint
pAction.sLocus = Locuses
pAction.sWE = WatchExpressions
pAction.sSG = StaticGraphics
pAction.sLabel = TextLabels
pAction.sButton = Buttons
pAction.pFigure = FigureCount
pAction.pPoint = PointCount
pAction.pLocus = LocusCount
pAction.pWE = WECount
pAction.pLabel = LabelCount
pAction.pSG = StaticGraphicCount
pAction.pButton = ButtonCount
'pAction.CurrentTransform = WorldTransform
End Sub
'
'Public Sub SnapshotDeleteFigure(ByVal Index As Long)
'Dim pAction As Action
'pAction.Type = actRemoveFigure
'pAction.pFigure = Index
'ReDim pAction.sFigure(1 To 1)
'pAction.sFigure(1) = Figures(Index)
'pAction.Group = 0
'
'For Z = 0 To Figures(Index).NumberOfChildren - 1
'    SnapshotDeleteFigure Figures(Index).Children(Z)
'    pAction.Group = pAction.Group + 1
'Next
'
'Select Case Figures(Index).FigureType
'    Case dsCircle_ArcCenterAndRadiusAndTwoPoints
'        SnapshotDeletePoint Figures(Index).Points(5)
'        SnapshotDeletePoint Figures(Index).Points(6)
'        pAction.Group = pAction.Group + 2
'    Case dsIntersect
'        SnapshotDeletePoint Figures(Index).Points(0)
'        SnapshotDeletePoint Figures(Index).Points(1)
'        pAction.Group = pAction.Group + 2
'    Case dsMiddlePoint, dsSimmPoint, dsSimmPointByLine, dsPointOnFigure, dsInvert
'        SnapshotDeletePoint Figures(Index).Points(0)
'        pAction.Group = pAction.Group + 1
'End Select
'
''Z = 0
''Do While Z <= FigureCount - 1
''    N = GetProperParentNumber(Figures(Z).FigureType)
''    If N > 0 Then
''        For Q = 0 To N - 1
''            If Figures(Z).Parents(Q) = Index Then
''                DeleteFigure Z, False, ShouldRecord
''                pAction.Group = pAction.Group + 1
''                GoTo NextZ
''            End If
''        Next Q
''    End If
''Z = Z + 1
''NextZ:
''Loop
'
'RecordAction pAction
'End Sub
'
'Public Sub SnapshotDeletePoint(ByVal Index As Long)
'Dim pAction As Action
'pAction.Type = actRemovePoint
'pAction.pPoint = Index
'ReDim pAction.sPoint(1 To 1)
'pAction.sPoint(1) = BasePoint(Index)
'pAction.Group = 0
'If BasePoint(Index).Locus <> 0 Then
'    ReDim pAction.sLocus(1 To 1)
'    pAction.sLocus(1) = Locuses(BasePoint(Index).Locus)
'    pAction.pLocus = BasePoint(Index).Locus
'End If
'
'For Z = 0 To FigureCount - 1
'    If IsParentPoint(Figures(Z), Index) Then SnapshotDeleteFigure Z
'    pAction.Group = pAction.Group + 1
'Next
'
'RecordAction pAction
'End Sub

Public Sub RecordAction(pAction As Action, Optional ByVal ShouldAddStrings As Boolean = True)
pAction.EventTime = Now
If pAction.Type = actGenericAction Or pAction.Type = actGenericReaction Then
    'do nothing
Else
    pAction.UndoString = GetString(ResUndoActionBase + 2 * pAction.Type)
    pAction.RedoString = Replace(pAction.UndoString, GetString(ResUndo), GetString(ResRedo))
End If

ActivityCount = ActivityCount + 1
ReDim Preserve Activity(1 To ActivityCount)
Activity(ActivityCount) = pAction

i_UndoAdded pAction.UndoString

If Not FromRedo And UndoneActivityCount > 0 Then
    ClearRedoBuffer
End If
IsDirty = True
End Sub

Public Function PrepareGenericAction(ByVal ActType As ResGenericUndoStrings) As Action
Dim pAction As Action
pAction.Type = actGenericAction
MakeStructureSnapshot pAction
pAction.UndoString = GetString(ActType)
pAction.RedoString = Replace(pAction.UndoString, GetString(ResUndo), GetString(ResRedo))
PrepareGenericAction = pAction
End Function

Public Sub RecordGenericAction(ByVal ActType As ResGenericUndoStrings)
RecordAction PrepareGenericAction(ActType)
End Sub

Public Sub RecordUndoneAction(pAction As Action)
'If pAction.Type = actGenericAction Or pAction.Type = actGenericReaction Then
'    'do nothing
'Else
'    pAction.UndoString = GetString(ResUndoActionBase + 2 * pAction.Type)
'    pAction.RedoString = Replace(pAction.UndoString, GetString(ResUndo), GetString(ResRedo))
'End If

UndoneActivityCount = UndoneActivityCount + 1
ReDim Preserve UndoneActivity(1 To UndoneActivityCount)
UndoneActivity(UndoneActivityCount) = pAction
i_UpdateRedoMenuStatus True
'FormMain.mnuRedo.Enabled = True
'FormMain.mnuRedo.Caption = GetString(ResRedo) '& IIf(setLanguage = langGerman, ":", "") & " " & GetString(ResUndoActionBase + 2 * pAction.Type)
End Sub

Public Sub ClearUndoBuffer()
ReDim Activity(1 To 1)
ActivityCount = 0
i_UpdateUndoMenuStatus False
End Sub

Public Sub ClearRedoBuffer()
ReDim UndoneActivity(1 To 1)
UndoneActivityCount = 0
i_UpdateRedoMenuStatus False
End Sub

Public Sub RestoreFromStructureSnapshot(pAction As Action)
Dim Z As Long
With pAction
    FigureCount = .pFigure
    RedimFigures 0, FigureCount - 1
    PointCount = .pPoint
    LocusCount = .pLocus
    LabelCount = .pLabel
    ButtonCount = .pButton
    Figures = .sFigure
    StaticGraphics = .sSG
    Buttons = .sButton
    BasePoint = .sPoint
    Locuses = .sLocus
    TextLabels = .sLabel
    StaticGraphicCount = .pSG
    'FormMain.ValueTable1.Clear
    'For Z = 1 To .pWE
    '    FormMain.ValueTable1.AddExpression .sWE(Z).Name, .sWE(Z).Expression, False
    'Next
    'WorldTransform = .CurrentTransform
End With
End Sub

Public Sub RemoveLastUndoActionFromBuffer()
If ActivityCount = 0 Then Exit Sub
ActivityCount = ActivityCount - 1
If ActivityCount > 0 Then ReDim Preserve Activity(1 To ActivityCount) Else ReDim Activity(1 To 1)
End Sub

Public Function UndoMemoryConsumption() As Long
On Local Error Resume Next
Dim MemSum As Long, Q As Long, Z As Long
If ActivityCount = 0 Then UndoMemoryConsumption = 0: Exit Function
MemSum = 0
For Q = 1 To ActivityCount
    With Activity(Q)
        MemSum = MemSum + Len(Activity(Q))
        'If True Then
            If UBound(.sFigure) >= LBound(.sFigure) Then
                For Z = LBound(.sFigure) To UBound(.sFigure)
                    MemSum = MemSum + FigureMemoryConsumption(.sFigure(Z))
                Next
            End If
        'End If
        'If Not .sPoint Is Nothing Then
            If UBound(.sPoint) >= LBound(.sPoint) Then
                For Z = LBound(.sPoint) To UBound(.sPoint)
                    MemSum = MemSum + Len(.sPoint(Z))
                Next
            End If
        'End If
        'If Not .sLabel Is Nothing Then
            If UBound(.sLabel) >= LBound(.sLabel) Then
                For Z = LBound(.sLabel) To UBound(.sLabel)
                    MemSum = MemSum + LabelMemoryConsumption(.sLabel(Z))
                Next
            End If
        'End If
        'If Not .sLocus Is Nothing Then
            If UBound(.sLocus) >= LBound(.sLocus) Then
                For Z = LBound(.sLocus) To UBound(.sLocus)
                    MemSum = MemSum + LocusMemoryConsumption(.sLocus(Z))
                Next
            End If
        'End If
    End With
Next
UndoMemoryConsumption = MemSum
End Function
