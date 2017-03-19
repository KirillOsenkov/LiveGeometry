Attribute VB_Name = "modInterfaceOut"
'This module contains event handlers that are called from pure
'abstract kernel to notify local concrete environment of kernel state changes,
'e.g. a mouse moved over abstract kernel virtual paper,
'       virtual paper scrolled,
'       tool must be selected, status message must be shown,
'       mouse pointer must be changed,
'       menu has to be popped up,
'       etc. etc. etc.
Option Explicit

'###############################################################
'Occurs when virtual paper is scrolled
'###############################################################

Public Sub i_Scrolled(Optional ByVal X As Boolean = True, Optional ByVal Y As Boolean = True)
FormMain.UpdateRulers -X - 2 * Y
End Sub

'###############################################################
'Occurs when mouse moves over virtual paper
'###############################################################

Public Sub i_MouseMoved(Optional ByVal X As Double, Optional ByVal Y As Double)
FormMain.UpdateCoords X, Y
ToPhysical X, Y
FormMain.MoveIndicX X
FormMain.MoveIndicY Y
End Sub

'###############################################################
'Occurs when user right-clicks a point; implies that a menu wants to pop up
'###############################################################

Public Sub i_PointRightClicked(ByVal AP As Long, APoints() As Long, AFigures() As Long)
Dim MaxZOrder As Long, Z As Long

With FormMain
    ActivePoint = AP
    
    'Is there a choose point ambiguity???
    If UBound(APoints) > 1 Then 'if Yes, then present user a beautiful point list
        MaxZOrder = 1
        .mnuChoosePoint.Visible = True
        For Z = 1 To UBound(APoints)
            If .mnuPointChoice.UBound < Z Then Load .mnuPointChoice(Z)
            .mnuPointChoice(Z).Caption = BasePoint(APoints(Z)).Name
            .mnuPointChoice(Z).Visible = True
            .mnuPointChoice(Z).Tag = APoints(Z)
            If BasePoint(APoints(Z)).ZOrder > BasePoint(APoints(MaxZOrder)).ZOrder Then MaxZOrder = Z
        Next Z
        .mnuPointChoice(MaxZOrder).Checked = True
    Else 'otherwise just hide it
        .mnuChoosePoint.Visible = False
    End If
    
    'Are there any figures under cursor???
    If UBound(AFigures) > 0 And IsFigure(AFigures(1)) And BasePoint(AP).Type = dsPoint Then
        .mnuSnapToFigure.Visible = True
        For Z = 1 To UBound(AFigures)
            If .mnuSnapTo.UBound < Z Then Load .mnuSnapTo(Z)
            .mnuSnapTo(Z).Caption = Figures(AFigures(Z)).Name
            .mnuSnapTo(Z).Visible = True
            .mnuSnapTo(Z).Tag = AFigures(Z)
        Next Z
    Else
        .mnuSnapToFigure.Visible = False
    End If
    
    'If AP is a figure point, present user an option to release it
    If BasePoint(AP).Type = dsPointOnFigure Then
        .mnuReleasePoint.Visible = True
    Else
        .mnuReleasePoint.Visible = False
    End If
    
    'Show locus-related details
    If BasePoint(AP).Locus = 0 Then
        .mnuCreateLocus.Visible = True
        .mnuCreateLocus.Checked = False
        .mnuClearLocus.Visible = False
    Else
        If Locuses(BasePoint(AP).Locus).Dynamic Then
            .mnuClearLocus.Visible = True
            .mnuCreateLocus.Visible = False
        Else
            .mnuCreateLocus.Visible = True
            .mnuCreateLocus.Checked = Locuses(BasePoint(AP).Locus).Enabled
            .mnuClearLocus.Visible = True
        End If
    End If
    
    .mnuShowPointName.Checked = BasePoint(AP).ShowName
    
    ' pop it up finally...
    .PopupMenu .mnuPointPopup  '<--- here it pops up!!!!!
    
    '...and cleanup the consequences...
    If .mnuCreateLocus.Visible = False Then .mnuCreateLocus.Visible = True
    
    Do While .mnuSnapTo.UBound > 1
        Unload .mnuSnapTo(.mnuSnapTo.UBound)
    Loop
    Do While .mnuPointChoice.UBound > 1
        Unload .mnuPointChoice(.mnuPointChoice.UBound)
    Loop
    .mnuPointChoice(1).Checked = False
End With
End Sub

'###############################################################
'Occurs when user right-clicks a figure; implies that a menu wants to pop up
'###############################################################

Public Sub i_FigureRightClicked(ByVal AF As Long, AFigures() As Long)
Dim Z As Long, MaxZOrder As Long, SGC As Long, P1 As Long, P2 As Long, P3 As Long, P4 As Long

With FormMain
    ActiveFigure = AF
    If UBound(AFigures) > 1 Then 'more than one figure happened to be under cursor
        MaxZOrder = 1
        .mnuChooseFigure.Visible = True
        For Z = 1 To UBound(AFigures)
            If .mnuFigureChoice.UBound < Z Then Load .mnuFigureChoice(Z)
            .mnuFigureChoice(Z).Caption = Figures(AFigures(Z)).Name
            .mnuFigureChoice(Z).Visible = True
            .mnuFigureChoice(Z).Tag = AFigures(Z)
            If Figures(AFigures(Z)).ZOrder > Figures(AFigures(MaxZOrder)).ZOrder Then MaxZOrder = Z
        Next Z
        .mnuFigureChoice(MaxZOrder).Checked = True
    Else
        .mnuChooseFigure.Visible = False 'no need to choose figures among single one...
    End If
    
    If Figures(ActiveFigure).FigureType = dsSegment Then 'vector operation
        SGC = 0
        For Z = 1 To StaticGraphicCount
            If StaticGraphics(Z).Type = sgVector Then
                P1 = StaticGraphics(Z).Points(1)
                P2 = StaticGraphics(Z).Points(2)
                P3 = Figures(ActiveFigure).Points(0)
                P4 = Figures(ActiveFigure).Points(1)
                If (P1 = P3 And P2 = P4) Or (P1 = P4 And P2 = P3) Then
                    If Not .mnuVectorProperties.Visible Then .mnuVectorProperties.Visible = True
                    If Not .mnuVectorDelete.Visible Then .mnuVectorDelete.Visible = True
                    SGC = SGC + 1
                    If .mnuVectorProp.UBound < SGC Then Load .mnuVectorProp(SGC)
                    .mnuVectorProp(SGC).Caption = BasePoint(P1).Name & BasePoint(P2).Name
                    .mnuVectorProp(SGC).Tag = Z
                    .mnuVectorProp(SGC).Visible = True
                    If .mnuVectorDel.UBound < SGC Then Load .mnuVectorDel(SGC)
                    .mnuVectorDel(SGC).Caption = BasePoint(P1).Name & BasePoint(P2).Name
                    .mnuVectorDel(SGC).Tag = Z
                    .mnuVectorDel(SGC).Visible = True
                End If
            End If
        Next
    End If
    
    If Figures(ActiveFigure).FigureType = dsMeasureDistance Or Figures(ActiveFigure).FigureType = dsMeasureAngle Then
        .mnuFigureProperties.Visible = False
        .mnuHideFigure.Visible = False
        .mnuFigureSep1.Visible = False
        .mnuDeleteFigure.Caption = GetString(ResMnuDeleteMeasurement)
        .mnuMeasurementProperties.Visible = True
    End If
    
    'popup the notorious menu...
    .PopupMenu .mnuFigurePopup
    
    '..and clean up afterwards...
    Do While .mnuFigureChoice.UBound > 1
        Unload .mnuFigureChoice(.mnuFigureChoice.UBound)
    Loop
    Do While .mnuVectorProp.UBound > 1
        Unload .mnuVectorProp(.mnuVectorProp.UBound)
    Loop
    Do While .mnuVectorDel.UBound > 1
        Unload .mnuVectorDel(.mnuVectorDel.UBound)
    Loop
    .mnuVectorProperties.Visible = False
    .mnuVectorDelete.Visible = False
    .mnuFigureChoice(1).Checked = False
    
    .mnuMeasurementProperties.Visible = False
    .mnuFigureProperties.Visible = True
    .mnuHideFigure.Visible = True
    .mnuFigureSep1.Visible = True
    .mnuDeleteFigure.Caption = GetString(ResMnuDeleteFigure)
End With
End Sub

'###############################################################
'Occurs when user right-clicks a label; implies that a menu wants to pop up
'###############################################################

Public Sub i_LabelRightClicked(ByVal AL As Long)
With FormMain
    ActiveLabel = AL
    If TextLabels(AL).Dynamic Then .mnuRecalcLabel.Visible = True Else .mnuRecalcLabel.Visible = False
    .mnuFixLabel.Checked = TextLabels(AL).Fixed
    .PopupMenu .mnuLabelPopup
End With
End Sub

'###############################################################
'Occurs when user right-clicks a point; implies that a menu wants to pop up
'###############################################################

Public Sub i_SGRightClicked(ByVal ASG As Long)
With FormMain
    ActiveStatic = ASG
    '.mnuSGProperties.Caption = GetString(ResStaticObjectBase + StaticGraphics(ActiveStatic).Type * 2) & " " & GetString(ResPropsTitle)
    .mnuSGProperties.Caption = GetString(ResPropertiesOfAPolygon + StaticGraphics(ActiveStatic).Type * 2)
    '.mnuSGDelete.Caption = GetString(ResMnuDeleteObject) & " (" & GetString(ResStaticObjectBase + StaticGraphics(ActiveStatic).Type * 2) & ")"
    .mnuSGDelete.Caption = GetString(ResDeletePolygon + StaticGraphics(ActiveStatic).Type * 2)
    .PopupMenu .mnuSGPopup
End With
End Sub

'###############################################################
'Occurs when user right-clicks a button; implies that a menu wants to pop up
'###############################################################

Public Sub i_ButtonRightClicked(ByVal AB As Long)
With FormMain
    ActiveButton = AB
    .mnuButtonMovable.Checked = Buttons(ActiveButton).Fixed
    .PopupMenu .mnuButtonPopup
End With
End Sub

'###############################################################
'need to show a message in the statusbar; links to concrete statusbar object
'###############################################################

Public Sub i_ShowStatus(Optional ByVal S As String = "")
FormMain.ShowStatus S
End Sub

'###############################################################
'need to set the mouse pointer in a proper way
'###############################################################

Public Sub i_SetMousePointer(ByVal MousePointer As CursorState)
FormMain.SetMousePointer MousePointer
End Sub

'###############################################################
'need to select a tool (make it active)
'###############################################################

Public Sub i_SelectTool(ByVal ToolNumber As DrawState)
FormMain.SelectTool ToolNumber
'ImitateMouseMove
End Sub

'=======================================================
'                  Request to enter "macro given select" interaction mode
'=======================================================

Public Sub i_EnterMacroGivenSelectMode()
FormMain.EnterMacroGivenSelectMode
End Sub

'=======================================================
'                  Request to enter "macro result select" interaction mode
'=======================================================

Public Sub i_EnterMacroResultSelectMode()
FormMain.EnterMacroResultSelectMode
End Sub

'=======================================================
'                  Request to exit "macro result select" interaction mode
'=======================================================

Public Sub i_ExitMacroResultSelectMode()
FormMain.ExitMacroCreateMode
End Sub

'=======================================================
'                  Request to enter "macro run" interaction mode
'=======================================================

Public Sub i_EnterMacroRunMode()
FormMain.EnterMacroRunMode
End Sub

Public Sub i_ExitMacroRunMode()
FormMain.ExitMacroRunMode
End Sub

'###############################################################
'Mouse click during macro givens selection...
'###############################################################
'Public Sub i_MacroGivensClick(ByVal AP As Long, ByVal AF As Long, APoints() As Long, AFigures() As Long)
Public Sub i_MacroGivensClick(objList As ObjectList)
Dim bFlagCount As Long, Z As Long

With FormMain
    bFlagCount = 0
    
    If objList.PointCount > 0 Then
        For Z = 1 To objList.PointCount
            bFlagCount = bFlagCount + 1
            If .mnuMacroObject.UBound < bFlagCount Then Load .mnuMacroObject(bFlagCount)
            .mnuMacroObject(bFlagCount).Caption = BasePoint(objList.Points(Z)).Name
            .mnuMacroObject(bFlagCount).Visible = True
            .mnuMacroObject(bFlagCount).Tag = -objList.Points(Z)
            .mnuMacroObject(bFlagCount).Checked = GivenPoints(objList.Points(Z)) > 0
        Next Z
    End If
    
    If objList.FigureCount > 0 And Not (objList.PointCount = 1) Then
        For Z = 1 To objList.FigureCount
            If Figures(objList.Figures(Z)).FigureType <> dsMeasureDistance And Figures(objList.Figures(Z)).FigureType <> dsMeasureAngle Then
                bFlagCount = bFlagCount + 1
                If .mnuMacroObject.UBound < bFlagCount Then Load .mnuMacroObject(bFlagCount)
                .mnuMacroObject(bFlagCount).Caption = Figures(objList.Figures(Z)).Name
                .mnuMacroObject(bFlagCount).Visible = True
                .mnuMacroObject(bFlagCount).Tag = objList.Figures(Z)
                .mnuMacroObject(bFlagCount).Checked = GivenFigures(objList.Figures(Z)) > 0
            End If
        Next Z
    End If
    
'    If UBound(AFigures) = LBound(AFigures) And IsPoint(AP) And IsFigure(AF) Then
'        bFlagCount = bFlagCount + 1
'        If .mnuMacroObject.UBound < bFlagCount Then Load .mnuMacroObject(bFlagCount)
'        .mnuMacroObject(bFlagCount).Caption = Figures(AF).Name
'        .mnuMacroObject(bFlagCount).Visible = True
'        .mnuMacroObject(bFlagCount).Tag = AF
'        .mnuMacroObject(bFlagCount).Checked = GivenFigures(AF) > 0
'    End If
'    If UBound(APoints) = LBound(APoints) And IsPoint(AP) And IsFigure(AF) Then
'        bFlagCount = bFlagCount + 1
'        If .mnuMacroObject.UBound < bFlagCount Then Load .mnuMacroObject(bFlagCount)
'        .mnuMacroObject(bFlagCount).Caption = BasePoint(AP).Name
'        .mnuMacroObject(bFlagCount).Visible = True
'        .mnuMacroObject(bFlagCount).Tag = -AP
'        .mnuMacroObject(bFlagCount).Checked = GivenPoints(AP) > 0
'    End If
    
    If bFlagCount > 0 Then ' should popup a menu
        If bFlagCount = 1 Then
            MacroObjectSelected .mnuMacroObject(1).Tag
'            With .mnuMacroObject(1)
'                If .Tag < 0 Then
'                    If GivenPoints(-.Tag) = 0 Then
'                        GivenPoints(-.Tag) = AddMacroGiven(tempMacro, dsPoint)
'                        ShowSelectedPoint Paper.hDC, -.Tag, True
'                    End If
'                Else
'                    If GivenFigures(.Tag) = 0 Then
'                        GivenFigures(.Tag) = AddMacroGiven(tempMacro, Figures(.Tag).FigureType)
'                        ShowSelectedFigure Paper.hDC, .Tag, True
'                    End If
'                End If
'            End With
        Else
            .PopupMenu .mnuChooseMacroObject
        End If
        
        Do While .mnuMacroObject.UBound > 1
            Unload .mnuMacroObject(.mnuMacroObject.UBound)
        Loop
        
        Exit Sub
    End If
    
'    If IsPoint(AP) Then
'        MacroObjectSelected -AP
''        If GivenPoints(AP) = 0 Then
''            GivenPoints(AP) = AddMacroGiven(tempMacro, dsPoint)
''            ShowSelectedPoint Paper.hDC, AP, True
''        End If
''        Exit Sub
'    End If
'    If IsFigure(AF) Then
'        MacroObjectSelected AF
''        If GivenFigures(AF) = 0 Then
''            GivenFigures(AF) = AddMacroGiven(tempMacro, Figures(AF).FigureType)
''            ShowSelectedFigure Paper.hDC, AF, True
''        End If
'    End If
End With
End Sub

'###############################################################
'Mouse click during macro Results selection...
'###############################################################
'Public Sub i_MacroResultsClick(ByVal AP As Long, ByVal AF As Long, ByVal ASG As Long, APoints() As Long, AFigures() As Long, ASGs() As Long, Optional ByVal AL As Long = 0)
Public Sub i_MacroResultsClick(objList As ObjectList)
'==================================================
' objList is a list of all objects under mouse cursor during the click
' usually one object
' prepares and calls popup menu to specify one object from objList
'==================================================
Dim bFlagCount As Long, Z As Long, tFig As Long

With FormMain
    bFlagCount = 0
    
    If objList.PointCount > 0 Then
        For Z = 1 To objList.PointCount
            If ResultPoints(objList.Points(Z)) >= 0 Then
                tFig = BasePoint(objList.Points(Z)).ParentFigure
                bFlagCount = bFlagCount + 1
                If .mnuMacroObject.UBound < bFlagCount Then Load .mnuMacroObject(bFlagCount)
                .mnuMacroObject(bFlagCount).Caption = Figures(tFig).Name
                .mnuMacroObject(bFlagCount).Visible = True
                .mnuMacroObject(bFlagCount).Tag = tFig
                .mnuMacroObject(bFlagCount).Checked = IsResultSelected(tFig)
            End If
        Next Z
    End If
    
    If objList.FigureCount > 0 And Not (objList.PointCount = 1 And objList.SGCount = 0) Then
        For Z = 1 To objList.FigureCount
            tFig = objList.Figures(Z)
            
            If ResultFigures(tFig) >= 0 Then
                bFlagCount = bFlagCount + 1
                If .mnuMacroObject.UBound < bFlagCount Then Load .mnuMacroObject(bFlagCount)
                .mnuMacroObject(bFlagCount).Caption = Figures(tFig).Name
                .mnuMacroObject(bFlagCount).Visible = True
                .mnuMacroObject(bFlagCount).Tag = tFig
                .mnuMacroObject(bFlagCount).Checked = IsResultSelected(tFig)
            End If
            
        Next Z
    End If
    
    If objList.SGCount > 0 Then
        For Z = 1 To objList.SGCount
            tFig = objList.SGs(Z)
            
            If ResultSGs(tFig) >= 0 Then
                bFlagCount = bFlagCount + 1
                If .mnuMacroObject.UBound < bFlagCount Then Load .mnuMacroObject(bFlagCount)
                .mnuMacroObject(bFlagCount).Caption = GetObjectName(gotSG, tFig)
                .mnuMacroObject(bFlagCount).Visible = True
                .mnuMacroObject(bFlagCount).Tag = -tFig
                .mnuMacroObject(bFlagCount).Checked = IsSGAdded(tFig)
            End If
            
        Next
    End If
    
    '======================================================================
    
    If bFlagCount > 0 Then
        If bFlagCount = 1 Then
            With .mnuMacroObject(1)
                MacroObjectSelected .Tag
'                Figures(.Tag).Tag = AddMacroResult(tempMacro, Figures(.Tag))
'                If Figures(.Tag).Tag <= 0 Then i_CancelMacro GetString(ResMacroErrBase - 2 * Figures(.Tag).Tag): Exit Sub
'                ShowSelectedFigure Paper.hDC, .Tag, True
            End With
        Else
            .PopupMenu .mnuChooseMacroObject
        End If
        
        Do While .mnuMacroObject.UBound > 1
            Unload .mnuMacroObject(.mnuMacroObject.UBound)
        Loop
        
        Exit Sub
    
    Else 'nothing added to the menu; only loci left...
        tFig = ObjectListGetUpperLocus(objList)
        If tFig > 0 Then
            If ResultLoci(tFig) > 0 Then
                ResultLoci(tFig) = 0
                PaperCls
                ShowAllWithResults
            ElseIf ResultLoci(tFig) = 0 Then
                ResultLoci(tFig) = 1
                PaperCls
                ShowAllWithResults
            End If
        End If
    End If
    
    '======================================================================
    
'    If IsPoint(AP) Then
'        If BasePoint(AP).Type <> dsPoint And Figures(BasePoint(AP).ParentFigure).Tag = 0 And GivenPoints(AP) = 0 Then
'            Figures(BasePoint(AP).ParentFigure).Tag = AddMacroResult(tempMacro, Figures(BasePoint(AP).ParentFigure))
'            If Figures(BasePoint(AP).ParentFigure).Tag <= 0 Then i_CancelMacro GetString(ResMacroErrBase - 2 * Figures(BasePoint(AP).ParentFigure).Tag): Exit Sub
'            ShowSelectedPoint Paper.hDC, AP, True
'            Exit Sub
'        End If
'    End If
'    If IsFigure(AF) Then
'        If Figures(AF).Tag = 0 And GivenFigures(AF) = 0 Then
'            Figures(AF).Tag = AddMacroResult(tempMacro, Figures(AF))
'            If Figures(AF).Tag <= 0 Then i_CancelMacro GetString(ResMacroErrBase - 2 * Figures(AF).Tag): Exit Sub
'            ShowSelectedFigure Paper.hDC, AF, True
'        End If
'        Exit Sub
'    End If
'    If IsSG(ASG) Then
'        If StaticGraphics(ASG).Tag = 0 Then
'            StaticGraphics(ASG).Tag = AddMacroSG(tempMacro, ASG)
'            If StaticGraphics(ASG).Tag <= 0 Then
'                i_CancelMacro GetString(ResMacroErrBase - 2 * Figures(AF).Tag): Exit Sub
'            End If
'            DrawStaticGraphicSelected Paper.hDC, ASG, True, True
'        End If
'    End If
'    If IsLocus(AL) Then
'        AP = Locuses(AL).ParentPoint
'        If IsPoint(AP) Then
'            If BasePoint(AP).Type <> dsPoint And Figures(BasePoint(AP).ParentFigure).Tag = 0 And GivenPoints(AP) = 0 Then
'                Figures(BasePoint(AP).ParentFigure).Tag = AddMacroResult(tempMacro, Figures(BasePoint(AP).ParentFigure))
'                If Figures(BasePoint(AP).ParentFigure).Tag <= 0 Then i_CancelMacro GetString(ResMacroErrBase - 2 * Figures(BasePoint(AP).ParentFigure).Tag): Exit Sub
'                ShowLocusSelected Paper.hDC, AL, , True
'                Exit Sub
'            End If
'        End If
'    End If
End With
End Sub

'###############################################################
'Exit macro creation mode
'###############################################################

Public Sub i_CancelMacro(Optional ByVal ErrStr As String)
InitZeroTags
FormMain.ExitMacroCreateMode
MacroCancel ErrStr
End Sub

'###############################################################
'Exit Macro launch mode
'###############################################################

'Public Sub i_CancelMacroRun()
'Dim Z As Long
'
'DragS.MacroCurrentObject = 1
'DragS.MacroObjectCount = 0
'ReDim DragS.MacroObjects(1 To 1)
'ReDim DragS.MacroObjectType(1 To 1)
'DragS.State = dscNormalState
'For Z = 1 To MaxDragNumbers: DragS.Number(Z) = 0:  Next
'FormMain.EnableMenu
'PaperCls
'ShowAll
'ImitateMouseMove
'End Sub

'###############################################################
'Link to CancelOperation
'###############################################################

Public Sub i_CancelOperation()
If Not FormMain.Fullscreen Then FormMain.EnableMenus mnsStandard
End Sub

'###############################################################
'Disable menu before starting operation
'###############################################################

Public Sub i_StartOperation()
If Not FormMain.Fullscreen Then FormMain.EnableMenus mnsCompleteDisable
End Sub

'=========================================================
'An action was performed; update the state of Undo menu
'=========================================================

Public Sub i_UndoAdded(ByVal ActionTypeString As String)
If DragS.State = dscNormalState Then FormMain.mnuUndo.Enabled = True
FormMain.mnuUndo.Caption = ActionTypeString 'GetString(ResUndo) & IIf(setLanguage = langGerman, ":", "") & " " & GetString(ResUndoActionBase + 2 * T)
End Sub

'=========================================================
' Check/uncheck Dynamic locus / SG button
'=========================================================
Public Sub i_CheckMainbarButton(ByVal Caption As String, Optional ByVal Check As Boolean = True)
Dim Z As Long

Z = FormMain.MenuBar(1).Item(Caption)
If Z <> 0 Then FormMain.MenuBar(1).CheckItem Z, Check
FormMain.MenuBar(1).Refresh
End Sub

'=========================================================
' Show figure list dialog
'=========================================================

Public Sub i_ShowFigureList()
If VisualFigureCount = 0 Then Exit Sub

frmFigureList.Show vbModal
End Sub

'=========================================================
' Enable/disable mnuUndo
'=========================================================

Public Sub i_UpdateUndoMenuStatus(Optional ByVal ShouldEnable)
Dim En As Boolean

If IsMissing(ShouldEnable) Then En = ActivityCount > 0 Else En = CBool(ShouldEnable)

FormMain.mnuUndo.Enabled = En
If Not En Then FormMain.mnuUndo.Caption = GetString(ResUndo)
End Sub

'=========================================================
' Enable/disable mnuRedo
'=========================================================

Public Sub i_UpdateRedoMenuStatus(Optional ByVal ShouldEnable)
Dim En As Boolean

If IsMissing(ShouldEnable) Then En = UndoneActivityCount > 0 Else En = CBool(ShouldEnable)

FormMain.mnuRedo.Enabled = En
If Not En Then FormMain.mnuRedo.Caption = GetString(ResRedo)
End Sub

'=========================================================
' Some dialog closed; so get focus back to the main container
'=========================================================
Public Sub i_NeedToSetFocus()
If FormMain.Enabled And FormMain.Visible Then FormMain.SetFocus
End Sub

'=========================================================
' Add a new macro menu item
'=========================================================
Public Sub i_AddMacroRunMenu(ByVal Caption As String)
FormMain.mnuMacroSep2.Visible = True
If FormMain.mnuMacroRun.UBound < MacroCount - 1 Then Load FormMain.mnuMacroRun(MacroCount - 1)
FormMain.mnuMacroRun(MacroCount - 1).Caption = Caption
FormMain.mnuMacroRun(MacroCount - 1).Visible = True
End Sub

'=========================================================
' Remove macro menu item
' MacroCount is still the same (not decremented)
' Index: ID º [1; MacroCount]
'=========================================================
Public Sub i_RemoveMacroRunMenu(ByVal Index As Long)
Dim Z As Long

If Index < 1 Or Index > MacroCount Then Exit Sub

If Index < MacroCount Then
    For Z = Index To MacroCount - 1
        FormMain.mnuMacroRun(Z - 1).Caption = FormMain.mnuMacroRun(Z).Caption
    Next
End If

If MacroCount > 1 Then Unload FormMain.mnuMacroRun(MacroCount - 1) Else FormMain.mnuMacroRun(0).Visible = False

FormMain.mnuMacroSep2.Visible = MacroCount > 1
End Sub

'=========================================================
' Notify the user that the point already exists.
' Let him make decision by showing the appropriate dialog.
'=========================================================
Public Sub i_AskAboutPointRename(ByVal Index As Long, ByVal NewName As String)
frmPointRename.OldName = BasePoint(Index).Name
frmPointRename.NewName = NewName
frmPointRename.Index = Index
frmPointRename.FillDialogStrings
frmPointRename.Show vbModal
End Sub
