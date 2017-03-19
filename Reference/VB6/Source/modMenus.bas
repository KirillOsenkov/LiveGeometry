Attribute VB_Name = "modMenus"
Option Explicit

Public Sub CreateMenus()
Dim BAG As Boolean, MenuPic As StdPicture, MenuPic2 As IPictureDisp, Z As Long
DrawingState = dsSelect
PrepareHelpFile

With FormMain

    .mnuPrint.Visible = False 'Printers.Count > 0
    .mnuFileSep1.Visible = .mnuPrint.Visible

    .mnuFile.Caption = "&" & GetString(ResFile)
    .mnuNew.Caption = "&" & GetString(ResNew)
    .mnuOpen.Caption = "&" & GetString(ResOpen)
    .mnuSave.Caption = "&" & GetString(ResSave)
    .mnuSaveAs.Caption = GetString(ResSaveAs)
    
    '.mnuSave.Visible = CanSave
    '.mnuSaveAs.Visible = CanSave
    
    .mnuPrint.Caption = GetString(ResPrint)
    .mnuExport.Caption = GetString(ResExport)
    .mnuBitmap.Caption = GetString(ResBMP)
    .mnuMetafile.Caption = GetString(ResWMF)
    .mnuEnhancedMetafile.Caption = GetString(ResEMF)
    .mnuExit.Caption = "&" & GetString(ResExit)
    
    .mnuEdit.Caption = "&" & GetString(ResEdit)
    .mnuUndo.Caption = GetString(ResUndo)
    .mnuUndo.Enabled = ActivityCount > 0
    .mnuRedo.Caption = GetString(ResRedo)
    .mnuRedo.Enabled = UndoneActivityCount > 0
    .mnuInsertLabel.Caption = GetString(ResInsertLabel)
    .mnuInsertButton.Caption = GetString(ResInsertButton)
    .mnuCalculator.Caption = GetString(ResCalculator)
    .mnuClearAll.Caption = GetString(ResClearAll)
    .mnuFileProps.Caption = GetString(ResFileProps)
    
    .mnuView.Caption = "&" & GetString(ResView)
    .mnuFigureList.Caption = GetString(ResFigureList)
    .mnuPointList.Caption = GetString(ResPointList)
    .mnuFullscreen.Caption = GetString(ResFullscreen)
    .mnuShowStatusbar.Caption = GetString(ResShowStatusbar)
    .mnuShowStatusbar.Checked = setShowStatusbar
    .mnuShowToolbar.Checked = setShowToolbar
    .mnuShowMainbar.Caption = GetString(ResShowMainbar)
    .mnuShowMainbar.Checked = setShowMainbar
    
    .mnuShowToolbar.Caption = GetString(ResShowToolbar)
    .mnuDemo.Caption = GetString(ResDemo)
    .mnuDemoOptions.Caption = GetString(ResDemoOptions)
    .mnuPolygon.Caption = GetString(ResStaticObjectBase + 2 * sgPolygon)
    .mnuBezier.Caption = GetString(ResStaticObjectBase + 2 * sgBezier)
    .mnuVector.Caption = GetString(ResStaticObjectBase + 2 * sgVector)
    .mnuAnalytic.Caption = GetString(ResMnuAnalytic)
    .mnuAnCircle.Caption = GetString(ResMnuAnCircle)
    .mnuAnLine.Caption = GetString(ResMnuAnLine)
    .mnuActiveAxes.Caption = GetString(ResActiveAxes)
    .mnuAnPoint.Caption = GetString(ResMnuAnPoint)
    
    .mnuFigures.Caption = GetString(ResMnuFigures)
    .mnuWE.Caption = GetString(ResWEWindow)
    .mnuFigCircles.Caption = GetString(ResToolCircles)
    .mnuFigLines.Caption = GetString(ResToolLines)
    .mnuFigPoints.Caption = GetString(ResToolPoints)
    .mnuFigConstruction.Caption = GetString(ResToolConstruction)
    .mnuFigMeasure.Caption = GetString(ResToolMeasure)
    
    .mnuToolArc.Caption = GetString(ResFigureBase + 2 * dsCircle_ArcCenterAndRadiusAndTwoPoints)
    .mnuToolCircle.Caption = GetString(ResFigureBase + 2 * dsCircle_CenterAndCircumPoint)
    .mnuToolCircleByRadius.Caption = GetString(ResFigureBase + 2 * dsCircle_CenterAndTwoPoints)
    .mnuToolIntersect.Caption = GetString(ResFigureBase + 2 * dsIntersect)
    .mnuToolInvert.Caption = GetString(ResFigureBase + 2 * dsInvert)
    .mnuToolLine.Caption = GetString(ResFigureBase + 2 * dsLine_2Points)
    .mnuToolMeasureAngle.Caption = GetString(ResFigureBase + 2 * dsMeasureAngle)
    .mnuToolMeasureDistance.Caption = GetString(ResFigureBase + 2 * dsMeasureDistance)
    .mnuToolMeasureArea.Caption = GetString(ResFigureBase + 2 * dsMeasureArea)
    .mnuToolMiddlePoint.Caption = GetString(ResFigureBase + 2 * dsMiddlePoint)
    .mnuToolParallelLine.Caption = GetString(ResFigureBase + 2 * dsLine_PointAndParallelLine)
    .mnuToolBisector.Caption = GetString(ResFigureBase + 2 * dsBisector)
    .mnuToolPoint.Caption = GetString(ResFigureBase + 2 * dsPoint)
    .mnuToolPointOnFigure.Caption = GetString(ResFigureBase + 2 * dsPointOnFigure)
    .mnuToolPerpendicularLine.Caption = GetString(ResFigureBase + 2 * dsLine_PointAndPerpendicularLine)
    .mnuToolRay.Caption = GetString(ResFigureBase + 2 * dsRay)
    .mnuToolReflectedPoint.Caption = GetString(ResFigureBase + 2 * dsSimmPointByLine)
    .mnuToolSegment.Caption = GetString(ResFigureBase + 2 * dsSegment)
    .mnuToolSymmPoint.Caption = GetString(ResFigureBase + 2 * dsSimmPoint)
    .mnuDynamicLocus.Caption = GetString(ResFigureBase + 2 * dsDynamicLocus)
    
    .mnuMacros.Caption = GetString(ResMnuMacros)
    .mnuMacroCreate.Caption = GetString(ResMnuMacroCreate)
    .mnuMacroLoad.Caption = GetString(ResMnuMacroLoad)
    .mnuMacroResults.Caption = GetString(ResMnuMacroSelectResults)
    .mnuMacroOrganize.Caption = GetString(ResMnuMacroOrganize)
    .mnuMacroSave.Caption = GetString(ResMnuMacroSave)

'    If .Docked.Visible Then
'        .mnuWE.Caption = GetString(ResHide) & " " & GetString(ResWEWindow)
'    Else
'        .mnuWE.Caption = GetString(ResShow) & " " & GetString(ResWEWindow)
'    End If
    .mnuOptions.Caption = "&" & GetString(ResOptions)
    .mnuLangEnglish.Checked = setLanguage = langEnglish
    .mnuLangRussian.Checked = setLanguage = langRussian
    .mnuLangGerman.Checked = setLanguage = langGerman
    .mnuLangUkrainian.Checked = setLanguage = langUkrainian
    .mnuSettings.Caption = "&" & GetString(ResOptions)
    .mnuHelp.Caption = "&" & GetString(ResHelp)
    .mnuHelpContents.Caption = "&" & GetString(ResHelpContents)

#If conTips = 1 Then
    .mnuTip.Caption = GetString(ResTipOfTheDay)
    .mnuTip.Visible = True
#End If

    .mnuAbout.Caption = GetString(ResAbout)
    
    .mnuChooseFigure.Caption = GetString(ResMnuChooseFigure)
    .mnuFigureProperties.Caption = GetString(ResMnuFigureProperties)
    .mnuHideFigure.Caption = GetString(ResHide)
    .mnuDeleteFigure.Caption = GetString(ResMnuDeleteFigure)
    
    .mnuChoosePoint.Caption = GetString(ResMnuChoosePoint)
    .mnuPointProperties.Caption = GetString(ResMnuPointProperties)
    .mnuShowPointName.Caption = GetString(ResShowName)
    .mnuReleasePoint.Caption = GetString(ResMnuReleasePoint)
    .mnuSnapToFigure.Caption = GetString(ResMnuSnapToFigure)
    .mnuHidePoint.Caption = GetString(ResHide)
    .mnuDeletePoint.Caption = GetString(ResMnuDeletePoint)
    .mnuVectorProperties.Caption = GetString(ResPropertiesOfAVector)
    .mnuVectorDelete.Caption = GetString(ResDeleteVector)
    .mnuMeasurementProperties.Caption = GetString(ResMnuMeasurementProperties)
    
    .mnuLabelProperties.Caption = GetString(ResMnuLabelProperties)
    .mnuRecalcLabel.Caption = GetString(ResMnuRecalcLabel)
    .mnuFixLabel.Caption = GetString(ResFix)
    .mnuDeleteLabel.Caption = GetString(ResMnuDeleteLabel)
    
    .mnuLocusProps.Caption = GetString(ResLocusProps)
    .mnuCreateLocus.Caption = GetString(ResCreateLocus)
    .mnuClearLocus.Caption = GetString(ResDeleteLocus)
    
    .mnuShowAxes.Caption = GetString(ResShowAxes)
    .mnuShowGrid.Caption = GetString(ResShowGrid)
    .mnuShowRulers.Caption = GetString(ResShowRulers)
    
    .mnuButtonDelete.Caption = GetString(ResDeleteButton)
    .mnuButtonMovable.Caption = GetString(ResFix)
    .mnuButtonProperties.Caption = GetString(ResButtonProperties)

    If InDebugMode Then .mnuDebug.Visible = True
    
    '=======================================================
    
    FillMenuBar , 1
    FillMenuBar
    
    
    '=======================================================
    
    'SetMenuItemBitmaps GetSubMenu(GetMenu(.hWnd), 0), 0, MF_BYPOSITION, GetPicture(ResIconMenuNew, , 0).handle, 0
    '********************************************************************
'    For Z = .MenuBar.LBound To .MenuBar.UBound
'        .MenuBar(Z).Clear
'    Next
'
'    .MenuBar(2).AddItem GetString(ResSelect), 1, GetString(ResSelect), , , GetPicture(dsSelect), setDisplayIconText
'    For Z = dsPoint To dsMeasureAngle
'        If Z = dsSegment Or Z = dsCircle_CenterAndCircumPoint Or Z = dsMiddlePoint Or Z = dsMeasureDistance Then BAG = True Else BAG = False
'        .MenuBar(2).AddItem GetString(ResFigureBase + Z * 2), 1, GetString(ResFigureBase + Z * 2), , , GetPicture(Z), setDisplayIconText, BAG
'    Next Z
'    Menus(2).Items(1).Checked = True
'
'
'    For Z = .MenuBar.LBound To .MenuBar.UBound
'        .MenuBar(Z).Refresh
'    Next
    '*******************************************************************
'1     .MenuBar(1).AddItem GetString(ResFile), 1, GetString(ResMnuBase)
'2     .MenuBar(1).AddItem GetString(ResNew), 2, GetString(ResMnuCreateNew)
'3     .MenuBar(1).AddItem GetString(ResOpen), 2, GetString(ResMnuOpenExisting)
'4     .MenuBar(1).AddItem GetString(ResSave), 2, GetString(ResMnuSave)
'5     .MenuBar(1).AddItem GetString(ResSaveAs), 2, GetString(ResMnuSaveAs)
'6     .MenuBar(1).AddItem GetString(ResExport), 2, GetString(ResExport)
'7     .MenuBar(1).AddItem GetString(ResBMP), 3, GetString(ResBMP)
'8     .MenuBar(1).AddItem GetString(ResExit), 2, GetString(ResMnuExit), , , , , True
'        '.MenuBar(1).AddItem GetString(ResEdit), 1, GetString(ResMnuEdit)
'9     .MenuBar(1).AddItem GetString(ResView), 1, GetString(ResMnuView)
'10     If .Docked.Visible Then
'        .MenuBar(1).AddItem GetString(ResHide) & " " & GetString(ResWEWindow), 2, GetString(ResHide) & " " & GetString(ResWEWindow)
'    Else
'        .MenuBar(1).AddItem GetString(ResShow) & " " & GetString(ResWEWindow), 2, GetString(ResShow) & " " & GetString(ResWEWindow)
'    End If
'11     .MenuBar(1).AddItem GetString(ResOptions), 1, GetString(ResMnuOptions)
'12     .MenuBar(1).AddItem GetString(ResSettings), 2, GetString(ResMnuSettings)
    'FormMain.MenuBar(1).AddItem GetString(ResHelp), 1
    'FormMain.MenuBar(1).AddItem GetString(ResHelpContents), 2
    'FormMain.MenuBar(1).AddItem GetString(ResAbout), 3
    
    '-----------------------------------------------------------------
'    .MenuBar(1).AddItem "W", 1, , , , , , , False
'    .MenuBar(1).AddItem GetString(ResMnuFigureProperties), 2, GetString(ResMnuFigureProperties), , True
'    .MenuBar(1).AddItem GetString(ResMnuDeleteFigure), 2, GetString(ResMnuDeleteFigure), , True, , , True
'    .MenuBar(1).AddItem GetString(ResMnuChooseFigure), 2, GetString(ResMnuChooseFigure), , True, , , , False
'    .MenuBar(1).AddItem "", 1, , , True
'    .MenuBar(1).AddItem GetString(ResMnuPointProperties), 2, GetString(ResMnuPointProperties), , True
'    .MenuBar(1).AddItem GetString(ResMnuSnapToFigure), 2, GetString(ResMnuSnapToFigure), , True
'    .MenuBar(1).AddItem GetString(ResMnuReleasePoint), 2, GetString(ResMnuReleasePoint), , True
'    .MenuBar(1).AddItem GetString(ResMnuDeletePoint), 2, GetString(ResMnuDeletePoint), , True, , , True
'    .MenuBar(1).AddItem GetString(ResMnuChoosePoint), 2, GetString(ResMnuChoosePoint), , True, , , , False
'    .MenuBar(1).AddItem "", 1, , , True
'    .MenuBar(1).AddItem GetString(ResMnuLabelProperties), 2, GetString(ResMnuLabelProperties), , True
'    .MenuBar(1).AddItem GetString(ResMnuDeleteLabel), 2, GetString(ResMnuDeleteLabel), , True, , , True
'    .MenuBar(1).AddItem "", 1, , , True
    
    '___________________________________________
'    If setGroupTools Then
'        Dim TempPic As StdPicture
'        ReDim MenuTransposition.Element1(1 To 17)
'        ReDim MenuTransposition.Element2(1 To 17)
'
'        MenuTransposition.Element1(1) = 2
'        MenuTransposition.Element2(1) = 0
'        MenuTransposition.Element1(2) = 3
'        MenuTransposition.Element2(2) = 13
'        MenuTransposition.Element1(3) = 4
'        MenuTransposition.Element2(3) = 9
'        MenuTransposition.Element1(4) = 5
'        MenuTransposition.Element2(4) = 10
'        MenuTransposition.Element1(5) = 6
'        MenuTransposition.Element2(5) = 11
'        MenuTransposition.Element1(6) = 7
'        MenuTransposition.Element2(6) = 12
'        MenuTransposition.Element1(7) = 8
'        MenuTransposition.Element2(7) = 13
'
'
'        MenuTransposition.Element1(8) = 10
'        MenuTransposition.Element2(8) = 1
'        MenuTransposition.Element1(9) = 11
'        MenuTransposition.Element2(9) = 2
'        MenuTransposition.Element1(10) = 12
'        MenuTransposition.Element2(10) = 3
'        MenuTransposition.Element1(11) = 13
'        MenuTransposition.Element2(11) = 4
'        MenuTransposition.Element1(12) = 14
'        MenuTransposition.Element2(12) = 5
'
'        MenuTransposition.Element1(13) = 16
'        MenuTransposition.Element2(13) = 6
'        MenuTransposition.Element1(14) = 17
'        MenuTransposition.Element2(14) = 7
'        MenuTransposition.Element1(15) = 18
'        MenuTransposition.Element2(15) = 8
'
'        MenuTransposition.Element1(16) = 20
'        MenuTransposition.Element2(16) = 14
'        MenuTransposition.Element1(17) = 21
'        MenuTransposition.Element2(17) = 15
'
'        .MenuBar(2).AddItem GetString(ResB_Points), 1, GetString(ResB_Points), , , LoadResPicture(ResIconB_Points, vbResIcon), GraphicsOnly
'            .MenuBar(2).AddItem GetString(ResFigurePoint), 2, GetString(ResFigurePoint), , , LoadResPicture(ResIconPoint, vbResIcon), TextAndGraphics
'            Menus(2).Items(2).Checked = True
'            .MenuBar(2).AddItem GetString(ResFigurePointOnFigure), 2, GetString(ResFigurePointOnFigure), , , LoadResPicture(ResIconPointOnFigure, vbResIcon), TextAndGraphics
'            .MenuBar(2).AddItem GetString(ResFigureMiddlePoint), 2, GetString(ResFigureMiddlePoint), , , LoadResPicture(ResIconMiddlePoint, vbResIcon), TextAndGraphics
'            .MenuBar(2).AddItem GetString(ResFigureSimmPoint), 2, GetString(ResFigureSimmPoint), , , LoadResPicture(ResIconSimmPoint, vbResIcon), TextAndGraphics
'            .MenuBar(2).AddItem GetString(ResFigureSimmPointByLine), 2, GetString(ResFigureSimmPointByLine), , , LoadResPicture(ResIconSimmPointByLine, vbResIcon), TextAndGraphics
'            .MenuBar(2).AddItem GetString(ResFigureInvert), 2, GetString(ResFigureInvert), , , LoadResPicture(ResIconInvert, vbResIcon), TextAndGraphics
'            .MenuBar(2).AddItem GetString(ResFigureIntersect), 2, GetString(ResFigureIntersect), , , LoadResPicture(ResIconIntersect, vbResIcon), TextAndGraphics
'
'        .MenuBar(2).AddItem GetString(ResB_Lines), 1, GetString(ResB_Lines), , , LoadResPicture(ResIconB_Lines, vbResIcon), GraphicsOnly
'            .MenuBar(2).AddItem GetString(ResFigureSegment), 2, GetString(ResFigureSegment), , , LoadResPicture(ResIconSegment, vbResIcon), TextAndGraphics
'            .MenuBar(2).AddItem GetString(ResFigureRay), 2, GetString(ResFigureRay), , , LoadResPicture(ResIconRay, vbResIcon), TextAndGraphics
'            .MenuBar(2).AddItem GetString(ResFigureLine_2Points), 2, GetString(ResFigureLine_2Points), , , LoadResPicture(ResIconLine_2Points, vbResIcon), TextAndGraphics
'            .MenuBar(2).AddItem GetString(ResFigureLine_PointAndParallelLine), 2, GetString(ResFigureLine_PointAndParallelLine), , , LoadResPicture(ResIconLine_PointAndParallelLine, vbResIcon), TextAndGraphics
'            .MenuBar(2).AddItem GetString(ResFigureLine_PointAndPerpendicularLine), 2, GetString(ResFigureLine_PointAndPerpendicularLine), , , LoadResPicture(ResIconLine_PointAndPerpendicularLine, vbResIcon), TextAndGraphics
'
'        .MenuBar(2).AddItem GetString(ResB_Circles), 1, GetString(ResB_Circles), , , LoadResPicture(ResIconB_Circles, vbResIcon), GraphicsOnly
'            .MenuBar(2).AddItem GetString(ResFigureCircle_CenterAndCircumPoint), 2, GetString(ResFigureCircle_CenterAndCircumPoint), , , LoadResPicture(ResIconCircle_CenterAndCircumPoint, vbResIcon), TextAndGraphics
'            .MenuBar(2).AddItem GetString(ResFigureCircle_CenterAndTwoPoints), 2, GetString(ResFigureCircle_CenterAndTwoPoints), , , LoadResPicture(ResIconCircle_CenterAndTwoPoints, vbResIcon), TextAndGraphics
'            .MenuBar(2).AddItem GetString(ResFigureCircle_ArcCenterAndRadiusAndTwoPoints), 2, GetString(ResFigureCircle_ArcCenterAndRadiusAndTwoPoints), , , LoadResPicture(ResIconCircle_ArcCenterAndRadiusAndTwoPoints, vbResIcon), TextAndGraphics
'
'        .MenuBar(2).AddItem GetString(ResB_Measurements), 1, GetString(ResB_Measurements), , , LoadResPicture(ResIconB_Measurements, vbResIcon), GraphicsOnly
'            .MenuBar(2).AddItem GetString(ResFigureMeasureDistance), 2, GetString(ResFigureMeasureDistance), , , LoadResPicture(ResIconMeasureDistance, vbResIcon), TextAndGraphics
'            .MenuBar(2).AddItem GetString(ResFigureMeasureAngle), 2, GetString(ResFigureMeasureAngle), , , LoadResPicture(ResIconMeasureAngle, vbResIcon), TextAndGraphics
    'Else '____________________________________________________________________


'    .MenuBar(2).AddItem GetString(ResSelect), 1, GetString(ResSelect), , , LoadResPicture(ResIconCursor, vbResIcon), GraphicsOnly
'    For Z = dsPoint To dsMeasureAngle
'        If Z = 7 Or Z = 12 Then BAG = True Else BAG = False
'        .MenuBar(2).AddItem GetString(ResFigureBase + Z * 2), 1, GetString(ResFigureBase + Z * 2), , , LoadResPicture(ResIconBase + Z, vbResIcon), GraphicsOnly, BAG
'    Next Z
'    Menus(2).Items(1).Checked = True
    'End If
End With
End Sub

Public Sub FillMenuBar(Optional ByVal MenuBarMode As MenuBarState = mbsToolBar, Optional ByVal MenuBarIndex As Long = 2, Optional ByVal ShouldClear As Boolean = True)
Dim Z As Long, BAG As Boolean 'begin a group

If ShouldClear Then FormMain.MenuBar(MenuBarIndex).Clear

Select Case MenuBarIndex
Case 1
    'FormMain.MenuBar(1).AddItem GetString(ResFile), 1
    FormMain.MenuBar(1).AddItem GetString(ResNew), 1, GetString(ResMnuCreateNew) & " (Ctrl+N)", , , GetPicture(ResIconMenuNew, , 0), GraphicsOnly, , , ResNew
    FormMain.MenuBar(1).AddItem GetString(ResOpen), 1, GetString(ResMnuOpenExisting) & " (Ctrl+O)", , , GetPicture(ResIconMenuOpen, , 0), GraphicsOnly, , , ResOpen
    'If CanSave Then
    FormMain.MenuBar(1).AddItem GetString(ResSave), 1, GetString(ResMnuSave) & " (Ctrl+S)", , , GetPicture(ResIconMenuSave, , 0), GraphicsOnly, , , ResSave
    FormMain.MenuBar(1).AddItem GetString(ResUndo), 1, GetString(ResUndo) & " (Ctrl+Z)", , , GetPicture(ResIconMenuUndo, , 0), GraphicsOnly, True, , ResUndo
    FormMain.MenuBar(1).AddItem GetString(ResRedo), 1, GetString(ResRedo) & " (Ctrl+R)", , , GetPicture(ResIconMenuRedo, , 0), GraphicsOnly, , , ResRedo
    'FormMain.MenuBar(1).AddItem GetString(ResFigureBase + 2 * dsDynamicLocus), 1, GetString(ResFigureBase + 2 * dsDynamicLocus), , , GetPicture(ResIconDynamicLocus), GraphicsOnly, True, , ResFigureBase + 2 * dsDynamicLocus
    'FormMain.MenuBar(1).AddItem GetString(ResStaticObjectBase + 2 * sgPolygon), 1, GetString(ResStaticObjectBase + 2 * sgPolygon), , , GetPicture(ResIconPolygon), GraphicsOnly, , , ResStaticObjectBase + 2 * sgPolygon
    FormMain.MenuBar(1).AddItem GetString(ResInsertLabel), 1, GetString(ResInsertLabel) & " (Enter)", , , GetPicture(ResIconLabel), GraphicsOnly, True, , ResInsertLabel
    FormMain.MenuBar(1).AddItem GetString(ResInsertButton), 1, GetString(ResInsertButton), , , GetPicture(ResIconButton), GraphicsOnly, , , ResInsertButton
    FormMain.MenuBar(1).AddItem GetString(ResCalculator), 1, GetString(ResCalculator) & " (F5)", , , GetPicture(ResIconCalc, vbResBitmap), GraphicsOnly, True, , ResCalcBase
    FormMain.MenuBar(1).AddItem GetString(ResMnuMacroCreate), 1, GetString(ResMnuMacroCreate), , , GetPicture(ResIconM), GraphicsOnly, , , ResMnuMacroCreate
'    FormMain.MenuBar(1).AddItem GetString(ResDemo), 1, GetString(ResDemo) & " (F8)", , , GetPicture(ResIconCamera), GraphicsOnly, , , ResDemo
    FormMain.MenuBar(1).AddItem GetString(ResOptions), 1, GetString(ResOptions) & " (F9)", , , GetPicture(ResIconSettings), GraphicsOnly, , , ResOptions
    FormMain.MenuBar(1).AddItem GetString(ResHelp), 1, GetString(ResHelp), , , GetPicture(ResIconHelp), GraphicsOnly, True, , ResHelpContents
    
    'FormMain.MenuBar(1).AddItem GetString(ResStaticObjectBase + 2 * sgVector), 1, GetString(ResStaticObjectBase + 2 * sgVector), , , GetPicture(ResIconVector), GraphicsOnly, , , ResStaticObjectBase + 2 * sgVector
    'FormMain.MenuBar(1).AddItem GetString(ResStaticObjectBase + 2 * sgBezier), 1, GetString(ResStaticObjectBase + 2 * sgBezier), , , GetPicture(ResIconBezier), GraphicsOnly, , , ResStaticObjectBase + 2 * sgBezier

Case 2
    
    'FillMenuTransposition
    
    Select Case MenuBarMode
    Case mbsToolBar
        'FormMain.MenuBar(MenuBarIndex).AddItem GetString(ResSelect), 1, GetString(ResSelect), , , GetPicture(dsSelect), setDisplayIconText
        
        For Z = 1 To MenuTransposition.Count
            Select Case Transpose(Z, MenuTransposition)
            Case dsSegment, dsCircle_CenterAndCircumPoint, dsMiddlePoint, dsMeasureDistance, dsDynamicLocus
                BAG = True
            Case Else
                BAG = False
            End Select
            FormMain.MenuBar(MenuBarIndex).AddItem GetString(ResFigureBase + 2 * Transpose(Z, MenuTransposition)), 1, GetString(ResFigureBase + 2 * Transpose(Z, MenuTransposition)), , , GetPicture(Z - 2), setDisplayIconText, BAG, , , , GetToolShortcut(Z)
            'FormMain.MenuBar(MenuBarIndex).AddItem GetString(ResFigureBase + 2 * Transpose(Z, MenuTransposition)), 1, GetString(ResFigureBase + 2 * Transpose(Z, MenuTransposition)), , , GetPicture(Transpose(Z, MenuTransposition)), setDisplayIconText, BAG
        Next
        
        'FormMain.MenuBar(MenuBarIndex).AddItem "Area", 1, "Measure area", , , GetPicture(18), GraphicsOnly
        
        'Menus(MenuBarIndex).Items(1).Checked = True
        FormMain.MenuBar(MenuBarIndex).CheckItem 1, True
        FormMain.MenuBar(MenuBarIndex).ToolbarMode = mbsToolBar
        
    
    Case mbsSelectObjectsFinish
        FormMain.MenuBar(MenuBarIndex).AddItem GetString(ResContinue), 1, GetString(ResContinue), , , GetPicture(ResIconOK), TextAndGraphics
        FormMain.MenuBar(MenuBarIndex).AddItem GetString(ResCancel), 1, GetString(ResCancel), , , GetPicture(ResIconCancel), TextAndGraphics, True
        FormMain.MenuBar(MenuBarIndex).ToolbarMode = mbsSelectObjectsFinish
    
    Case mbsDemo
        FormMain.MenuBar(MenuBarIndex).AddItem GetString(ResStepFirst), 1, GetString(ResStepFirst), , , GetPicture(ResIconFrameFirst), GraphicsOnly
        FormMain.MenuBar(MenuBarIndex).AddItem GetString(ResStepPrevious), 1, GetString(ResStepPrevious), , , GetPicture(ResIconFramePrevious), GraphicsOnly
        FormMain.MenuBar(MenuBarIndex).AddItem GetString(ResStepNext), 1, GetString(ResStepNext), , , GetPicture(ResIconFrameNext), GraphicsOnly
        FormMain.MenuBar(MenuBarIndex).AddItem GetString(ResStepLast), 1, GetString(ResStepLast), , , GetPicture(ResIconFrameLast), GraphicsOnly
        FormMain.MenuBar(MenuBarIndex).AddItem GetString(ResAutoDemo), 1, GetString(ResAutoDemo), , , GetPicture(ResIconCamera), TextAndGraphics, True
        FormMain.MenuBar(MenuBarIndex).AddItem GetString(ResUsualMode), 1, GetString(ResUsualMode), , , GetPicture(ResIconOK), TextAndGraphics, True
        FormMain.MenuBar(MenuBarIndex).ToolbarMode = mbsDemo
    
    Case mbsMacroGivens
        FormMain.MenuBar(MenuBarIndex).AddItem GetString(ResMnuMacroSelectResults), 1, GetString(ResMnuMacroSelectResults), , , GetPicture(ResIconOK), TextAndGraphics
        FormMain.MenuBar(MenuBarIndex).AddItem GetString(ResCancel), 1, GetString(ResCancel), , , GetPicture(ResIconCancel), TextAndGraphics, True
        FormMain.MenuBar(MenuBarIndex).ToolbarMode = mbsMacroGivens
        
    Case mbsMacroResults
        FormMain.MenuBar(MenuBarIndex).AddItem GetString(ResMnuMacroSave), 1, GetString(ResMnuMacroSave), , , GetPicture(ResIconOK), TextAndGraphics
        FormMain.MenuBar(MenuBarIndex).AddItem GetString(ResCancel), 1, GetString(ResCancel), , , GetPicture(ResIconCancel), TextAndGraphics, True
        FormMain.MenuBar(MenuBarIndex).ToolbarMode = mbsMacroResults
        
    Case mbsMacroRun
        FormMain.MenuBar(MenuBarIndex).AddItem GetString(ResCancel), 1, GetString(ResCancel), , , GetPicture(ResIconCancel), TextAndGraphics
        FormMain.MenuBar(MenuBarIndex).ToolbarMode = mbsMacroRun
        
    Case mbsCancel
        FormMain.MenuBar(MenuBarIndex).AddItem GetString(ResCancel), 1, GetString(ResCancel), , , GetPicture(ResIconCancel), TextAndGraphics
        FormMain.MenuBar(MenuBarIndex).ToolbarMode = mbsCancel
        
    End Select
    
    'Select Case MenuBarMode
    'Case mbsToolBar
    '    FormMain.MenuBar(MenuBarIndex).AddItem GetString(ResSelect), 1, GetString(ResSelect), , , GetPicture(dsSelect), setDisplayIconText
    '    For Z = dsPoint To dsMeasureAngle
    '        If Z = dsSegment Or Z = dsCircle_CenterAndCircumPoint Or Z = dsMiddlePoint Or Z = dsMeasureDistance Then BAG = True Else BAG = False
    '        FormMain.MenuBar(MenuBarIndex).AddItem GetString(ResFigureBase + Z * 2), 1, GetString(ResFigureBase + Z * 2), , , GetPicture(Z), setDisplayIconText, BAG
    '        If Z = dsLine_PointAndPerpendicularLine Then
    '            FormMain.MenuBar(MenuBarIndex).AddItem GetString(ResFigureBase + dsBisector * 2), 1, GetString(ResFigureBase + dsBisector * 2), , , GetPicture(dsBisector), setDisplayIconText
    '        End If
    '    Next Z
    '    Menus(MenuBarIndex).Items(1).Checked = True
    '    FormMain.MenuBar(MenuBarIndex).ToolbarMode = mbsToolBar
    '
    'Case mbsSelectObjectsFinish
    '    FormMain.MenuBar(MenuBarIndex).AddItem GetString(ResContinue), 1, GetString(ResContinue)
    '    FormMain.MenuBar(MenuBarIndex).AddItem GetString(ResCancel), 1, GetString(ResCancel), , , , , True
    '    FormMain.MenuBar(MenuBarIndex).ToolbarMode = mbsSelectObjectsFinish
    'End Select
End Select '// case of which MenuBar?

FormMain.MenuBar(MenuBarIndex).Refresh
End Sub

Public Sub FillMenuTransposition(Optional ByVal MenuBarMode As MenuBarState = mbsToolBar, Optional ByVal MenuBarIndex As Long = 2, Optional ByVal ShouldClear As Boolean = True)
Dim Z As Long
With MenuTransposition
    TranspositionClear MenuTransposition
    
    Select Case MenuBarMode
    Case mbsToolBar
        
        TranspositionAdd 1, dsSelect, MenuTransposition
        
        For Z = dsPoint To dsLine_PointAndPerpendicularLine
            TranspositionAdd Z + 2, Z, MenuTransposition
        Next
        
        TranspositionAdd dsLine_PointAndPerpendicularLine + 3, dsBisector, MenuTransposition
        
        For Z = dsCircle_CenterAndCircumPoint To dsPointOnFigure
            TranspositionAdd Z + 3, Z, MenuTransposition
        Next
        
        TranspositionAdd dsMeasureDistance + 3, dsDynamicLocus, MenuTransposition
        TranspositionAdd dsMeasureDistance + 4, dsPolygon, MenuTransposition
        
        TranspositionAdd dsMeasureDistance + 5, dsMeasureDistance, MenuTransposition
        TranspositionAdd dsMeasureAngle + 5, dsMeasureAngle, MenuTransposition
        TranspositionAdd dsMeasureAngle + 6, dsMeasureArea, MenuTransposition
    
    Case mbsSelectObjectsFinish
    End Select
End With
End Sub

Public Sub FillMRU()
Dim Z As Long
Dim MRUCount As Long, tStr As String, ShouldResave As Boolean
On Local Error Resume Next

Do While MRUList.Count > 0
    MRUList.Remove 1
Loop

MRUCount = Val(GetSetting(AppName, "MRU", "Count", "-1"))
If MRUCount <> -1 Then
    For Z = 1 To MRUCount
        tStr = GetSetting(AppName, "MRU", "File" & Z)
        If tStr <> "" And Dir(tStr) <> "" Then
            If ERR.Number = 0 Then MRUList.Add tStr
        End If
        ERR.Clear
    Next Z
    If MRUList.Count < MRUCount Then ShouldResave = True
    MRUCount = MRUList.Count
End If

If ShouldResave Then SaveMRU Else UpdateMRUMenu
End Sub

Public Sub SaveMRU()
Dim Z As Long, bRemoved As Boolean
On Local Error Resume Next

Z = 0
Do While Z < MRUList.Count
    Z = Z + 1
    bRemoved = False
    If Dir(MRUList(Z)) = "" Then MRUList.Remove Z: Z = Z - 1: bRemoved = True
    If ERR.Number = 52 And Not bRemoved Then MRUList.Remove Z: Z = Z - 1: ERR.Clear
Loop

SaveSetting AppName, "MRU", "Count", MRUList.Count
If MRUList.Count > 0 Then
    For Z = 1 To MRUList.Count
        SaveSetting AppName, "MRU", "File" & Z, MRUList(Z)
    Next Z
End If
UpdateMRUMenu
End Sub

Public Sub UpdateMRUMenu()
Dim Z As Long
Do While FormMain.mnuMRUFile.UBound > MRUList.Count And FormMain.mnuMRUFile.UBound > 1: Unload FormMain.mnuMRUFile(FormMain.mnuMRUFile.UBound): Loop

With FormMain
    If MRUList.Count > 0 Then
        .mnuFileSep4.Visible = True
        For Z = 1 To MRUList.Count
            If .mnuMRUFile.UBound < Z Then Load .mnuMRUFile(Z)
            .mnuMRUFile(Z).Caption = "&" & Z & " " & RetrieveName(MRUList(MRUList.Count + 1 - Z))
            .mnuMRUFile(Z).Visible = True
        Next
    Else
        .mnuMRUFile(1).Visible = False
        .mnuFileSep4.Visible = False
    End If
End With
End Sub

Public Sub AddMRUItem(ByVal tStr As String)
Dim Z As Long
Z = 1
Do While Z <= MRUList.Count
    If LCase(MRUList(Z)) = LCase(tStr) Then MRUList.Remove Z Else Z = Z + 1
Loop
If MRUList.Count >= MRUMax Then MRUList.Remove 1
MRUList.Add tStr
SaveMRU
End Sub

Public Sub SelectActiveTool(ByVal TheTool As DrawState)
With FormMain
    .mnuToolPoint.Checked = TheTool = dsPoint
    .mnuToolSegment.Checked = TheTool = dsSegment
    .mnuToolRay.Checked = TheTool = dsRay
    .mnuToolLine.Checked = TheTool = dsLine_2Points
    .mnuToolParallelLine.Checked = TheTool = dsLine_PointAndParallelLine
    .mnuToolPerpendicularLine.Checked = TheTool = dsLine_PointAndPerpendicularLine
    .mnuToolCircle.Checked = TheTool = dsCircle_CenterAndCircumPoint
    .mnuToolCircleByRadius.Checked = TheTool = dsCircle_CenterAndTwoPoints
    .mnuToolArc.Checked = TheTool = dsCircle_ArcCenterAndRadiusAndTwoPoints
    .mnuToolMiddlePoint.Checked = TheTool = dsMiddlePoint
    .mnuToolSymmPoint.Checked = TheTool = dsSimmPoint
    .mnuToolReflectedPoint.Checked = TheTool = dsSimmPointByLine
    .mnuToolIntersect.Checked = TheTool = dsIntersect
    .mnuToolInvert.Checked = TheTool = dsInvert
    .mnuToolPointOnFigure.Checked = TheTool = dsPointOnFigure
    .mnuToolMeasureDistance.Checked = TheTool = dsMeasureDistance
    .mnuToolMeasureAngle.Checked = TheTool = dsMeasureAngle
End With
End Sub

Public Function GetToolShortcut(ByVal Index As Long) As Long
Select Case Index
Case 1
    GetToolShortcut = vbKeyQ
Case 2
    GetToolShortcut = vbKeyP
Case 3
    GetToolShortcut = vbKeyS
Case 4
    GetToolShortcut = vbKeyY
Case 5
    GetToolShortcut = vbKeyL
Case 6
    GetToolShortcut = vbKeyN
Case 7
    GetToolShortcut = vbKeyE
Case 8
    GetToolShortcut = vbKeyB
Case 9
    GetToolShortcut = vbKeyC
Case 10
    GetToolShortcut = vbKeyR
Case 11
    GetToolShortcut = vbKeyA
Case 12
    GetToolShortcut = vbKeyM
Case 13
    GetToolShortcut = vbKeyO
Case 14
    GetToolShortcut = vbKeyT
Case 15
    GetToolShortcut = vbKeyV
Case 16
    GetToolShortcut = vbKeyI
Case 17
    GetToolShortcut = vbKeyZ
Case 18
    GetToolShortcut = vbKeyD
Case 19
    GetToolShortcut = vbKeyW
Case 20
    GetToolShortcut = vbKeyG
Case 21
    GetToolShortcut = vbKeyJ
Case 22
    GetToolShortcut = vbKeyK
End Select
End Function
