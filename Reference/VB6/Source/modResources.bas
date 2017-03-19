Attribute VB_Name = "modResources"
Option Explicit

'=========================================================
'                              FIGURE NAMES and ICON IDs
'=========================================================

Public Const ResFigureBase = 400
Public Const ResSelect = 398

Public Enum ResFigureNames
    ResFigurePoint = ResFigureBase + dsPoint * 2
    ResFigureSegment = ResFigureBase + dsSegment * 2
    ResFigureRay = ResFigureBase + dsRay * 2
    ResFigureLine_2Points = ResFigureBase + dsLine_2Points * 2
    ResFigureLine_PointAndParallelLine = ResFigureBase + dsLine_PointAndParallelLine * 2
    ResFigureLine_PointAndPerpendicularLine = ResFigureBase + dsLine_PointAndPerpendicularLine * 2
    ResFigureCircle_CenterAndCircumPoint = ResFigureBase + dsCircle_CenterAndCircumPoint * 2
    ResFigureCircle_CenterAndTwoPoints = ResFigureBase + dsCircle_CenterAndTwoPoints * 2
    ResFigureCircle_ArcCenterAndRadiusAndTwoPoints = ResFigureBase + dsCircle_ArcCenterAndRadiusAndTwoPoints * 2
    ResFigureMiddlePoint = ResFigureBase + dsMiddlePoint * 2
    ResFigureSimmPoint = ResFigureBase + dsSimmPoint * 2
    ResFigureSimmPointByLine = ResFigureBase + dsSimmPointByLine * 2
    ResFigureInvert = ResFigureBase + dsInvert * 2
    ResFigureIntersect = ResFigureBase + dsIntersect * 2
    ResFigurePointOnFigure = ResFigureBase + dsPointOnFigure * 2
    ResFigureMeasureDistance = ResFigureBase + dsMeasureDistance * 2
    ResFigureMeasureAngle = ResFigureBase + dsMeasureAngle * 2
    ResFigureAnLineGeneral = ResFigureBase + dsAnLineGeneral * 2
    ResFigureAnLineCanonic = ResFigureBase + dsAnLineCanonic * 2
    ResFigureAnLineNormal = ResFigureBase + dsAnLineNormal * 2
    ResFigureAnLineNormalPoint = ResFigureBase + dsAnLineNormalPoint * 2
    ResFigureAnCircle = ResFigureBase + dsAnCircle * 2
    ResFigureBisector = ResFigureBase + dsBisector * 2
End Enum

Public Const ResIconOK = 50
Public Const ResIconCancel = 51
Public Const ResIconFrameFirst = 52
Public Const ResIconFramePrevious = 53
Public Const ResIconFrameNext = 54
Public Const ResIconFrameLast = 55
Public Const ResIconCamera = 56
 
Public Const ResIconLabel = 60
Public Const ResIconButton = 61
Public Const ResIconDynamicLocus = 62
Public Const ResIconPolygon = 63
Public Const ResIconBezier = 64
Public Const ResIconVector = 65
Public Const ResIconCalc = 66
Public Const ResIconM = 67
Public Const ResIconSettings = 68
Public Const ResIconHelp = 69

Public Const ResIconMenuNew = 101
Public Const ResIconMenuOpen = 102
Public Const ResIconMenuSave = 103
Public Const ResIconMenuUndo = 104
Public Const ResIconMenuRedo = 105

Public Const ResBMPCamera = 120

'=========================================================
'                           RESOURCE STRING IDENTIFIERS
'=========================================================
Public Const ResCaption = 102
Public Const ResFile = 104
Public Const ResEdit = 106
Public Const ResView = 108
Public Const ResClearAll = 110
Public Const ResComment = 112
Public Const ResOptions = 114
Public Const ResHelp = 116
Public Const ResNumberOfItems = 118
Public Const ResNew = 120
Public Const ResOpen = 122
Public Const ResSave = 124
Public Const ResSaveAs = 126
Public Const ResExit = 128
Public Const ResExport = 130
Public Const ResBMP = 132
Public Const ResInsertLabel = 134
Public Const ResLabelProperties = 136
Public Const ResFileProps = 138
Public Const ResUndo = 140
Public Const ResShow = 142
Public Const ResWEWindow = 146
Public Const ResMnuAnalytic = 148
Public Const ResMnuAnPoint = 150
Public Const ResMnuAnLine = 152
Public Const ResMnuAnCircle = 154
Public Const ResMnuFigures = 156
Public Const ResCanonic = 158
Public Const ResNormal = 160
Public Const ResLinePoint = 162
Public Const ResGuideVector = 164
Public Const ResLineEquation = 166
Public Const ResNormalVector = 168
Public Const ResBold = 170
Public Const ResItalic = 172
Public Const ResUnderline = 174
Public Const ResTransparent = 176
Public Const ResSize = 178
Public Const ResColor = 180
Public Const ResBackColor = 182
Public Const ResShowCoordinates = 184
Public Const ResFont = 186
Public Const ResWMF = 188
Public Const ResEMF = 190
Public Const ResPressF1ForHelp = 192
Public Const ResWhatsThis = 194
Public Const ResRightClickForHelp = 196
Public Const ResRedo = 198
Public Const ResB_Points = 200
Public Const ResB_Lines = 202
Public Const ResB_Circles = 204
Public Const ResB_Measurements = 206
Public Const ResCancel = 208
Public Const ResClock = 210
Public Const ResFigureList = 212
Public Const ResPointList = 214
Public Const ResFullscreen = 216
Public Const ResPrint = 218
Public Const ResName = 220
Public Const ResEquation = 222
Public Const ResAppearance = 224
Public Const ResTipOfTheDay = 226
Public Const ResWallpaper = 228
Public Const ResStatistics = 230
Public Const ResLabels = 232
Public Const ResActiveAxes = 234
Public Const ResXAxis = 236
Public Const ResYAxis = 238
Public Const ResShowAxesMarks = 240
Public Const ResRulerColor = 242
Public Const ResBytes = 244
Public Const ResToolbarColor = 246
Public Const ResJSPHTML = 248
Public Const ResContinue = 250
Public Const ResLabel = 252
Public Const ResLocus = 254
Public Const ResDemo = 256
Public Const ResUsualMode = 258
Public Const ResMeasureDistance = 260
Public Const ResMeasureAngle = 262
Public Const ResDisabled = 264
Public Const ResLocusDetails = 266
Public Const ResLocusDetailsHigh = 268
Public Const ResLocusType = 270
Public Const ResStepFirst = 272
Public Const ResStepPrevious = 274
Public Const ResStepNext = 276
Public Const ResStepLast = 278
Public Const ResEmptyDemo = 280
Public Const ResAutoDemo = 282
Public Const ResDefault = 284
Public Const ResShowMainbar = 286
Public Const ResArea = 288
Public Const ResOtherColor = 290
Public Const ResSaveTransparentEMF = 292
Public Const ResCalculator = 294
Public Const ResAngleMark = 296
Public Const ResHideMeasurementText = 298
Public Const ResNone = 300
Public Const ResMnuMacros = 302
Public Const ResMnuMacroCreate = 304
Public Const ResMnuMacroSelectResults = 306
Public Const ResMnuMacroSave = 308
Public Const ResMnuMacroLoad = 310
Public Const ResMnuMacroOrganize = 312
Public Const ResMnuMacroProperties = 314
Public Const ResMacro = 316
Public Const ResLocateObject = 318
Public Const ResButtonProperties = 320
Public Const ResCaptionName = 322
Public Const ResType = 324
Public Const ResButton = 326
Public Const ResInsertButton = 328
Public Const ResShowHideObjects = 330
Public Const ResObjectList = 332
Public Const ResSelectObjects = 334
Public Const ResInitiallyVisible = 336
Public Const ResObjectsNotSelected = 338
Public Const ResMoveButton = 340
Public Const ResDeleteButton = 342
Public Const ResFix = 344
Public Const ResMessageButton = 346
Public Const ResButtons = 348
Public Const ResHelpContents = 350
Public Const ResAbout = 352
Public Const ResFill = 354
Public Const ResPlaySound = 356
Public Const ResLaunchFile = 358
Public Const ResMnuChooseFigure = 360
Public Const ResMnuChoosePoint = 362
Public Const ResAddRemoveObjects = 364
Public Const ResCreateLabelWithMeasurement = 366
Public Const ResPropertiesOfAPoint = 368
Public Const ResMnuLabelProperties = 370
Public Const ResMnuDeleteLabel = 372
Public Const ResMnuRecalcLabel = 374
Public Const ResMnuDeleteMeasurement = 376
Public Const ResMoveMeasurement = 378
Public Const ResMnuDeleteFigure = 380
Public Const ResMnuFigureProperties = 382
Public Const ResMnuDeletePoint = 384
Public Const ResMnuPointProperties = 386
Public Const ResMnuSnapToFigure = 388
Public Const ResMnuReleasePoint = 390
Public Const ResMnuDeleteObject = 392
Public Const ResMnuMeasurementProperties = 394
Public Const ResApply = 396

Public Const ResFillStyleBase = 480

Public Const ResSetPoints = 500
Public Const ResSetFigures = 502
Public Const ResSetPaper = 504
Public Const ResSetInterface = 506
Public Const ResSetMisc = 508

Public Const ResAutoShowPointName = 522
Public Const ResColors = 524
Public Const ResPaperColor = 526
Public Const ResGridColor = 528
Public Const ResAxesColor = 530
Public Const ResMacroAutoloadPath = 532
Public Const ResLocateMacroAutoloadPath = 534
Public Const ResPleaseWaitLoadingMacros = 536
Public Const ResLoading = 538
Public Const ResSaving = 540
Public Const ResShowStatusbar = 542
Public Const ResShowToolbar = 544
Public Const ResShowTooltips = 546
Public Const ResRightClickForMenu = 548
Public Const ResUndoActionBase = 550

Public Const ResStaticObjects = 650
Public Const ResStaticObjectBase = 652

Public Const ResDigitNumber = 674
Public Const ResDecimalPrecision = 676 '508
Public Const ResDistancePrecision = 678 '510
Public Const ResAnglePrecision = 680 '512
Public Const ResNumberPrecision = 682 '514
Public Const ResShowAxes = 684 '516
Public Const ResShowGrid = 686 '518
Public Const ResShowRulers = 688 '520
'===================================
Public Const ResSetNewPointProperties = 690 ' // added Dec 14
Public Const ResSetBasePoints = 692
Public Const ResSetFigurePoints = 694
Public Const ResSetDependentPoints = 696
Public Const ResSetPointNames = 698
Public Const ResDrawModeSolid = 700
Public Const ResDrawModeTransparentOnLight = 702
Public Const ResDrawModeTransparentOnDark = 704
Public Const ResDrawModeInvert = 706

Public Const ResDeletePolygon = 710
Public Const ResDeleteBezier = 712
Public Const ResDeleteVector = 714

Public Const ResPropertiesOfAPolygon = 720
Public Const ResPropertiesOfABezier = 722
Public Const ResPropertiesOfAVector = 724

Public Const ResMacroCreation = 730
Public Const ResDoNotShowThisDialog = 732
Public Const ResMnuMacroSelectGivens = 734
Public Const ResMacroGivenPrompt = 736
Public Const ResMacroResultSelection = 738
Public Const ResMacroChosenSuchGivens = 740
Public Const ResMacroChosenNoGivens = 742
Public Const ResReturnAndReselect = 744
Public Const ResMacroThisGivenPrompt = 746
Public Const ResNowSelectResults = 748
Public Const ResMacroNeverthelessSelectResults = 750
Public Const ResMacroNoResultsWithoutGivens = 752
Public Const ResMacroSelectResults = 754
Public Const ResMacroCompletingTask = 756
Public Const ResMacroChosenSuchResults = 758
Public Const ResMacroEnterNameAndDescription = 760
Public Const ResDescription = 762
Public Const ResMacroOKSaveAs = 764
Public Const ResMacroList = 766
Public Const ResMacroAdd = 768
Public Const ResMacroRemove = 770
Public Const ResMacroDuringCreation = 772
Public Const ResMacroShowCreateDialog = 774
Public Const ResMacroShowResultsDialog = 776
Public Const ResSaveAsDefaults = 778
Public Const ResLoadDefaults = 780
Public Const ResPropertiesOfPoints = 782
Public Const ResSelectPointFromList = 784
Public Const ResSelectFigureFromList = 786
Public Const ResPropertiesOfFigures = 788
Public Const ResFigures = 790
Public Const ResDGFFiles = 792
Public Const ResDGMFiles = 794

Public Const ResCalcBase = 800
Public Const ResCalcDistance = 0
Public Const ResCalcDistanceB = 1
Public Const ResCalcAngle = 2
Public Const ResCalcAngleB = 3
Public Const ResCalcOAngle = 4
Public Const ResCalcOAngleB = 5
Public Const ResCalcXangle = 6
Public Const ResCalcXangleB = 7
Public Const ResCalcX = 8
Public Const ResCalcXB = 9
Public Const ResCalcY = 10
Public Const ResCalcYB = 11
Public Const ResCalcNorm = 12
Public Const ResCalcNormB = 13
Public Const ResCalcArg = 14
Public Const ResCalcArgB = 15
Public Const ResCalcArea = 16
Public Const ResCalcAreaB = 17
Public Const ResCalcIf = 18
Public Const ResCalcMax = 19
Public Const ResCalcMin = 20
Public Const ResCalcDegree = 21
Public Const ResCalcPi = 22
Public Const ResCalcE = 23
Public Const ResCalcPower = 24
Public Const ResCalcSqr = 25
Public Const ResCalcLog = 26
Public Const ResCalcExp = 27
Public Const ResCalcAbs = 28
Public Const ResCalcSin = 29
Public Const ResCalcCos = 30
Public Const ResCalcTg = 31
Public Const ResCalcCtg = 32
Public Const ResCalcRound = 33
Public Const ResCalcFact = 34
Public Const ResCalcRnd = 35
Public Const ResCalcRandom = 36
Public Const ResCalcRad = 37
Public Const ResCalcBrackets = 38
Public Const ResCalcASin = 39
Public Const ResCalcACos = 40
Public Const ResCalcATg = 41
Public Const ResCalcACtg = 42

Public Const ResGivenHintBase = 890
Public Const ResGivenHintPoint = 890
Public Const ResGivenHintSegmentRayLine = 892
Public Const ResGivenHintCircle = 894

Public Const ResTitles = 900
Public Const ResSupport = 940
Public Const ResDemoOptions = 942
Public Const ResDemoParticipatingItems = 944
Public Const ResDemoStepDescription = 946
Public Const ResDemoDelay = 948
Public Const ResSeconds = 950
Public Const ResRenamePoint = 952 ' // Added March 12
Public Const ResRenameWhatToDo = 954
Public Const ResRenameChooseAnotherName = 956
Public Const ResRenameAssignAnother = 958
Public Const ResRenameSwapNames = 960

Public Const ResMsgNew = 1000
Public Const ResMsgInputInteger = 1002
Public Const ResMsgExpressedBy = 1004
Public Const ResMsgFrom = 1006
Public Const ResMsgTo = 1008
Public Const ResMsgNoName = 1010
Public Const ResMsgCantDragThisPoint = 1012
Public Const ResMsgCannotOpenFile = 1014
Public Const ResCurrentTool = 1016
Public Const ResConvertCoordConfirm = 1018
Public Const ResMsgCannotEvaluate = 1020
Public Const ResMsgABUnequalTo0 = 1022
Public Const ResImaginaryCircle = 1024
Public Const ResMsgObjectAlreadyExists = 1026
Public Const ResError = 1028
Public Const ResMsgDoYouReallyWantToCancel = 1030
Public Const ResHelpFileNotFound = 1032
Public Const ResIndependentColor = 1034
Public Const ResClickToClosePolygon = 1036
Public Const ResMovePointName = 1038
Public Const ResRemindToSaveFile = 1040
Public Const ResLoadCursors = 1042
Public Const ResCursorSensitivity = 1044
Public Const ResSetPaperGradient = 1046
Public Const ResSetGroupTools = 1048
Public Const ResMsgConfirmation = 1050

Public Const ResSetNoFlicker = 1054
Public Const ResSetGradientFill = 1056
Public Const ResSetSelectToolOnce = 1058
Public Const ResUntitled = 1060
Public Const ResPropsTitle = 1062
Public Const ResCreateLocus = 1064
Public Const ResDeleteLocus = 1066
Public Const ResLocusProps = 1068
Public Const ResDrawMode = 1070
Public Const ResDrawStyle = 1072
Public Const ResDrawWidth = 1074
Public Const ResForeColor = 1076
Public Const ResVisible = 1078
Public Const ResTheLength = 1080
Public Const ResTheRadius = 1082
Public Const ResFillColor = 1084
Public Const ResShape = 1086
Public Const ResHide = 1088
Public Const ResShowName = 1090
Public Const ResNameColor = 1092
Public Const ResLocusColor = 1094
Public Const ResFigureColor = 1096
Public Const ResDependentColor = 1098
Public Const ResAddPoint = 1100
Public Const ResDragPoint = 1102
Public Const ResLocate = 1104
Public Const ResFirst = 1106
Public Const ResSecond = 1108
Public Const ResThird = 1110
Public Const ResFourth = 1112
Public Const ResFifth = 1114
Public Const ResPoint = 1116
Public Const ResLineRayOrSegment = 1118
Public Const ResCircleCenter = 1120
Public Const ResRadiusStart = 1122
Public Const ResRadiusEnding = 1124
Public Const ResArcStart = 1126
Public Const ResArcEnd = 1128
Public Const ResSymmCenter = 1130
Public Const ResFigure = 1132
Public Const ResPressEsc = 1134
Public Const ResCreateParLine = 1136
Public Const ResCreatePerpLine = 1138
Public Const ResFasten = 1140
Public Const ResCreateSegment = 1142
Public Const ResCreateRay = 1144
Public Const ResCreateLine = 1146
Public Const ResCreateCircle = 1148
Public Const ResCreateArc = 1150
Public Const ResCreateMiddlePoint = 1152
Public Const ResCreateSymmetricPoint = 1154
Public Const ResCreateIntersection = 1156
Public Const ResClickToDragThisBar = 1158
Public Const ResCloseTheApp = 1160
Public Const ResMinimizeTheApp = 1162
Public Const ResCreatePointOnFigure = 1164
Public Const ResLocateAngleVertex = 1166
Public Const ResWorkingPleaseWait = 1168
Public Const ResMovePoint = 1170
Public Const ResMoveLabel = 1172
Public Const ResShiftSnapToGrid = 1174
Public Const ResCtrlSnapToFigure = 1176
Public Const ResDoubleClickProps = 1178
Public Const ResReleaseButtonTo = 1180
Public Const ResCircle = 1182
Public Const ResCreateInvertedPoint = 1184
Public Const ResSelectGivens = 1186
Public Const ResSelectResults = 1188
Public Const ResEnterMacroName = 1190
Public Const ResUnableToSaveFile = 1192
Public Const ResRename = 1194
Public Const ResPromptForSave = 1196
Public Const ResBuyDG = 1198
Public Const ResMnuBase = 1200
Public Const ResMnuManageFiles = ResMnuBase + 0
Public Const ResMnuCreateNew = ResMnuBase + 2
Public Const ResMnuOpenExisting = ResMnuBase + 4
Public Const ResMnuSave = ResMnuBase + 6
Public Const ResMnuSaveAs = ResMnuBase + 8
Public Const ResMnuExit = ResMnuBase + 10
Public Const ResMnuOptions = ResMnuBase + 12
Public Const ResMnuSettings = ResMnuBase + 14
Public Const ResMnuEdit = ResMnuBase + 16
Public Const ResMnuView = ResMnuBase + 18
Public Const ResCoordinates = 1220
Public Const ResCursorCoordinates = 1222
Public Const ResToolPoints = 1224
Public Const ResToolLines = 1226
Public Const ResToolCircles = 1228
Public Const ResToolConstruction = 1230
Public Const ResToolMeasure = 1232

Public Const ResChoiceAmbiguity = 1240
Public Const ResPrivNoAlter = 1242

Public Const ResEnterExpression = 1300
Public Const ResWatchExpressions = 1302
Public Const ResAddExpression = 1304
Public Const ResEvalErrorBase = 1306
Public Const ResMacroErrBase = 1340

Public Const ResDes_LocusOfPoint = 1400
Public Const ResDes_StepAofB = 1402
Public Const ResDes_Point = 1404
Public Const ResDes_DynamicLocus = 1406
Public Const ResDes_Polygon = 1408
Public Const ResDes_Bezier = 1410
Public Const ResDes_Vector = 1412
Public Const ResDes_PointOnFigure = 1414
Public Const ResDes_Middlepoint = 1416
Public Const ResDes_SimmPoint = 1418
Public Const ResDes_SimmPointByLine = 1420
Public Const ResDes_Intersect = 1422
Public Const ResDes_Invert = 1424
Public Const ResDes_ArcEndPoint = 1426
Public Const ResDes_Segment = 1428
Public Const ResDes_Ray = 1430
Public Const ResDes_Line = 1432
Public Const ResDes_LineParallel = 1434
Public Const ResDes_LinePerpendicular = 1436
Public Const ResDes_Bisector = 1438
Public Const ResDes_Circle = 1440
Public Const ResDes_CircleByRadius = 1442
Public Const ResDes_Arc = 1444
Public Const ResDes_PointCoord = 1446

Public Const ResTipDidYouKnow = 1500
Public Const ResTipPrevious = 1502
Public Const ResTipNext = 1504
Public Const ResShowTips = 1506

'===================================================
Public Const ResUndoStringBase = 1600
'Public Const ResUndoPointPropertiesChange = 1244

Public Enum ResGenericUndoStrings
    ResUndoPointPropertiesChange = 1600
    ResUndoDeletePoint = 1602
    ResUndoDeleteFigure = 1604
    ResUndoClearAll = 1606
    ResUndoButtonPush = 1608
    ResUndoMacro = 1610
    ResUndoRenamePoint = 1612
    ResUndoCreatePolygon = 1614
    ResUndoCreateBezier = 1616
    ResUndoCreateVector = 1618
    ResUndoFigurePropertiesChange = 1620
    ResUndoSnapPoint = 1622
    ResUndoReleasePoint = 1624
End Enum

'=========================================================
'                                   HELP TOPIC MAP
'=========================================================

Public Const ResHlp_Contents = 1
Public Const ResHlp_Interface_Settings = 2
Public Const ResHlp_Interface_Expressions = 37
Public Const ResHlp_Interface_Expression = 62
Public Const ResHlp_Interface_CalcPanel = 267
Public Const ResHlp_About = 61
Public Const ResHlp_PopToolBase = 101
Public Const ResHlp_PopBase = 279

Public Sub DisplayHelpTopic(ByVal Number As Long)
If App.HelpFile <> "" Then
    CD.HelpFile = App.HelpFile
    CD.HelpCommand = HELP_CONTEXT
    CD.ShowHelp Number
Else
    MsgBox GetString(ResHelpFileNotFound) & " " & ProgramPath, vbExclamation
End If
End Sub

Public Function GetString(ByVal ID As Long, Optional ByVal LanguageID As Long = EmptyVar) As String
If LanguageID = EmptyVar Then GetString = LoadResString(ID + setLanguage) Else GetString = LoadResString(ID + LanguageID)
End Function

Public Function GetPicture(ByVal ID As Long, Optional ByVal rType As LoadResConstants = vbResBitmap, Optional ByVal lIconSize As Long = -1) As IPictureDisp
On Local Error Resume Next
If lIconSize = -1 Then lIconSize = IIf(setIconSize = 16, 200, 300)
Set GetPicture = LoadResPicture(ID + lIconSize, rType)
End Function
