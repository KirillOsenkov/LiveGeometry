Attribute VB_Name = "modSettings"
Option Explicit

'=========================================================
' Settings variables
'=========================================================

Public setLanguage As Long
Public setNoFlicker As Boolean
Public setGradientFill As Boolean
Public setToolSelectOnce As Boolean
Public Const setGroupTools As Boolean = False
Public setPaperGradient As Boolean
Public setDistancePrecision As Long
Public setAnglePrecision As Long
Public setNumberPrecision As Long
Public setFormatAngle As String
Public setFormatDistance As String
Public setFormatNumber As String
Public setMacroAutoloadPath As String
Public setLocusDetails As Long
Public setLocusDetailsHigh As Long
Public setAutoShowPointName As Boolean
Public setCursorSensitivity As Long
Public setShowMacroCreateDialog As Boolean
Public setShowMacroResultsDialog As Boolean
Public setShowTips As Boolean
Public setCurrentTip As Long

Public setcolPaper1 As Long
Public setcolPaper2 As Long
Public setcolGrid As Long
Public setcolAxes As Long
Public setcolRuler As Long
Public setcolToolbar As Long

Public setShowStatusbar As Boolean
Public setShowToolbar As Boolean
Public setShowMainbar As Boolean
Public setShowClock As Boolean
Public setShowTooltips As Boolean
Public setShowCoord As Boolean
Public setShowAxes As Boolean
Public setShowGrid As Boolean
Public setShowRulers As Boolean

Public setTransparentEMF As Boolean
Public setDynamicLocusType As Long
Public setWallpaper As String
Public setShowAxesMarks As Boolean
Public setLoadCursors As Boolean

Public setdefcolPoint As Long
Public setdefcolPointFill As Long
Public setdefcolDependentPoint As Long
Public setdefcolDependentPointFill As Long
Public setdefcolFigurePoint As Long
Public setdefcolFigurePointFill As Long
Public setdefcolPointName As Long

Public setdefPointFill As Long
Public setdefPointFontCharset As Long
Public setdefPointFontName As String
Public setdefPointFontSize As Long
Public setdefPointFontBold As Boolean
Public setdefPointFontItalic As Boolean
Public setdefPointFontUnderline As Boolean
Public setdefPointSize As Long

Public setdefcolLocus As Long
Public setdefPointShape As Long
Public setdefFigurePointShape As Long
Public setdefDependentPointShape As Long

Public setdefcolFigure As Long
Public setdefcolFigureFill As Long

Public setdefFigureDrawWidth As Long
Public setdefLocusDrawWidth As Long
Public setdefLocusType As Long
Public setdefLabelFont As String
Public setdefLabelCharset As Long
Public setdefLabelBold As Boolean
Public setdefLabelItalic As Boolean
Public setdefLabelUnderline As Boolean
Public setdefLabelTransparent As Boolean
Public setdefcolLabelColor As Long
Public setdefcolLabelBackColor As Long
Public setdefLabelFontSize As Long
Public setIconSize As Long
Public setDisplayIconText As Long

'==================================================

Public Function GetBoolSetting(AppName As String, Section As String, Key As String, Default As Boolean) As Boolean
Dim S As String
On Local Error Resume Next

S = GetSetting(AppName, Section, Key, Default)
If S = "" Then
    GetBoolSetting = Default
Else
    If S = "0" Or S = "False" Then
        GetBoolSetting = False
    ElseIf S = "1" Or S = "True" Then
        GetBoolSetting = True
    Else
        GetBoolSetting = CBool(S)
    End If
End If
End Function

Public Sub GetSettings()
On Local Error Resume Next

Dim NewAlign As AlignConstants, Z As Long

If GetSetting(AppName, "General", "Language") = "" Then
    setLanguage = conLanguage
    SaveSetting AppName, "General", "Language", setLanguage
Else
    setLanguage = Val(GetSetting(AppName, "General", "Language"))
End If

If GetSetting(AppName, "General", "AppPath", "") = "" Then
    SaveSetting AppName, "General", "AppPath", AddDirSep(App.Path)
ElseIf Dir(GetSetting(AppName, "General", "AppPath", ""), 23) = "" Then
    SaveSetting AppName, "General", "AppPath", AddDirSep(App.Path)
End If

setIconSize = Val(GetSetting(AppName, "Interface", "IconSize", IconSize))
setDisplayIconText = Val(GetSetting(AppName, "Interface", "DisplayIconText", GraphicsOnly))
setGradientFill = GetBoolSetting(AppName, "Interface", "GradientFill", IIf(ShouldGradientFill, "1", "0"))
setToolSelectOnce = GetBoolSetting(AppName, "General", "ToolSelectOnce", "0")
setDistancePrecision = Val(GetSetting(AppName, "General", "DistancePrecision", "1"))
setNumberPrecision = Val(GetSetting(AppName, "General", "NumberPrecision", "1"))
setAnglePrecision = Val(GetSetting(AppName, "General", "AnglePrecision", "0"))
setLocusDetails = Val(GetSetting(AppName, "General", "LocusDetails", DynamicLocusPointCount))
setLocusDetailsHigh = Val(GetSetting(AppName, "General", "LocusDetailsHigh", HighQualityDynamicLocusPointCount))
setMacroAutoloadPath = GetSetting(AppName, "General", "MacroAutoloadPath", "")
If Dir(setMacroAutoloadPath, 23) = "" Then setMacroAutoloadPath = ""
If (GetFileAttributes(setMacroAutoloadPath) And vbDirectory) <> vbDirectory Then setMacroAutoloadPath = ""

setPaperGradient = GetBoolSetting(AppName, "Interface", "PaperGradientFill", 0)
setcolPaper1 = GetNearestColor(Paper.hDC, Val(GetSetting(AppName, "Interface", "PaperColor1", colPaperGradient1)))
setcolPaper2 = Val(GetSetting(AppName, "Interface", "PaperColor2", colPaperGradient2))
setcolGrid = Val(GetSetting(AppName, "Interface", "GridColor", colGridColor))
setcolAxes = Val(GetSetting(AppName, "Interface", "AxesColor", colAxes))
setShowAxes = GetBoolSetting(AppName, "Interface", "ShowAxes", "0")
setShowGrid = GetBoolSetting(AppName, "Interface", "ShowGrid", "0")

setLoadCursors = GetBoolSetting(AppName, "Interface", "LoadCursors", "1")
setShowAxesMarks = GetBoolSetting(AppName, "Interface", "ShowAxesMarks", "1")
setAutoShowPointName = GetBoolSetting(AppName, "General", "AutoShowPointName", "1")
setShowMacroCreateDialog = GetBoolSetting(AppName, "General", "ShowMacroCreateDialog", "1")
setShowMacroResultsDialog = GetBoolSetting(AppName, "General", "ShowMacroResultsDialog", "1")
setShowTips = GetBoolSetting(AppName, "General", "ShowTips", "0")
setCurrentTip = Val(GetSetting(AppName, "General", "CurrentTip", 1))

setTransparentEMF = GetBoolSetting(AppName, "General", "TransparentEMF", "0")
setWallpaper = GetSetting(AppName, "Interface", "Wallpaper", "")
If setWallpaper <> "" Then
    If Dir(setWallpaper) = "" Then setWallpaper = ""
End If

setShowRulers = GetBoolSetting(AppName, "Interface", "ShowRulers", "0")
setcolRuler = Val(GetSetting(AppName, "Interface", "RulerColor", colRulerBackColor))
setcolToolbar = Val(GetSetting(AppName, "Interface", "ToolbarColor", EnsureRGB(vbButtonFace)))

nPaperColor1 = setcolPaper1
nPaperColor2 = setcolPaper2
nGridColor = setcolGrid
nAxesColor = setcolAxes
nGradientPaper = setPaperGradient
nShowAxes = setShowAxes
nShowGrid = setShowGrid

setCursorSensitivity = Val(GetSetting(AppName, "General", "CursorSensitivity", "4"))
setShowStatusbar = GetBoolSetting(AppName, "Interface", "ShowStatusbar", True)
setShowToolbar = GetBoolSetting(AppName, "Interface", "ShowToolbar", True)
setShowMainbar = GetBoolSetting(AppName, "Interface", "ShowMainbar", True)
setShowClock = GetBoolSetting(AppName, "Interface", "ShowClock", False)
setShowCoord = GetBoolSetting(AppName, "Interface", "ShowCoord", False)
setShowTooltips = GetBoolSetting(AppName, "Interface", "ShowTooltips", True)

'==============================================
' Default object colors
'==============================================

setdefcolPoint = Val(GetSetting(AppName, "Defaults", "PointColor", colPointColor))
setdefcolPointFill = Val(GetSetting(AppName, "Defaults", "PointFillColor", colPointFillColor))
setdefcolFigurePoint = Val(GetSetting(AppName, "Defaults", "SemiDependent", colSemiDependentForeColor))
setdefcolFigurePointFill = Val(GetSetting(AppName, "Defaults", "FigurePointFillColor", colFigurePointFillColor))
setdefcolDependentPoint = Val(GetSetting(AppName, "Defaults", "Dependent", colDependent))
setdefcolDependentPointFill = Val(GetSetting(AppName, "Defaults", "DependentPointFillColor", colDependentPointFillColor))

setdefcolFigure = Val(GetSetting(AppName, "Defaults", "FigureColor", colFigureForeColor))
setdefcolFigureFill = Val(GetSetting(AppName, "Defaults", "FigureFillColor", colFigureFillColor))
setdefcolLocus = Val(GetSetting(AppName, "Defaults", "LocusColor", colLocusColor))
setdefLocusType = Val(GetSetting(AppName, "Defaults", "LocusType", 0))
setdefcolLabelColor = Val(GetSetting(AppName, "Defaults", "LabelColor", colTextColor))
setdefcolLabelBackColor = Val(GetSetting(AppName, "Defaults", "LabelBackColor", colPaperGradient1))

setdefPointFill = Val(GetSetting(AppName, "Defaults", "PointFill", 0))
setdefcolPointName = Val(GetSetting(AppName, "Defaults", "PointNameColor", colTextColor))
setdefPointFontName = GetSetting(AppName, "Defaults", "PointFontName", defSLabelFontName)
setdefPointFontCharset = GetSetting(AppName, "Defaults", "PointFontCharset", DefaultFontCharset)
setdefPointFontSize = GetSetting(AppName, "Defaults", "PointFontSize", defSLabelFontSize)
setdefPointFontBold = GetSetting(AppName, "Defaults", "PointFontBold", False)
setdefPointFontItalic = GetSetting(AppName, "Defaults", "PointFontItalic", False)
setdefPointFontUnderline = GetSetting(AppName, "Defaults", "PointFontUnderline", False)

setdefPointShape = Val(GetSetting(AppName, "Defaults", "PointShape", defPointShape))
setdefFigurePointShape = Val(GetSetting(AppName, "Defaults", "FigurePointShape", defFigurePointShape))
setdefDependentPointShape = Val(GetSetting(AppName, "Defaults", "DependentPointShape", defDependentPointShape))

setdefPointSize = Val(GetSetting(AppName, "Defaults", "PointSize", defPointSize))

setdefFigureDrawWidth = Val(GetSetting(AppName, "Defaults", "FigureDrawWidth", defFigureDrawWidth))
setdefLocusDrawWidth = Val(GetSetting(AppName, "Defaults", "LocusDrawWidth", defFigureDrawWidth))
setdefLabelFont = GetSetting(AppName, "Defaults", "LabelFont", defLabelFontName)
setdefLabelCharset = GetSetting(AppName, "Defaults", "LabelCharset", DefaultFontCharset)
setdefLabelBold = GetBoolSetting(AppName, "Defaults", "LabelFontBold", defLabelFontBold)
setdefLabelItalic = GetBoolSetting(AppName, "Defaults", "LabelFontItalic", defLabelFontitalic)
setdefLabelUnderline = GetBoolSetting(AppName, "Defaults", "LabelFontUnderline", defLabelFontUnderline)
setdefLabelTransparent = True 'CBool(GetSetting(AppName, "Defaults", "LabelFontTransparent", True))
setdefLabelFontSize = Val(GetSetting(AppName, "Defaults", "LabelFontSize", defLabelFontSize))

FillDefColorArray ColorArray
For Z = 0 To 15
    ColorArray(Z) = Val(GetSetting(AppName, "Custom colors", "Color" & Format(Z), Format(ColorArray(Z))))
Next

setNoFlicker = True
'Paper.AutoRedraw = setNoFlicker
End Sub

Public Function BoolToString(B As Boolean) As String
If B Then BoolToString = "1" Else BoolToString = "0"
End Function

Public Sub SaveSettings()
Dim Z As Long

For Z = 0 To 15
    SaveSetting AppName, "Custom colors", "Color" & Z, Format(ColorArray(Z))
Next

SaveSetting AppName, "Interface", "GradientFill", BoolToString(setGradientFill)
SaveSetting AppName, "Interface", "IconSize", setIconSize
SaveSetting AppName, "Interface", "DisplayIconText", setDisplayIconText
SaveSetting AppName, "General", "ToolSelectOnce", BoolToString(setToolSelectOnce)
SaveSetting AppName, "General", "ShowMacroCreateDialog", BoolToString(setShowMacroCreateDialog)
SaveSetting AppName, "General", "ShowMacroResultsDialog", BoolToString(setShowMacroResultsDialog)
SaveSetting AppName, "General", "ShowTips", BoolToString(setShowTips)

SaveSetting AppName, "Interface", "LoadCursors", BoolToString(setLoadCursors)
'SaveSetting AppName, "General", "GroupTools", Format(-CInt(setGroupTools))
SaveSetting AppName, "General", "DistancePrecision", setDistancePrecision
SaveSetting AppName, "General", "NumberPrecision", setNumberPrecision
SaveSetting AppName, "General", "AnglePrecision", setAnglePrecision
SaveSetting AppName, "General", "LocusDetails", setLocusDetails
SaveSetting AppName, "General", "LocusDetailsHigh", setLocusDetailsHigh
SaveSetting AppName, "General", "MacroAutoloadPath", setMacroAutoloadPath

SaveSetting AppName, "Interface", "ShowAxesMarks", BoolToString(setShowAxesMarks)
SaveSetting AppName, "Interface", "ShowRulers", BoolToString(setShowRulers)
SaveSetting AppName, "General", "AutoShowPointName", BoolToString(setAutoShowPointName)
SaveSetting AppName, "General", "CursorSensitivity", Format(setCursorSensitivity)
SaveSetting AppName, "General", "TransparentEMF", BoolToString(setTransparentEMF)

SaveSetting AppName, "Interface", "PaperGradientFill", BoolToString(setPaperGradient)
SaveSetting AppName, "Interface", "ShowAxes", BoolToString(setShowAxes)
SaveSetting AppName, "Interface", "ShowGrid", BoolToString(setShowGrid)

SaveSetting AppName, "Interface", "PaperColor1", setcolPaper1
SaveSetting AppName, "Interface", "PaperColor2", setcolPaper2
SaveSetting AppName, "Interface", "GridColor", setcolGrid
SaveSetting AppName, "Interface", "AxesColor", setcolAxes

SaveSetting AppName, "Interface", "RulerColor", setcolRuler
SaveSetting AppName, "Interface", "ToolbarColor", setcolToolbar
SaveSetting AppName, "Interface", "ShowStatusbar", BoolToString(setShowStatusbar)
SaveSetting AppName, "Interface", "ShowToolbar", BoolToString(setShowToolbar)
SaveSetting AppName, "Interface", "ShowMainbar", BoolToString(setShowMainbar)
SaveSetting AppName, "Interface", "ShowClock", BoolToString(setShowClock)
SaveSetting AppName, "Interface", "ShowTooltips", BoolToString(setShowTooltips)
SaveSetting AppName, "Interface", "ShowCoord", BoolToString(setShowCoord)
SaveSetting AppName, "Interface", "Wallpaper", setWallpaper

SaveSetting AppName, "Defaults", "PointFill", setdefPointFill
SaveSetting AppName, "Defaults", "PointNameColor", setdefcolPointName
SaveSetting AppName, "Defaults", "PointFontName", setdefPointFontName
SaveSetting AppName, "Defaults", "PointFontCharset", setdefPointFontCharset
SaveSetting AppName, "Defaults", "PointFontSize", setdefPointFontSize
SaveSetting AppName, "Defaults", "PointFontBold", BoolToString(setdefPointFontBold)
SaveSetting AppName, "Defaults", "PointFontItalic", BoolToString(setdefPointFontItalic)
SaveSetting AppName, "Defaults", "PointFontUnderline", BoolToString(setdefPointFontUnderline)

SaveSetting AppName, "Defaults", "PointShape", setdefPointShape
SaveSetting AppName, "Defaults", "FigurePointShape", setdefFigurePointShape
SaveSetting AppName, "Defaults", "DependentPointShape", setdefDependentPointShape

SaveSetting AppName, "Defaults", "PointSize", setdefPointSize

'==============================================
' Default object colors
'==============================================

SaveSetting AppName, "Defaults", "PointColor", setdefcolPoint
SaveSetting AppName, "Defaults", "PointFillColor", setdefcolPointFill
SaveSetting AppName, "Defaults", "SemiDependent", setdefcolFigurePoint
SaveSetting AppName, "Defaults", "FigurePointFillColor", setdefcolFigurePointFill
SaveSetting AppName, "Defaults", "Dependent", setdefcolDependentPoint
SaveSetting AppName, "Defaults", "DependentPointFillColor", setdefcolDependentPointFill

SaveSetting AppName, "Defaults", "FigureColor", setdefcolFigure
SaveSetting AppName, "Defaults", "FigureFillColor", setdefcolFigureFill
SaveSetting AppName, "Defaults", "LocusColor", setdefcolLocus
SaveSetting AppName, "Defaults", "LocusType", setdefLocusType
SaveSetting AppName, "Defaults", "LabelColor", setdefcolLabelColor
SaveSetting AppName, "Defaults", "LabelBackColor", setdefcolLabelBackColor


SaveSetting AppName, "Defaults", "FigureDrawWidth", setdefFigureDrawWidth
SaveSetting AppName, "Defaults", "LocusDrawWidth", setdefLocusDrawWidth
SaveSetting AppName, "Defaults", "LabelFont", setdefLabelFont
SaveSetting AppName, "Defaults", "LabelCharset", setdefLabelCharset
SaveSetting AppName, "Defaults", "LabelFontBold", BoolToString(setdefLabelBold)
SaveSetting AppName, "Defaults", "LabelFontItalic", BoolToString(setdefLabelItalic)
SaveSetting AppName, "Defaults", "LabelFontUnderline", BoolToString(setdefLabelUnderline)
SaveSetting AppName, "Defaults", "LabelFontTransparent", BoolToString(setdefLabelTransparent)
SaveSetting AppName, "Defaults", "LabelFontSize", setdefLabelFontSize

With FormMain
    For Z = .MenuBar.LBound To .MenuBar.UBound
        SaveSetting AppName, "General", "MenuAlign" & Z, Format(.MenuBar(Z).Align)
    Next
End With

SaveSetting AppName, "General", "Language", Format(setLanguage)
End Sub

Public Sub GetDefaults()
Paper.ScaleMode = ScaleModeConstants.vbPixels
FormMain.XRuler.ForeColor = colRulerForeColor
FormMain.YRuler.ForeColor = colRulerForeColor
FormMain.BackColor = colFormBackGround
FormMain.StatusBar.AutoRedraw = StatusBarAutoRedraw
setFormatAngle = GetFormatString(setAnglePrecision)
setFormatDistance = GetFormatString(setDistancePrecision)
setFormatNumber = GetFormatString(setNumberPrecision)
FromRedo = False
Set curArrow = LoadResPicture(101, vbResCursor)
Set curArrowPlus = LoadResPicture(102, vbResCursor)
Set curArrowMinus = LoadResPicture(103, vbResCursor)
Set curArrowCross = LoadResPicture(104, vbResCursor)
Set curArrowDrag = LoadResPicture(105, vbResCursor)
Set curArrowSelect = LoadResPicture(106, vbResCursor)
setPaperCursorArrow = 99
setPaperCursorCross = 99
setPaperCursorDrag = 99
End Sub

Public Function ShouldGradientFill() As Boolean
If ColorResolution > 8 Then ShouldGradientFill = True
End Function

Public Function ColorResolution() As Long
ColorResolution = GetDeviceCaps(GetWindowDC(GetDesktopWindow), BITSPIXEL)
End Function

Public Sub SaveLastPaths()
SaveSetting AppName, "Paths", "LastFigurePath", LastFigurePath
SaveSetting AppName, "Paths", "LastMacroPath", LastMacroPath
SaveSetting AppName, "Paths", "LastHTMLPath", LastHTMLPath
SaveSetting AppName, "Paths", "LastPicturePath", LastPicturePath
SaveSetting AppName, "Paths", "LastFileLinkPath", LastFileLinkPath
SaveSetting AppName, "Paths", "LastSoundPath", LastSoundPath
SaveSetting AppName, "Paths", "LastTexturePath", LastTexturePath
End Sub

Public Sub LoadLastPaths()
LastFigurePath = GetSetting(AppName, "Paths", "LastFigurePath", ProgramPath)
LastMacroPath = GetSetting(AppName, "Paths", "LastMacroPath", ProgramPath)
LastHTMLPath = GetSetting(AppName, "Paths", "LastHTMLPath", ProgramPath)
LastPicturePath = GetSetting(AppName, "Paths", "LastPicturePath", ProgramPath)
LastFileLinkPath = GetSetting(AppName, "Paths", "LastFileLinkPath", ProgramPath)
LastSoundPath = GetSetting(AppName, "Paths", "LastSoundPath", ProgramPath)
LastTexturePath = GetSetting(AppName, "Paths", "LastTexturePath", ProgramPath)
If Not IsValidPath(LastFigurePath) Then LastFigurePath = ProgramPath
If Not IsValidPath(LastMacroPath) Then LastMacroPath = ProgramPath
If Not IsValidPath(LastHTMLPath) Then LastHTMLPath = ProgramPath
If Not IsValidPath(LastPicturePath) Then LastPicturePath = ProgramPath
If Not IsValidPath(LastFileLinkPath) Then LastFileLinkPath = ProgramPath
If Not IsValidPath(LastSoundPath) Then LastSoundPath = ProgramPath
If Not IsValidPath(LastTexturePath) Then LastTexturePath = ProgramPath
End Sub

Public Sub FillDefColorArray(ColArray() As Long)
'ColArray(0) = &HFFFAFA
'ColArray(1) = &HF8F8FF
'ColArray(2) = &HFFFBF0
'ColArray(3) = &HFFDAB9
'ColArray(4) = &HFFFACD
'ColArray(5) = &HF0FFFF
'ColArray(6) = &HE6E6FA
'ColArray(7) = &HFFF0F5
'ColArray(8) = &HFFE4E1
'ColArray(9) = &H708090
'ColArray(10) = &H191970
'ColArray(11) = &H80
'ColArray(12) = &H41690
'ColArray(13) = &HAFEEEE
'ColArray(14) = &HFF7F
'ColArray(15) = &HEEE8AA
ColArray(0) = 10658555
ColArray(1) = 10862044
ColArray(2) = 10537677
ColArray(3) = 12181950
ColArray(4) = 13881771
ColArray(5) = 15254721
ColArray(6) = 14728415
ColArray(7) = 12961221
ColArray(8) = 8684793
ColArray(9) = 9684729
ColArray(10) = 10742513
ColArray(11) = 11467197
ColArray(12) = 15329707
ColArray(13) = 15903413
ColArray(14) = 16300270
ColArray(15) = 14211288
End Sub

