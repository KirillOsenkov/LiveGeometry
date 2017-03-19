Attribute VB_Name = "modFileIO"
Option Explicit

Public Function OpenFile(ByVal FName As String, Optional ByVal ShowProgress As Boolean = True, Optional ByVal ProgressQuiet As Boolean = False, Optional ByVal ShouldChangeCaption As Boolean = True) As Boolean
On Error GoTo EH:

Const ProgressStep = 100
Dim ErrStr As String, TotalProgress As Long, CurrentProgress As Long, ReadLength As Long, TotalLength As Long, ProgressStepCount As Long
Dim EqPos As Long, CurObj As Long, CurSect As String, CurSectType As Long
Dim AName As String, AValue As String, CurLine As Long
Dim BuildVersion As Long
Dim curChildren As Long, curParents As Long, curPoints As Long, curLetters As Long, curAuxInfo As Long, curAuxPointsX As Long, curAuxPointsY As Long
Dim curButtons As Long, curFigures As Long, curLabels As Long, curLoci As Long, curSG As Long, curWE As Long
Dim tWEName As String, tWEExpression As String
Dim tX As Double, tY As Double
Dim tHide As Boolean
Dim Z As Long, A As String, S As String, Q As Long
Dim RemChar As String

If Dir$(FName) = "" Then ErrStr = "Cannot find file": GoTo EH

ClearAll , Not ProgressQuiet, False

DrawingName = FName
If ShouldChangeCaption Then FormMain.Caption = GetString(ResCaption) & " - " & RetrieveName(FName)

If ShowProgress Then
    TotalLength = FileLen(FName)
    ReadLength = 0
    ProgressShow GetString(ResLoading) & " " & RetrieveName(FName) & "...", ProgressQuiet
End If

BuildVersion = App.Revision
CurSect = "General"
CurSectType = 0
CurLine = 0
ProgressStepCount = 0

Open FName For Input As #1
    Do While True
        Do
            If EOF(1) Then A = "E": Exit Do
            Line Input #1, A
            CurLine = CurLine + 1
            ProgressStepCount = ProgressStepCount + 1
            If ProgressStepCount = ProgressStep Then ProgressStepCount = 0
            
            If ShowProgress Then
                ReadLength = ReadLength + Len(A) + 2
                If ProgressStepCount = 0 Then ProgressUpdate ReadLength / TotalLength
            End If
            
            If Left(A, 1) = "[" And Right(A, 1) = "]" Then
                If CurSectType = 2 Then
                    If Not tHide Then
                        Figures(CurObj).Hide = Not Figures(CurObj).Visible
                    Else
                        tHide = False
                    End If
                End If
                
                CurSect = GetSectionHeader(A): A = ";" & A
                
                If CurSect = "General" Then
                    CurSectType = 0
                    CurObj = 0
                
                ElseIf Left(CurSect, 5) = "Point" Then
                    CurSectType = 1
                    CurObj = Val(Right(CurSect, Len(CurSect) - 5))
                    If Not IsPoint(CurObj) Then CurSectType = -1 Else FillPointWithDefaults BasePoint(CurObj)
                    '
                ElseIf Left(CurSect, 6) = "Figure" Then
                    CurSectType = 2
                    CurObj = Val(Right(CurSect, Len(CurSect) - 6))
                    If Not IsFigure(CurObj) Then
                        CurSectType = -1
                    Else
                        FillFigureWithDefaults Figures(CurObj)
                        curAuxInfo = 0
                        curAuxPointsX = 0
                        curAuxPointsY = 0
                        curChildren = 0
                        curLetters = 0
                        curParents = 0
                        tHide = False
                    End If
                    '
                ElseIf Left(CurSect, 5) = "Label" Then
                    CurSectType = 3
                    CurObj = Val(Right(CurSect, Len(CurSect) - 5))
                    If Not IsLabel(CurObj) Then
                        CurSectType = -1
                    Else
                        TextLabels(CurObj).Transparent = True
                        TextLabels(CurObj).Visible = True
                        TextLabels(CurObj).Borders = tbsNone
                        TextLabels(CurObj).FontBold = False
                        TextLabels(CurObj).FontItalic = False
                        TextLabels(CurObj).FontUnderline = False
                        TextLabels(CurObj).BackColor = setdefcolLabelBackColor
                        TextLabels(CurObj).InDemo = True
                    End If
                    '
                ElseIf Left(CurSect, 15) = "WatchExpression" Then
                    CurSectType = 4
                    CurObj = Val(Right(CurSect, Len(CurSect) - 15))
'                    If CurObj = 1 Then
'                        For Z = 0 To FigureCount - 1
'                            If Figures(Z).FigureType = dsAnPoint Then
'                                Figures(Z).XTree = BuildTree(Figures(Z).XS)
'                                If Figures(Z).XTree.Erroneous Then
'                                    Figures(Z).XTree = BuildTree("0")
'                                End If
'                                Figures(Z).YTree = BuildTree(Figures(Z).YS)
'                                If Figures(Z).YTree.Erroneous Then
'                                    Figures(Z).YTree = BuildTree("0")
'                                End If
'                            End If
'                            RecalcAuxInfo Z   '?????
'                        Next
'                    End If
                    '
                ElseIf Left(CurSect, 5) = "Locus" Then
                    CurSectType = 5
                    CurObj = Val(Right(CurSect, Len(CurSect) - 5))
                    If CurObj > LocusCount Then
                        CurSectType = -1
                    Else
                        Locuses(CurObj).LocusNumber = 1
                        ReDim Locuses(CurObj).LocusNumbers(1 To 1)
                    End If
                    
                    '
                ElseIf Left(CurSect, 2) = "SG" Then
                    CurSectType = 6
                    CurObj = Val(Right(CurSect, Len(CurSect) - 2))
                    If Not IsSG(CurObj) Then
                        CurSectType = -1
                    Else
                        StaticGraphics(CurObj).InDemo = True
                    End If
                    '
                ElseIf Left(CurSect, 6) = "Button" Then
                    CurSectType = 7
                    CurObj = Val(Right(CurSect, Len(CurSect) - 6))
                    If Not IsButton(CurObj) Then
                        CurSectType = -1
                    Else
                        With Buttons(CurObj)
                            .Caption = ""
                            .Charset = setdefLabelCharset
                            .BackColor = GetSysColor(vbButtonFace + SysColorTranslationBase)
                            .ForeColor = GetSysColor(vbButtonText + SysColorTranslationBase)
                            .Visible = True
                            .InDemo = True
                            .InitiallyVisible = False
                            .Message = ""
                            .FontName = setdefLabelFont
                            .FontBold = setdefLabelBold
                            .FontItalic = setdefLabelItalic
                            .FontSize = setdefLabelFontSize
                            .FontUnderline = setdefLabelUnderline
                            .LogicalPosition.P1.X = 0
                            .LogicalPosition.P1.Y = 0
                        End With
                    End If
                    '
                ElseIf CurSect = "Misc" Then
                    CurSectType = 8
                    CurObj = 0
                End If
            End If
            
            EqPos = InStr(A, "=")
            If A <> "" Then RemChar = Left(A, 1) Else RemChar = ""
        Loop Until RemChar <> "" And RemChar <> ";" And EqPos <> 0
        
        If A = "E" Then Exit Do
        
        AName = Left(A, EqPos - 1)
        AValue = Right(A, Len(A) - EqPos)
        
        Select Case CurSectType
            Case 0 ' [General]
                Select Case AName
                    Case "FileFormatVersion"
                        BuildVersion = Val(Right(AValue, Len(AValue) - InStrRev(AValue, ".")))
                    Case "PointCount"
                        PointCount = Val(AValue)
                        If PointCount > 0 Then RedimBasePoint 1, PointCount
                    Case "FigureCount"
                        FigureCount = Val(AValue)
                        If FigureCount > 0 Then RedimFigures 0, FigureCount - 1
                    Case "LabelCount"
                        LabelCount = Val(AValue)
                        If LabelCount > 0 Then ReDim TextLabels(1 To LabelCount) Else ReDim TextLabels(1 To 1)
                    Case "WatchExpressionCount"
                        'WECount = Val(AValue)
                        'If WECount > 0 Then ReDim WatchExpressions(1 To WECount)
                    Case "StaticGraphicCount"
                        StaticGraphicCount = Val(AValue)
                        If StaticGraphicCount > 0 Then ReDim StaticGraphics(1 To StaticGraphicCount)
                    Case "ButtonCount"
                        ButtonCount = Val(AValue)
                        If ButtonCount > 0 Then ReDim Buttons(1 To ButtonCount)
                    Case "LocusCount"
                        LocusCount = Val(AValue)
                        If LocusCount > 0 Then ReDim Locuses(1 To LocusCount)
                    Case "PaperColor1"
                        nPaperColor1 = Val(AValue)
                        Paper.BackColor = nPaperColor1
                    Case "PaperColor2"
                        nPaperColor2 = Val(AValue)
                    Case "AxesColor"
                        nAxesColor = Val(AValue)
                    Case "GridColor"
                        nGridColor = Val(AValue)
                    Case "GradientPaper"
                        nGradientPaper = FromBoolean(AValue)
                    Case "ShowAxes"
                        nShowAxes = FromBoolean(AValue)
                        FormMain.mnuShowAxes.Checked = nShowAxes
                    Case "ShowGrid"
                        nShowGrid = FromBoolean(AValue)
                        FormMain.mnuShowGrid.Checked = nShowGrid
                    Case "Description"
                        nDescription = ToMultiLine(AValue)
                    Case "WorkArea"
                        WorkArea.P1.X = Val(GetParameter(AValue, 1, True))
                        WorkArea.P1.Y = Val(GetParameter(AValue, 2, True))
                        WorkArea.P2.X = Val(GetParameter(AValue, 3, True))
                        WorkArea.P2.Y = Val(GetParameter(AValue, 4, True))
                        If GetParameter(AValue, 5, True) <> "" Then
                            WorkArea.P1.X = 0
                            WorkArea.P1.Y = 0
                            WorkArea.P2.X = 0
                            WorkArea.P2.Y = 0
                        End If
                    Case "DemoInterval"
                        nDemoInterval = Val(AValue)
                        If nDemoInterval < 500 Or nDemoInterval > 15000 Then nDemoInterval = defDemoInterval
                End Select
            Case 1 ' [Point]
                With BasePoint(CurObj)
                    Select Case AName
                        Case "Enabled"
                            .Enabled = FromBoolean(AValue)
                        Case "FillColor"
                            .FillColor = Val(AValue)
                        Case "FillStyle"
                            .FillStyle = Val(AValue)
                        Case "ForeColor"
                            .ForeColor = Val(AValue)
                        Case "Hide"
                            .Hide = FromBoolean(AValue)
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
                            'PointNames.Add AValue
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
                            MovePoint CurObj, .X, .Y
                        Case "ZOrder"
                            .ZOrder = Val(AValue)
                    End Select
                End With
            Case 2 ' [Figure]
                With Figures(CurObj)
                    If Left(AName, 8) = "Children" Then
                        curChildren = Right(AName, Len(AName) - 8)
                        If curChildren < .NumberOfChildren Then .Children(curChildren) = Val(AValue)
                    ElseIf Left(AName, 7) = "Parents" Then
                        curParents = Right(AName, Len(AName) - 7)
                        If curParents < GetProperParentNumber(.FigureType) Then .Parents(curParents) = Val(AValue)
                    ElseIf Left(AName, 6) = "Points" Then
                        curPoints = Right(AName, Len(AName) - 6)
                        If curPoints < .NumberOfPoints Then .Points(curPoints) = Val(AValue)
'                    ElseIf Left(AName, 7) = "Letters" Then
'                        curLetters = Right(AName, Len(AName) - 7)
'                        If curLetters < .NumberOfPoints Then .Letters(curLetters) = Val(AValue)
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
                            'FigureNames.Add AValue
                        Case "FigureType"
                            .FigureType = Val(AValue)
                            If GetProperParentNumber(.FigureType) > 0 Then ReDim .Parents(0 To GetProperParentNumber(.FigureType) - 1)
                            If .FigureType = dsMeasureAngle Then
                                .AuxInfo(2) = defAngleMarkRadius
                            End If
                        Case "FigureTypeString"
                            'do nothing
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
                        Case "YS"
                            .YS = AValue
                        Case "Hide"
                            .Hide = FromBoolean(AValue)
                            tHide = True
                        Case "InDemo"
                            .InDemo = FromBoolean(AValue)
                        Case "Desc"
                            .Description = AValue
                        Case "ZOrder"
                            .ZOrder = Val(AValue)
                    End Select
                End With
            Case 3 'LABEL
                With TextLabels(CurObj)
                    Select Case AName
                        Case "Borders"
                            .Borders = Val(AValue)
                        Case "Charset"
                            .Charset = Val(AValue)
                        Case "Caption"
                            .Caption = ToMultiLine(AValue)
                            .DisplayName = .Caption
                            .LenDisplayName = Len(.Caption)
                            .Charset = PickCharset(.Caption)
                        Case "FontName"
                            .FontName = Trim(AValue)
                            If Not FontExists(.FontName) Then .FontName = "Arial"
                        Case "FontSize"
                            .FontSize = Val(AValue)
                        Case "FontBold"
                            .FontBold = FromBoolean(AValue)
                        Case "FontItalic"
                            .FontItalic = FromBoolean(AValue)
                        Case "FontUnderline"
                            .FontUnderline = FromBoolean(AValue)
                        Case "ForeColor"
                            .ForeColor = Val(AValue)
                        Case "BackColor"
                            .BackColor = Val(AValue)
                        Case "Transparent"
                            .Transparent = FromBoolean(AValue)
                        Case "Hide"
                            .Hide = FromBoolean(AValue)
                        Case "InDemo"
                            .InDemo = FromBoolean(AValue)
                        Case "Desc"
                            .Description = AValue
                        'Case "Visible"
                        '    .Visible = FromBoolean(AValue)
                        Case "Fixed"
                            .Fixed = FromBoolean(AValue)
                        Case "X"
                            .LogicalPosition.P1.X = Val(AValue)
                        Case "Y"
                            .LogicalPosition.P1.Y = Val(AValue)
                    End Select
                End With
            Case 4 'WE
                Select Case AName
                    Case "Name"
                        tWEName = AValue
                    Case "Expression"
                        tWEExpression = AValue
                        'FormMain.ValueTable1.AddExpression tWEName, tWEExpression, False
                End Select
            Case 5 'LOCUS
                With Locuses(CurObj)
                    If Left(AName, 5) = "Point" And Right(AName, 1) = "X" Then
                        curPoints = Mid(AName, 6, Len(AName) - 7)
                        tX = Val(AValue)
                    End If
                    If Left(AName, 5) = "Point" And Right(AName, 1) = "Y" Then
                        curPoints = Mid(AName, 6, Len(AName) - 7)
                        If curPoints <= .LocusPointCount Then
                            tY = Val(AValue)
                            .LocusPoints(curPoints).X = tX
                            .LocusPoints(curPoints).Y = tY
                            ToPhysical tX, tY
                            .LocusPixels(curPoints).X = tX
                            .LocusPixels(curPoints).Y = tY
                        End If
                    End If
                    
                    Select Case AName
                        Case "DrawWidth"
                            .DrawWidth = Val(AValue)
                        Case "Dynamic"
                            .Dynamic = FromBoolean(AValue)
                            If .Dynamic Then
                                
                            End If
                        Case "Enabled"
                            .Enabled = FromBoolean(AValue)
                        Case "ForeColor"
                            .ForeColor = Val(AValue)
                        Case "LocusPointCount"
                            .LocusPointCount = Val(AValue)
                            If .LocusPointCount > 0 Then
                                ReDim .LocusPoints(1 To .LocusPointCount)
                                ReDim .LocusPixels(1 To .LocusPointCount)
                            Else
                                ReDim .LocusPoints(1 To 1)
                                ReDim .LocusPixels(1 To 1)
                            End If
                            .LocusNumbers(1) = .LocusPointCount
                        Case "ParentPoint" 'Important: ParentPoint MUST Precede ParentFigure
                            .ParentPoint = Val(AValue)
                            .ParentFigure = GetLocusParentFigure(.ParentPoint)
                        Case "ParentFigure"
                            .ParentFigure = Val(AValue)
                        Case "Type"
                            .Type = Val(AValue)
                        Case "Visible"
                            .Visible = FromBoolean(AValue)
                        Case "Hide"
                            .Hide = FromBoolean(AValue)
                        Case "Pieces"
                            Dim Sa() As String
                            Sa = Split(AValue, ";")
                            .LocusNumber = UBound(Sa) + 1
                            ReDim .LocusNumbers(1 To .LocusNumber)
                            For Q = 1 To .LocusNumber
                                .LocusNumbers(Q) = Sa(Q - 1)
                            Next
                    End Select
                End With
            Case 6 'SG
                With StaticGraphics(CurObj)
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
                        Case "Hide"
                            .Hide = FromBoolean(AValue)
                        Case "InDemo"
                            .InDemo = FromBoolean(AValue)
                        Case "Desc"
                            .Description = AValue
                    End Select
                End With
            
            Case 7 'BUTTON
                With Buttons(CurObj)
                    If Left(AName, 7) = "Button#" Then
                        curButtons = Val(Right(AName, Len(AName) - 8))
                        ObjectListAdd .ObjectListAux, gotButton, Val(AValue)
                    ElseIf Left(AName, 7) = "Figure#" Then
                        curFigures = Val(Right(AName, Len(AName) - 8))
                        ObjectListAdd .ObjectListAux, gotFigure, Val(AValue)
                    ElseIf Left(AName, 6) = "Point#" Then
                        curPoints = Val(Right(AName, Len(AName) - 7))
                        ObjectListAdd .ObjectListAux, gotPoint, Val(AValue)
                    ElseIf Left(AName, 6) = "Label#" Then
                        curLabels = Val(Right(AName, Len(AName) - 7))
                        ObjectListAdd .ObjectListAux, gotLabel, Val(AValue)
                    ElseIf Left(AName, 6) = "Locus#" Then
                        curLoci = Val(Right(AName, Len(AName) - 7))
                        ObjectListAdd .ObjectListAux, gotLocus, Val(AValue)
                    ElseIf Left(AName, 3) = "SG#" Then
                        curSG = Val(Right(AName, Len(AName) - 4))
                        ObjectListAdd .ObjectListAux, gotSG, Val(AValue)
                    End If
                    
                    Select Case AName
                    Case "Type"
                        .Type = Val(AValue)
                    Case "Caption"
                        .Caption = AValue
                        .Charset = PickCharset(.Caption)
                    Case "Hide"
                        .Hide = FromBoolean(AValue)
                    Case "InDemo"
                        .InDemo = FromBoolean(AValue)
                    Case "Desc"
                        .Description = AValue
                    'Case "Visible"
                    '    .Visible = FromBoolean(AValue)
                    Case "Fixed"
                        .Fixed = FromBoolean(AValue)
                    Case "Charset"
                        .Charset = Val(AValue)
                    Case "FontName"
                        .FontName = Trim(AValue)
                        If Not FontExists(.FontName) Then .FontName = "Arial"
                    Case "Forecolor"
                        .ForeColor = Val(AValue)
                    Case "FontSize"
                        .FontSize = Val(AValue)
                    Case "FontBold"
                        .FontBold = FromBoolean(AValue)
                    Case "FontItalic"
                        .FontItalic = FromBoolean(AValue)
                    Case "FontUnderline"
                        .FontUnderline = FromBoolean(AValue)
                    Case "InitiallyVisible"
                        .InitiallyVisible = FromBoolean(AValue)
                        .CurrentState = CLng(.InitiallyVisible)
                    Case "Message"
                        .Message = ToMultiLine(AValue)
                    Case "Remind"
                        .RemindToSaveFile = FromBoolean(AValue)
                    Case "Path"
                        .Path = AValue
                    Case "X"
                        .LogicalPosition.P1.X = Val(AValue)
                    Case "Y"
                        .LogicalPosition.P1.Y = Val(AValue)
                    End Select
                End With
                
            Case 8 ' [Misc]
                Select Case AName
                Case "AltP"
                    privNoAlter = FromBoolean(AValue)
                End Select
                
                
        End Select
    Loop
Close #1





For Z = 0 To FigureCount - 1
    If Figures(Z).FigureType = dsAnPoint Then
        Figures(Z).XTree = BuildTree(Figures(Z).XS)
        If Figures(Z).XTree.Erroneous Then
            If InDebugMode Then
                S = ""
                S = S & "Error in OpenFile: Recalc loop after input loop;" & vbCrLf
                S = S & "Figures(" & Z & ").XS = " & Figures(Z).XS & vbCrLf
                S = S & "Error = " & GetString(ResEvalErrorBase - 2 + 2 * WasThereAnErrorEvaluatingLastExpression) & vbCrLf
                S = S & "Err = " & ERR.Number & vbCrLf
                S = S & "Error$ = " & ERR.Description & vbCrLf
                S = S & "Figures(" & Z & ").XTree.BranchCount = " & Figures(Z).XTree.BranchCount & vbCrLf
                MsgBox S, vbExclamation
            End If
            Figures(Z).XTree = BuildTree(Trim(Str(BasePoint(Figures(Z).Points(0)).X)))
        End If
        Figures(Z).YTree = BuildTree(Figures(Z).YS)
        If Figures(Z).YTree.Erroneous Then
            If InDebugMode Then
                S = ""
                S = S & "Error in OpenFile: Recalc loop after input loop;" & vbCrLf
                S = S & "Figures(" & Z & ").YS = " & Figures(Z).YS & vbCrLf
                S = S & "Error = " & GetString(ResEvalErrorBase - 2 + 2 * WasThereAnErrorEvaluatingLastExpression) & vbCrLf
                S = S & "Err = " & ERR.Number & vbCrLf
                S = S & "Error$ = " & ERR.Description & vbCrLf
                S = S & "Figures(" & Z & ").YTree.BranchCount = " & Figures(Z).YTree.BranchCount & vbCrLf
                MsgBox S, vbExclamation
            End If
            Figures(Z).YTree = BuildTree(Trim(Str(BasePoint(Figures(Z).Points(0)).Y)))
        End If
    End If
    If Figures(Z).FigureType = dsDynamicLocus Then BuildDynamicLocusDependency Figures(Z), Figures(Z).Points(0), Figures(Z).Points(1)
    RecalcAuxInfo Z
Next

RecalcStaticGraphics
RefillObjectDescriptions
SnapAllPointNamesToPoints

For Z = 1 To LabelCount
    ParseLabel Z
    GetLabelSize Z
Next

For Z = 1 To ButtonCount
    GetButtonSize Z
Next

If (WorkArea.P2.X > WorkArea.P1.X) And (WorkArea.P2.Y > WorkArea.P1.Y) Then CalculateDefaultTransform WorkArea
AutorunButtons

'FormMain.mnuPointList.Enabled = PointCount > 0
'FormMain.mnuFigureList.Enabled = FigureCount > 0

If ShowProgress Then
    ProgressClose
    If FormMain.Enabled And FormMain.Visible Then FormMain.SetFocus
End If

ClearUndoBuffer
IsDirty = False

ShowAll
ErrStr = "OK"

OpenFile = True
Exit Function

EH:
Reset
If ShowProgress Then ProgressClose
MsgBox GetString(ResMsgCannotOpenFile) & vbCrLf & GetString(ResError) & ": " & ErrStr & vbCrLf & ERR.Description & vbCrLf & "Line number " & CurLine & ":" & vbCrLf & A, vbOKOnly + vbCritical, GetString(ResCaption)
ClearAll
DrawingName = GetString(ResUntitled)
FormMain.Caption = GetString(ResCaption) + " - " + GetString(ResUntitled)
If FormMain.Enabled And FormMain.Visible Then FormMain.SetFocus
End Function

Public Sub SaveFile(ByVal FName As String, Optional ByVal ShowProgress As Boolean = True)
On Local Error GoTo EH:
Dim TotalProgress As Long, CurrentProgress As Long, TempIsVisual As Boolean, Z As Long, Q As Long, NumOfParents As Long
Const ProgressStep = 10

If ShowProgress Then
    TotalProgress = PointCount + FigureCount + LocusCount
    ProgressShow GetString(ResSaving) & " " & RetrieveName(FName) & "..."
End If

Open FName For Output As #1
    Print #1, "[General]"
    Print #1, "FileFormatVersion=" & App.Major & "." & App.Minor & "." & App.Revision
    Print #1, "PointCount=" & PointCount
    Print #1, "FigureCount=" & FigureCount
    Print #1, "LabelCount=" & LabelCount
    Print #1, "WatchExpressionCount=" & WECount
    Print #1, "StaticGraphicCount=" & StaticGraphicCount
    Print #1, "ButtonCount=" & ButtonCount
    Print #1, "LocusCount=" & LocusCount
    Print #1, "PaperColor1=" & nPaperColor1
    Print #1, "PaperColor2=" & nPaperColor2
    Print #1, "AxesColor=" & nAxesColor
    Print #1, "GridColor=" & nGridColor
    Print #1, "GradientPaper=" & ToBoolean(nGradientPaper)
    Print #1, "ShowAxes=" & ToBoolean(nShowAxes)
    Print #1, "ShowGrid=" & ToBoolean(nShowGrid)
    Print #1, "Description=" & ToSingleLine(nDescription)
    Print #1, "Window=" & Trim(PaperScaleWidth) & "x" & Trim(PaperScaleHeight)
    Print #1, "WorkArea=(" & Trim(Str(CanvasBorders.P1.X)) & "," & Trim(Str(CanvasBorders.P1.Y)) & "," & Trim(Str(CanvasBorders.P2.X)) & "," & Trim(Str(CanvasBorders.P2.Y)) & ")"
    Print #1, "DemoInterval=" & Trim(nDemoInterval)
    Print #1,
    
    If PointCount <> 0 Then
        For Z = 1 To PointCount
            Print #1, "[Point" & Z & "]"
            
            'Print #1, "DrawWidth=" & BasePoint(Z).DrawWidth
            If Not BasePoint(Z).Enabled Then Print #1, "Enabled=False" ' & ToBoolean(BasePoint(Z).Enabled)
            Print #1, "FillColor=" & BasePoint(Z).FillColor
            Print #1, "FillStyle=" & BasePoint(Z).FillStyle
            Print #1, "ForeColor=" & BasePoint(Z).ForeColor
            If BasePoint(Z).Hide Then Print #1, "Hide=True" ' & ToBoolean(BasePoint(Z).Hide)
            If Not BasePoint(Z).InDemo Then Print #1, "InDemo=False"
            If BasePoint(Z).Description <> GetObjectDescription(gotPoint, Z) Then Print #1, "Desc=" & BasePoint(Z).Description
            If BasePoint(Z).LabelOffsetX <> 0 Then Print #1, "LabelOffsetX=" & Trim(Str(BasePoint(Z).LabelOffsetX))
            If BasePoint(Z).LabelOffsetY <> -setdefPointSize \ 2 + 1 Then Print #1, "LabelOffsetY=" & Trim(Str(BasePoint(Z).LabelOffsetY))
            'Print #1, "LabelOffsetY=" & Trim(Str(BasePoint(Z).LabelOffsetY))
            If BasePoint(Z).Locus <> 0 Then Print #1, "Locus=" & BasePoint(Z).Locus
            Print #1, "Name=" & BasePoint(Z).Name
            Print #1, "NameColor=" & BasePoint(Z).NameColor
            If BasePoint(Z).ParentFigure <> 0 Then Print #1, "ParentFigure=" & BasePoint(Z).ParentFigure
            Print #1, "PhysicalWidth=" & Trim(Str(BasePoint(Z).PhysicalWidth))
            Print #1, "Shape=" & BasePoint(Z).Shape
            If BasePoint(Z).ShowCoordinates Then Print #1, "ShowCoordinates=True" ' & ToBoolean(BasePoint(Z).ShowCoordinates)
            Print #1, "ShowName=" & ToBoolean(BasePoint(Z).ShowName)
            If BasePoint(Z).Tag <> 0 Then Print #1, "Tag=" & BasePoint(Z).Tag
            Print #1, "Type=" & BasePoint(Z).Type
            If Not BasePoint(Z).Visible Then Print #1, "Visible=False" ' & ToBoolean(BasePoint(Z).Visible)
            'Print #1, "Width=" & Str(BasePoint(Z).Width)
            Print #1, "X=" & Trim(Str(BasePoint(Z).X))
            Print #1, "Y=" & Trim(Str(BasePoint(Z).Y))
            Print #1, "ZOrder=" & BasePoint(Z).ZOrder
            
            If Z <> PointCount Then Print #1,
            
            If Z Mod ProgressStep = 0 Then
                CurrentProgress = CurrentProgress + ProgressStep
                ProgressUpdate CurrentProgress / TotalProgress
            End If
        Next Z
    End If
    
    '==================================================
    '   MISC SECTION
    '==================================================
    
    Print #1, "[Misc]"
    Print #1, "AltP=" & ToBoolean(privNoAlter)
    Print #1,
    
    '==================================================
    '   FIGURES
    '==================================================
    
    If FigureCount <> 0 Then
        For Z = 0 To FigureCount - 1
            Print #1, "[Figure" & Z & "]"
            TempIsVisual = IsVisual(Z)
            
            If TempIsVisual Then
                If Figures(Z).DrawMode <> defFigureDrawMode Then Print #1, "DrawMode=" & Figures(Z).DrawMode
                If Figures(Z).DrawStyle <> defFigureDrawStyle Then Print #1, "DrawStyle=" & Figures(Z).DrawStyle
                Print #1, "DrawWidth=" & Figures(Z).DrawWidth
                If Figures(Z).FillColor <> colPolygonFillColor Then Print #1, "FillColor=" & Figures(Z).FillColor
                If Figures(Z).FillStyle <> 6 Then Print #1, "FillStyle=" & Figures(Z).FillStyle
            End If
            If Figures(Z).FigureType = dsMeasureAngle Then Print #1, "DrawStyle=" & Figures(Z).DrawStyle
            Print #1, "FigureName=" & Figures(Z).Name
            Print #1, "FigureType=" & Figures(Z).FigureType
            Print #1, "FigureTypeString=" & GetString(ResFigureBase + Figures(Z).FigureType * 2, langEnglish)
            If TempIsVisual Or Figures(Z).FigureType = dsMeasureAngle Then Print #1, "ForeColor=" & Figures(Z).ForeColor
            Print #1, "NumberOfChildren=" & Figures(Z).NumberOfChildren
            Print #1, "NumberOfPoints=" & Figures(Z).NumberOfPoints
            If Not Figures(Z).Visible Then Print #1, "Visible=False" ' & ToBoolean(Figures(Z).Visible)
            If Figures(Z).XS <> "" Then Print #1, "XS=" & Figures(Z).XS
            If Figures(Z).YS <> "" Then Print #1, "YS=" & Figures(Z).YS
            If Figures(Z).Hide Then Print #1, "Hide=True" ' & ToBoolean(Figures(Z).Hide)
            If Not Figures(Z).InDemo Then Print #1, "InDemo=False"
            If Figures(Z).Description <> GetObjectDescription(gotFigure, Z) Then Print #1, "Desc=" & Figures(Z).Description
            If TempIsVisual Then Print #1, "ZOrder=" & Figures(Z).ZOrder
            
            If Figures(Z).NumberOfChildren <> 0 Then
                For Q = 0 To Figures(Z).NumberOfChildren - 1
                    Print #1, "Children" & Q & "=" & Figures(Z).Children(Q)
                Next
            End If
            NumOfParents = GetProperParentNumber(Figures(Z).FigureType)
            If NumOfParents <> 0 Then
                For Q = 0 To NumOfParents - 1
                    Print #1, "Parents" & Q & "=" & Figures(Z).Parents(Q)
                Next
            End If
            For Q = 0 To Figures(Z).NumberOfPoints - 1
                Print #1, "Points" & Q & "=" & Figures(Z).Points(Q)
            Next
            For Q = 1 To AuxCount
                If Figures(Z).AuxInfo(Q) <> 0 Then Print #1, "AuxInfo(" & Q & ")=" & Trim(Str(Figures(Z).AuxInfo(Q)))
                If Figures(Z).AuxPoints(Q).X <> 0 Then Print #1, "AuxPoints(" & Q & ").X=" & Trim(Str(Figures(Z).AuxPoints(Q).X))
                If Figures(Z).AuxPoints(Q).Y <> 0 Then Print #1, "AuxPoints(" & Q & ").Y=" & Trim(Str(Figures(Z).AuxPoints(Q).Y))
            Next
            
            Print #1,
            
            If Z Mod ProgressStep = 0 Then
                CurrentProgress = CurrentProgress + ProgressStep
                ProgressUpdate CurrentProgress / TotalProgress
            End If
        Next Z
    End If
    
    If LabelCount <> 0 Then
        For Z = 1 To LabelCount
            Print #1, "[Label" & Z & "]"
            
            If TextLabels(Z).Borders <> 0 Then Print #1, "Borders=" & TextLabels(Z).Borders
            Print #1, "Caption=" & ToSingleLine(TextLabels(Z).Caption)
            Print #1, "Charset=" & TextLabels(Z).Charset
            Print #1, "FontName=" & TextLabels(Z).FontName
            Print #1, "FontSize=" & TextLabels(Z).FontSize
            If TextLabels(Z).FontBold Then Print #1, "FontBold=" & ToBoolean(TextLabels(Z).FontBold)
            If TextLabels(Z).FontItalic Then Print #1, "FontItalic=" & ToBoolean(TextLabels(Z).FontItalic)
            If TextLabels(Z).FontUnderline Then Print #1, "FontUnderline=" & ToBoolean(TextLabels(Z).FontUnderline)
            Print #1, "ForeColor=" & TextLabels(Z).ForeColor
            If TextLabels(Z).BackColor <> 0 Then Print #1, "BackColor=" & TextLabels(Z).BackColor
            If Not TextLabels(Z).Transparent Then Print #1, "Transparent=False"
            'If Not TextLabels(Z).Visible Then Print #1, "Visible=" & ToBoolean(TextLabels(Z).Visible)
            If TextLabels(Z).Hide Then Print #1, "Hide=True"
            If Not TextLabels(Z).InDemo Then Print #1, "InDemo=False"
            If TextLabels(Z).Description <> GetObjectDescription(gotLabel, Z) Then Print #1, "Desc=" & TextLabels(Z).Description
            If TextLabels(Z).Fixed Then Print #1, "Fixed=True"
            Print #1, "X=" & Trim(Str(TextLabels(Z).LogicalPosition.P1.X))
            Print #1, "Y=" & Trim(Str(TextLabels(Z).LogicalPosition.P1.Y))
            
            Print #1,
        Next
    End If
    
    If WECount > 0 Then
        For Z = 1 To WECount
            Print #1, "[WatchExpression" & Z & "]"
            
            Print #1, "Name=" & WatchExpressions(Z).Name
            Print #1, "Expression=" & WatchExpressions(Z).Expression
            
            Print #1,
        Next Z
    End If
    
    If LocusCount > 0 Then
        For Z = 1 To LocusCount
            Print #1, "[Locus" & Z & "]"
            
            Print #1, "DrawWidth=" & Locuses(Z).DrawWidth
            Print #1, "Dynamic=" & ToBoolean(Locuses(Z).Dynamic)
            Print #1, "Enabled=" & ToBoolean(Locuses(Z).Enabled)
            Print #1, "ForeColor=" & Locuses(Z).ForeColor
            Print #1, "LocusPointCount=" & Locuses(Z).LocusPointCount
            Print #1, "ParentPoint=" & Locuses(Z).ParentPoint
            Print #1, "ParentFigure=" & Locuses(Z).ParentFigure
            Print #1, "Type=" & Locuses(Z).Type
            If Locuses(Z).Hide Then Print #1, "Hide=True"
            Print #1, "Visible=" & ToBoolean(Locuses(Z).Visible)
            If Not Locuses(Z).Dynamic Then
                
                If Locuses(Z).LocusNumber > 1 Then
                    Dim Sa() As String
                    
                    ReDim Sa(1 To Locuses(Z).LocusNumber)
                    For Q = 1 To Locuses(Z).LocusNumber
                        Sa(Q) = Locuses(Z).LocusNumbers(Q)
                    Next Q
                    
                    Print #1, "Pieces=" & Join(Sa, ";")
                End If
                
                For Q = 1 To Locuses(Z).LocusPointCount
                    Print #1, "Point" & Q & ".X=" & Trim(Str(Locuses(Z).LocusPoints(Q).X))
                    Print #1, "Point" & Q & ".Y=" & Trim(Str(Locuses(Z).LocusPoints(Q).Y))
                Next Q
            End If
            
            Print #1,
            
            If Z Mod ProgressStep = 0 Then
                CurrentProgress = CurrentProgress + ProgressStep
                ProgressUpdate CurrentProgress / TotalProgress
            End If
        Next Z
    End If
    
    If StaticGraphicCount > 0 Then
        For Z = 1 To StaticGraphicCount
            With StaticGraphics(Z)
                Print #1, "[SG" & Z & "]"
                
                Print #1, "DrawMode=" & .DrawMode
                Print #1, "DrawStyle=" & .DrawStyle
                Print #1, "DrawWidth=" & .DrawWidth
                Print #1, "FillColor=" & .FillColor
                Print #1, "FillStyle=" & .FillStyle
                Print #1, "ForeColor=" & .ForeColor
                Print #1, "NumberOfPoints=" & .NumberOfPoints
                Print #1, "Type=" & .Type
                If .Hide Then Print #1, "Hide=True"
                If Not .InDemo Then Print #1, "InDemo=False"
                If .Description <> GetObjectDescription(gotSG, Z) Then Print #1, "Desc=" & .Description
                Print #1, "Visible=" & ToBoolean(.Visible)
                For Q = 1 To .NumberOfPoints
                    Print #1, "Point" & Q & "=" & .Points(Q)
                Next
                
                Print #1,
            End With
        Next
    End If
    
    If ButtonCount > 0 Then
        For Z = 1 To ButtonCount
            With Buttons(Z)
                Print #1, "[Button" & Z & "]"
                
                Print #1, "Type=" & .Type
                Print #1, "Caption=" & .Caption
                If .Hide Then Print #1, "Hide=True"
                If Not .InDemo Then Print #1, "InDemo=False"
                If .Description <> GetObjectDescription(gotButton, Z) Then Print #1, "Desc=" & .Description
                'If Not .Visible Then Print #1, "Visible=" & ToBoolean(.Visible)
                If .Fixed Then Print #1, "Fixed=" & ToBoolean(.Fixed)
                If .Charset <> 0 Then Print #1, "Charset=" & .Charset
                Print #1, "FontName=" & .FontName
                Print #1, "FontSize=" & .FontSize
                If .FontBold Then Print #1, "FontBold=" & ToBoolean(.FontBold)
                If .FontItalic Then Print #1, "FontItalic=" & ToBoolean(.FontItalic)
                If .FontUnderline Then Print #1, "FontUnderline=" & ToBoolean(.FontUnderline)
                If .ForeColor <> 0 Then Print #1, "Forecolor=" & .ForeColor
                Print #1, "X=" & Trim(Str(.LogicalPosition.P1.X))
                Print #1, "Y=" & Trim(Str(.LogicalPosition.P1.Y))
                
                Select Case .Type
                Case butShowHide
                    If .InitiallyVisible Then Print #1, "InitiallyVisible=" & ToBoolean(.InitiallyVisible)
                    With .ObjectListAux
                        If .ButtonCount > 0 Then
                            For Q = 1 To .ButtonCount
                                Print #1, "Button#" & Q & "=" & .Buttons(Q)
                            Next
                        End If
                        If .FigureCount > 0 Then
                            For Q = 1 To .FigureCount
                                Print #1, "Figure#" & Q & "=" & .Figures(Q)
                            Next
                        End If
                        If .LabelCount > 0 Then
                            For Q = 1 To .LabelCount
                                Print #1, "Label#" & Q & "=" & .Labels(Q)
                            Next
                        End If
                        If .LocusCount > 0 Then
                            For Q = 1 To .LocusCount
                                Print #1, "Locus#" & Q & "=" & .Loci(Q)
                            Next
                        End If
                        If .PointCount > 0 Then
                            For Q = 1 To .PointCount
                                Print #1, "Point#" & Q & "=" & .Points(Q)
                            Next
                        End If
                        If .SGCount > 0 Then
                            For Q = 1 To .SGCount
                                Print #1, "SG#" & Q & "=" & .SGs(Q)
                            Next
                        End If
                    End With
                Case butMsgBox
                    Print #1, "Message=" & ToSingleLine(.Message)
                Case butPlaySound
                    Print #1, "Path=" & .Path
                Case butLaunchFile
                    Print #1, "Path=" & .Path
                    If .RemindToSaveFile Then Print #1, "Remind=True"
                End Select
                
                Print #1,
                
            End With
        Next
    End If
Close #1

If ShowProgress Then
    ProgressClose
    If FormMain.Enabled And FormMain.Visible Then FormMain.SetFocus
End If

If Dir(FName) = "" Then GoTo EH:
IsDirty = False
Exit Sub

EH:
Reset
MsgBox GetString(ResUnableToSaveFile) & vbCrLf & ERR.Description, vbExclamation
End Sub

Public Sub SaveBMP(ByVal FName As String)
'If Not Paper.AutoRedraw Then
'    Paper.AutoRedraw = True
'    PaperCls
'    ShowAll
'    SavePicture Paper.Image, FName
'    PaperCls
'    Paper.AutoRedraw = False
'    ShowAll
'Else
SavePicture Paper.Image, FName
'End If
End Sub

Public Sub SaveWMF(ByVal FName As String)
Dim hDC As Long, hMF As Long, hFont As Long, hOldFont As Long
Dim Z As Long
Dim Q As Long, i As Long, X As Double, Y As Double, tX As Double, tY As Double
Dim P As POINTAPI
Dim hPen As Long, hOldPen As Long, Col As Long

hDC = CreateMetaFile(vbNullString)

PrepareHDC hDC
hFont = CreateFont(setdefPointFontSize * -20 / Screen.TwipsPerPixelY, _
    0, 0, 0, IIf(setdefPointFontBold, 700, 400), -CLng(setdefPointFontItalic), -CLng(setdefPointFontUnderline), 0, 0, _
    0, 0, 2, 0, setdefPointFontName & vbNullChar)
hOldFont = SelectObject(hDC, hFont)

If Not setTransparentEMF Then PaperCls hDC

'==============================================
'ShowAll hDC

If setWallpaper <> "" Then DrawWallPaper hDC Else If nGradientPaper Then Gradient hDC, nPaperColor1, nPaperColor2, 0, 0, PaperScaleWidth, PaperScaleHeight, False ' Else PaperCls

'--------------------------------------------------------------------------------------------
If nShowGrid Then
    Col = EnsureRGB(nGridColor)
    
    hPen = CreatePen(PS_SOLID, 1, nGridColor)
    hOldPen = SelectObject(hDC, hPen)
    
    Q = Round(CanvasBorders.P1.X)
    For X = Round(CanvasBorders.P1.X) To Round(CanvasBorders.P2.X)
        tX = X
        tY = 0
        ToPhysical tX, tY
        MoveToEx hDC, tX, 0, P
        LineTo hDC, tX, PaperScaleHeight
    Next
    
    Q = Round(CanvasBorders.P1.Y)
    For Y = Round(CanvasBorders.P1.Y) To Round(CanvasBorders.P2.Y)
        tX = 0
        tY = Y
        ToPhysical tX, tY
        MoveToEx hDC, 0, tY, P
        LineTo hDC, PaperScaleWidth, tY
    Next
    
    SelectObject hDC, hOldPen
    DeleteObject hPen
End If
'--------------------------------------------------------------------------------------------

If nShowAxes Then ShowAxes hDC

For Z = 1 To StaticGraphicCount
    DrawStaticGraphic hDC, Z
Next

For Z = 1 To LabelCount
    ShowLabel hDC, Z
Next

For Z = 1 To LocusCount
    ShowLocus hDC, Z
Next

'==============================================

If FigureCount > 0 Then
    For Z = 0 To FigureCount - 1
        OrderedFigures(Z) = -1
    Next
    For Z = 0 To FigureCount - 1
        Q = Figures(Z).ZOrder
        If Q >= 0 And Q < FigureCount Then
            If OrderedFigures(Q) = -1 Then
                OrderedFigures(Q) = Z
            Else
                Figures(Z).ZOrder = GenerateNewFigureZOrder
                OrderedFigures(Figures(Z).ZOrder) = Z
            End If
        Else
            Figures(Z).ZOrder = GenerateNewFigureZOrder
            OrderedFigures(Figures(Z).ZOrder) = Z
        End If
    Next
End If

For Z = 0 To FigureCount - 1
    DrawFigure hDC, OrderedFigures(Z), False
Next

'===============================================
If PointCount > 0 Then
    ReDim OrderedPoints(1 To PointCount)
    For Z = 1 To PointCount
        Q = BasePoint(Z).ZOrder
        If Q > 0 And Q <= PointCount Then
            If OrderedPoints(Q) = 0 Then
                OrderedPoints(Q) = Z
            Else
                BasePoint(Z).ZOrder = GenerateNewPointZOrder
                OrderedPoints(BasePoint(Z).ZOrder) = Z
            End If
        Else
            BasePoint(Z).ZOrder = GenerateNewPointZOrder
            OrderedPoints(BasePoint(Z).ZOrder) = Z
        End If
    Next
End If

For Z = 1 To PointCount
    ShowPoint hDC, OrderedPoints(Z)
Next

For Z = 1 To ButtonCount
    ShowButton hDC, Z
Next

'==============================================

SelectObject hDC, hOldFont
DeleteObject hFont

hMF = CloseMetaFile(hDC)
WriteMetafile FName, hMF, CInt(PaperScaleWidth), CInt(PaperScaleHeight)
DeleteMetaFile hMF

PaperCls
ShowAll
End Sub

Public Sub SaveEMF(ByVal FName As String)
Dim hDC As Long, hMF As Long, hFont As Long, hOldFont As Long, lpRect As RECT
Dim UnitMM As Double
UnitMM = Paper.ScaleX(1, vbPixels, vbMillimeters) * 100
lpRect.Left = 0
lpRect.Top = 0
lpRect.Right = Paper.ScaleWidth * UnitMM
lpRect.Bottom = Paper.ScaleHeight * UnitMM
hDC = CreateEnhMetaFile(hDC, FName & vbNullChar, lpRect, GetString(ResCaption) & vbNullChar & RetrieveName(DrawingName) & vbNullChar & vbNullChar)

PrepareHDC hDC
hFont = CreateFont(setdefPointFontSize * -20 / Screen.TwipsPerPixelY, _
    0, 0, 0, IIf(setdefPointFontBold, 700, 400), -CLng(setdefPointFontItalic), -CLng(setdefPointFontUnderline), 0, 0, _
    0, 0, 2, 0, setdefPointFontName & vbNullChar)
hOldFont = SelectObject(hDC, hFont)

If Not setTransparentEMF Then PaperCls hDC
ShowAll hDC

SelectObject hDC, hOldFont
DeleteObject hFont

hMF = CloseEnhMetaFile(hDC)

'Dim hWmf As Long
'If hMF <> 0 And MsgBox(GetString(ResSave) & " " & extWMF & "?", vbQuestion Or vbYesNo) = vbYes Then
'    hWmf = ConvertEMF2WMF(hMF)
'    If hWmf <> 0 Then
'        WriteMetafile Left(FName, Len(FName) - 3) & extWMF, hWmf, CInt(PaperScaleWidth), CInt(PaperScaleHeight)
'        DeleteMetaFile hWmf
'    End If
'End If

DeleteEnhMetaFile hMF
End Sub

Public Sub SaveJSPHTML(ByVal FName As String)
Dim Z As Long, Z2 As Long, A As String, B As String, LineNum As Long
Dim TempPoints() As Long, TempFigures() As Long
On Local Error GoTo EH:

ReDim TempPoints(1 To PointCount)
ReDim TempFigures(0 To FigureCount - 1)

Open FName For Output As #1
    Print #1, "<HTML>"
    Print #1,
    
    Print #1, "<HEAD>"
    Print #1, "    <TITLE>" & RetrieveName(DrawingName) & "</TITLE>"
    Print #1, "</HEAD>"
    Print #1,
    
    Print #1, "<BODY>"
    
    If nDescription <> "" Then
        Z = 1
        Z2 = InStr(Z, nDescription, vbCrLf)
        If Z2 = 0 Then
            Print #1, nDescription
        Else
            Do
                Print #1, Mid(nDescription, Z, Z2 - Z)
                Z = Z2 + 2
                Z2 = InStr(Z, nDescription, vbCrLf)
            Loop Until Z2 = 0
            Print #1, Right(nDescription, Len(nDescription) - Z + 1)
        End If
    End If
    
    Print #1, "<APPLET "
    Print #1, "    CODE = ""GSP.class"""
    Print #1, "    ARCHIVE = ""JSPDR3.JAR"""
    Print #1, "    CODEBASE = ""JSP"""
    Print #1, "    WIDTH = " & PaperScaleWidth
    Print #1, "    HEIGHT = " & PaperScaleHeight
    Print #1, "    ALIGN=LEFT>"
    Print #1,
    Print #1, "    <PARAM NAME=Offscreen VALUE=1 <!--Change this to 0 to turn off double-buffering-->"
    Print #1, "    <PARAM NAME=Frame VALUE=1 <!--Draw frame around applet window-->"
    Print #1, "    <PARAM NAME=BackRed VALUE=" & Red(nPaperColor1) & " <!--Red component of the background RGB-->"
    Print #1, "    <PARAM NAME=BackGreen VALUE=" & Green(nPaperColor1) & " <!--Green component of the background RGB-->"
    Print #1, "    <PARAM NAME=BackBlue VALUE=" & Blue(nPaperColor1) & " <!--Blue component of the background RGB-->"
    
    Print #1, "    <PARAM NAME=Construction VALUE="""
    LineNum = 1
    
    For Z = 1 To PointCount
        With BasePoint(Z)
            If .Type = dsPoint Then
                A = "{" & LineNum & "} "
                A = A & IIf(.Enabled Or .Hide, "", "Fixed") & "Point("
                A = A & Round(BasePoint(Z).PhysicalX) & "," & Round(BasePoint(Z).PhysicalY) & ")"
                
                B = "color(" & Red(.FillColor) & ", " & Green(.FillColor) & ", " & Blue(.FillColor) & ")"
                If .ShowName Then B = "label('" & .Name & "')," & B
                If .Locus > 0 Then
                    If Not Locuses(.Locus).Dynamic Then B = "traced," & B
                End If
                If .Hide Then B = "hidden," & B
                B = " [" & B & "];"
                A = A & B
                
                Print #1, A
                TempPoints(Z) = LineNum
                LineNum = LineNum + 1
            Else
                TempPoints(Z) = .ParentFigure
            End If
        End With
    Next
    
    
    For Z = 0 To FigureCount - 1
        If TempFigures(Z) = 0 Then
            With Figures(Z)
                TempFigures(Z) = LineNum
                Select Case Figures(Z).FigureType
                
                Case dsSegment
                    A = "{" & LineNum & "} Segment("
                    A = A & TempPoints(.Points(0)) & "," & TempPoints(.Points(1)) & ")"
                    
                    B = "color(" & Red(.ForeColor) & ", " & Green(.ForeColor) & ", " & Blue(.ForeColor) & ")"
                    If .DrawWidth > 1 Then B = "thick," & B
                    If .Hide Then B = "hidden"
                    B = " [" & B & "];"
                    A = A & B
                    
                    Print #1, A
                    LineNum = LineNum + 1
                
                Case dsRay
                    A = "{" & LineNum & "} Ray("
                    A = A & TempPoints(.Points(1)) & "," & TempPoints(.Points(0)) & ")"
                    
                    B = "color(" & Red(.ForeColor) & ", " & Green(.ForeColor) & ", " & Blue(.ForeColor) & ")"
                    If .DrawWidth > 1 Then B = "thick," & B
                    If .Hide Then B = "hidden"
                    B = " [" & B & "];"
                    A = A & B
                    
                    Print #1, A
                    LineNum = LineNum + 1
                
                Case dsLine_2Points
                    A = "{" & LineNum & "} Line("
                    A = A & TempPoints(.Points(0)) & "," & TempPoints(.Points(1)) & ")"
                    
                    B = "color(" & Red(.ForeColor) & ", " & Green(.ForeColor) & ", " & Blue(.ForeColor) & ")"
                    If .DrawWidth > 1 Then B = "thick," & B
                    If .Hide Then B = "hidden"
                    B = " [" & B & "];"
                    A = A & B
                    
                    Print #1, A
                    LineNum = LineNum + 1
                
                Case dsLine_PointAndPerpendicularLine
                    A = "{" & LineNum & "} Perpendicular("
                    A = A & TempFigures(.Parents(0)) & "," & TempPoints(.Points(0)) & ")"
                    
                    B = "color(" & Red(.ForeColor) & ", " & Green(.ForeColor) & ", " & Blue(.ForeColor) & ")"
                    If .DrawWidth > 1 Then B = "thick," & B
                    If .Hide Then B = "hidden"
                    B = " [" & B & "];"
                    A = A & B
                    
                    Print #1, A
                    LineNum = LineNum + 1
                
                Case dsLine_PointAndParallelLine
                    A = "{" & LineNum & "} Parallel("
                    A = A & TempFigures(.Parents(0)) & "," & TempPoints(.Points(0)) & ")"
                    
                    B = "color(" & Red(.ForeColor) & ", " & Green(.ForeColor) & ", " & Blue(.ForeColor) & ")"
                    If .DrawWidth > 1 Then B = "thick," & B
                    If .Hide Then B = "hidden"
                    B = " [" & B & "];"
                    A = A & B
                    
                    Print #1, A
                    LineNum = LineNum + 1
                
                Case dsPointOnFigure
                    If IsLine(.Parents(0)) Or IsCircle(.Parents(0)) Then
                        TempPoints(.Points(0)) = LineNum
                        ' Str inserted instead of format because Java needs 1.2 instead of 1,2
                        A = "{" & LineNum & "} Point on object(" & TempFigures(.Parents(0)) & "," & Str(IIf(Figures(.Parents(0)).FigureType >= dsSegment And Figures(.Parents(0)).FigureType <= dsLine_2Points, 1 - .AuxInfo(1), .AuxInfo(1))) & ") [color(" & Red(BasePoint(.Points(0)).FillColor) & ", " & Green(BasePoint(.Points(0)).FillColor) & ", " & Blue(BasePoint(.Points(0)).FillColor) & ")"
                        If BasePoint(.Points(0)).Hide Then A = A & ",hidden"
                        If BasePoint(.Points(0)).ShowName Then A = A & ",label('" & BasePoint(.Points(0)).Name & "')"
                        A = A & "];"
                    Else
                        
                    End If
                    
                    Print #1, A
                    LineNum = LineNum + 1
                    
                Case dsDynamicLocus
                    A = "{" & LineNum & "} Locus("
                    A = A & TempPoints(.Points(0)) & ","
                    A = A & TempPoints(.Points(1)) & ","
                    A = A & TempFigures(Figures(BasePoint(.Points(1)).ParentFigure).Parents(0)) & ","
                    A = A & Locuses(BasePoint(.Points(0)).Locus).LocusPointCount & ") [color(" & Red(Locuses(BasePoint(.Points(0)).Locus).ForeColor) & ", " & Green(Locuses(BasePoint(.Points(0)).Locus).ForeColor) & ", " & Blue(Locuses(BasePoint(.Points(0)).Locus).ForeColor) & ")"
                    If Locuses(BasePoint(.Points(0)).Locus).DrawWidth > 1 Then A = A & ",thick"
                    A = A & "];"
                    
                    Print #1, A
                    LineNum = LineNum + 1
                
                Case dsIntersect
                    If IsLine(.Parents(0)) And IsLine(.Parents(1)) Then
                        TempPoints(.Points(0)) = LineNum
                        A = "{" & LineNum & "} Intersect(" & TempFigures(.Parents(0)) & "," & TempFigures(.Parents(1)) & ") [color(" & Red(BasePoint(.Points(0)).FillColor) & ", " & Green(BasePoint(.Points(0)).FillColor) & ", " & Blue(BasePoint(.Points(0)).FillColor) & ")"
                        If BasePoint(.Points(0)).Hide Then A = A & ",hidden"
                        If BasePoint(.Points(0)).Locus > 0 Then
                            If Not Locuses(BasePoint(.Points(0)).Locus).Dynamic Then A = A & ",traced"
                        End If
                        If BasePoint(.Points(0)).ShowName Then A = A & ",label('" & BasePoint(.Points(0)).Name & "')"
                        A = A & "];"
                        
                        Print #1, A
                        LineNum = LineNum + 1
                    Else
                        TempPoints(.Points(0)) = LineNum
                        A = "{" & LineNum & "} Intersect1(" & Minimum(TempFigures(.Parents(0)), TempFigures(.Parents(1))) & "," & Maximum(TempFigures(.Parents(0)), TempFigures(.Parents(1))) & ") [color(" & Red(BasePoint(.Points(0)).FillColor) & ", " & Green(BasePoint(.Points(0)).FillColor) & ", " & Blue(BasePoint(.Points(0)).FillColor) & ")"
                        If BasePoint(.Points(0)).Hide Then A = A & ",hidden"
                        If BasePoint(.Points(0)).Locus > 0 Then
                            If Not Locuses(BasePoint(.Points(0)).Locus).Dynamic Then A = A & ",traced"
                        End If
                        If BasePoint(.Points(0)).ShowName Then A = A & ",label('" & BasePoint(.Points(0)).Name & "')"
                        A = A & "];"
                        
                        Print #1, A
                        LineNum = LineNum + 1
                        
                        TempPoints(.Points(1)) = LineNum
                        A = "{" & LineNum & "} Intersect2(" & Minimum(TempFigures(.Parents(0)), TempFigures(.Parents(1))) & "," & Maximum(TempFigures(.Parents(0)), TempFigures(.Parents(1))) & ") [color(" & Red(BasePoint(.Points(1)).FillColor) & ", " & Green(BasePoint(.Points(1)).FillColor) & ", " & Blue(BasePoint(.Points(1)).FillColor) & ")"
                        If BasePoint(.Points(1)).Hide Then A = A & ",hidden"
                        If BasePoint(.Points(1)).Locus > 0 Then
                            If Not Locuses(BasePoint(.Points(1)).Locus).Dynamic Then A = A & ",traced"
                        End If
                        If BasePoint(.Points(1)).ShowName Then A = A & ",label('" & BasePoint(.Points(1)).Name & "')"
                        A = A & "];"
                        
                        Print #1, A
                        LineNum = LineNum + 1
                    End If
                    
                Case dsCircle_CenterAndCircumPoint
                    A = "{" & LineNum & "} Circle("
                    A = A & TempPoints(.Points(0)) & "," & TempPoints(.Points(1)) & ")"
                    
                    B = "color(" & Red(.ForeColor) & ", " & Green(.ForeColor) & ", " & Blue(.ForeColor) & ")"
                    If .DrawWidth > 1 Then B = "thick," & B
                    If .Hide Then B = "hidden"
                    B = " [" & B & "];"
                    A = A & B
                    
                    Print #1, A
                    LineNum = LineNum + 1
                    
                    If .FillStyle <> 6 Then
                        A = "{" & LineNum & "} Circle interior("
                        A = A & (LineNum - 1) & ")"
                        
                        B = "color(" & Red(.FillColor) & ", " & Green(.FillColor) & ", " & Blue(.FillColor) & ")"
                        B = " [" & B & "];"
                        A = A & B
                        
                        Print #1, A
                        LineNum = LineNum + 1
                    End If
                    
                Case dsCircle_CenterAndTwoPoints
                    Z2 = FindFigureWithPoints(dsSegment, .Points(0), .Points(1))
                    If Z2 = -1 Then
                        Z2 = LineNum
                        Print #1, "{" & LineNum & "} Segment(" & TempPoints(.Points(0)) & "," & TempPoints(.Points(1)) & ") [hidden];"
                        LineNum = LineNum + 1
                        TempFigures(Z) = LineNum
                    Else
                        If Z2 < Z Then
                            Z2 = TempFigures(Z2)
                        Else
                            A = "{" & LineNum & "} Segment(" & TempPoints(Figures(Z2).Points(0)) & "," & TempPoints(Figures(Z2).Points(1)) & ") [color(" & Red(Figures(Z2).ForeColor) & ", " & Green(Figures(Z2).ForeColor) & ", " & Blue(Figures(Z2).ForeColor) & ")"
                            TempFigures(Z2) = LineNum
                            If Figures(Z2).Hide Then A = A & ",hidden"
                            If Figures(Z2).DrawWidth > 1 Then A = A & ",thick"
                            A = A & "];"
                            Print #1, A
                            Z2 = LineNum
                            LineNum = LineNum + 1
                            TempFigures(Z) = LineNum
                        End If
                    End If
                
                    A = "{" & LineNum & "} Circle by radius("
                    A = A & TempPoints(.Points(2)) & "," & Z2 & ")"
                    
                    B = "color(" & Red(.ForeColor) & ", " & Green(.ForeColor) & ", " & Blue(.ForeColor) & ")"
                    If .DrawWidth > 1 Then B = "thick," & B
                    If .Hide Then B = "hidden"
                    B = " [" & B & "];"
                    A = A & B
                    
                    Print #1, A
                    LineNum = LineNum + 1
                    
                    If .FillStyle <> 6 Then
                        A = "{" & LineNum & "} Circle interior("
                        A = A & (LineNum - 1) & ")"
                        
                        B = "color(" & Red(.FillColor) & ", " & Green(.FillColor) & ", " & Blue(.FillColor) & ")"
                        B = " [" & B & "];"
                        A = A & B
                        
                        Print #1, A
                        LineNum = LineNum + 1
                    End If
                    
                Case dsMiddlePoint
                    Z2 = FindFigureWithPoints(dsSegment, .Points(1), .Points(2))
                    If Z2 = -1 Then
                        Z2 = LineNum
                        Print #1, "{" & LineNum & "} Segment(" & TempPoints(.Points(1)) & "," & TempPoints(.Points(2)) & ") [hidden];"
                        LineNum = LineNum + 1
                    Else
                        If Z2 < Z Then
                            Z2 = TempFigures(Z2)
                        Else
                            A = "{" & LineNum & "} Segment(" & TempPoints(Figures(Z2).Points(0)) & "," & TempPoints(Figures(Z2).Points(1)) & ") [color(" & Red(Figures(Z2).ForeColor) & ", " & Green(Figures(Z2).ForeColor) & ", " & Blue(Figures(Z2).ForeColor) & ")"
                            TempFigures(Z2) = LineNum
                            If Figures(Z2).Hide Then A = A & ",hidden"
                            If Figures(Z2).DrawWidth > 1 Then A = A & ",thick"
                            A = A & "];"
                            Print #1, A
                            Z2 = LineNum
                            LineNum = LineNum + 1
                        End If
                    End If
                    
                    TempPoints(.Points(0)) = LineNum
                    A = "{" & LineNum & "} Midpoint(" & Z2 & ") [color(" & Red(BasePoint(.Points(0)).FillColor) & ", " & Green(BasePoint(.Points(0)).FillColor) & ", " & Blue(BasePoint(.Points(0)).FillColor) & ")"
                    If BasePoint(.Points(0)).Hide Then A = A & ",hidden"
                    If BasePoint(.Points(0)).Locus > 0 Then
                        If Not Locuses(BasePoint(.Points(0)).Locus).Dynamic Then A = A & ",traced"
                    End If
                    If BasePoint(.Points(0)).ShowName Then A = A & ",label('" & BasePoint(.Points(0)).Name & "')"
                    A = A & "];"
                    
                    Print #1, A
                    LineNum = LineNum + 1
                
                Case dsSimmPoint
                    TempPoints(.Points(0)) = LineNum
                    A = "{" & LineNum & "} Dilation(" & TempPoints(.Points(1)) & "," & TempPoints(.Points(2)) & ",-1) [color(" & Red(BasePoint(.Points(0)).FillColor) & ", " & Green(BasePoint(.Points(0)).FillColor) & ", " & Blue(BasePoint(.Points(0)).FillColor) & ")"
                    If BasePoint(.Points(0)).Hide Then A = A & ",hidden"
                    If BasePoint(.Points(0)).Locus > 0 Then
                        If Not Locuses(BasePoint(.Points(0)).Locus).Dynamic Then A = A & ",traced"
                    End If
                    If BasePoint(.Points(0)).ShowName Then A = A & ",label('" & BasePoint(.Points(0)).Name & "')"
                    A = A & "];"
                    
                    Print #1, A
                    LineNum = LineNum + 1
                    
                Case dsSimmPointByLine
                    TempPoints(.Points(0)) = LineNum
                    A = "{" & LineNum & "} Reflection(" & TempPoints(.Points(1)) & "," & TempFigures(.Parents(0)) & ") [color(" & Red(BasePoint(.Points(0)).FillColor) & ", " & Green(BasePoint(.Points(0)).FillColor) & ", " & Blue(BasePoint(.Points(0)).FillColor) & ")"
                    If BasePoint(.Points(0)).Hide Then A = A & ",hidden"
                    If BasePoint(.Points(0)).Locus > 0 Then
                        If Not Locuses(BasePoint(.Points(0)).Locus).Dynamic Then A = A & ",traced"
                    End If
                    If BasePoint(.Points(0)).ShowName Then A = A & ",label('" & BasePoint(.Points(0)).Name & "')"
                    A = A & "];"
                    
                    Print #1, A
                    LineNum = LineNum + 1
                    
                End Select
            End With
        End If
    Next
    
    '????? Several times adding the same required segment for midpoints and circles by radius
    
    For Z = 1 To StaticGraphicCount
        With StaticGraphics(Z)
            If .Type = sgPolygon Then
                A = "{" & LineNum & "} Polygon("
                For Z2 = 1 To .NumberOfPoints - 1
                    A = A & TempPoints(.Points(Z2)) & IIf(Z2 = .NumberOfPoints - 1, ")", ",")
                Next
                A = A & " [color(" & Red(.FillColor) & ", " & Green(.FillColor) & ", " & Blue(.FillColor) & ")];"
                
                Print #1, A
                LineNum = LineNum + 1
            End If
        End With
    Next
    
    Print #1, """>"
    Print #1, "    Sorry, this page requires a Java-compatible web browser."
    Print #1, "</APPLET>"
    
    For Z = 1 To LabelCount
    Next
    
    Print #1, "</BODY>"
    Print #1,
    
    Print #1, "</HTML>"
Close #1

Exit Sub
EH:
Reset
MsgBox GetString(ResError) & " #" & ERR.Number & ": " & ERR.Description, vbExclamation
End Sub

Public Function CanOverwriteFile(ByVal FName As String) As Boolean
Dim A As String

If Dir(FName) = "" Then CanOverwriteFile = True: Exit Function

Open FName For Input As #1
    Do While Not EOF(1)
        Line Input #1, A
        If A = "AltP=True" Then
            CanOverwriteFile = False
            Close #1
            Exit Function
        End If
    Loop
Close #1

CanOverwriteFile = True
End Function
