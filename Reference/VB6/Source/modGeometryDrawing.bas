Attribute VB_Name = "modDrawGeometry"
Option Explicit

Public Sub ShowAll(Optional ByVal hDC As Long = -1, Optional ByVal ShowGrayed As Boolean = False, Optional ByVal ShouldRefresh As Boolean = True)
Dim Z As Long
Dim Q As Long, i As Long

If hDC = -1 Then hDC = Paper.hDC
If setWallpaper <> "" Then DrawWallPaper hDC Else If nGradientPaper Then Gradient hDC, nPaperColor1, nPaperColor2, 0, 0, PaperScaleWidth, PaperScaleHeight, False ' Else PaperCls

If nShowGrid Then ShowGrid hDC, ShowGrayed
If nShowAxes Then ShowAxes hDC, ShowGrayed

For Z = 1 To StaticGraphicCount
    DrawStaticGraphic hDC, Z, , , ShowGrayed
Next

For Z = 1 To LabelCount
    ShowLabel hDC, Z, , , , ShowGrayed
Next

For Z = 1 To LocusCount
    ShowLocus hDC, Z, , , , ShowGrayed
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
    DrawFigure hDC, OrderedFigures(Z), False, , ShowGrayed
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
    ShowPoint hDC, OrderedPoints(Z), , , , ShowGrayed
Next

For Z = 1 To ButtonCount
    ShowButton hDC, Z, , , , , , ShowGrayed
Next

If ShouldRefresh Then Paper.Refresh
End Sub

Public Sub ShowProperAll()
If DragS.State = dscSelectObjects Then
    ShowSelectedAll TempObjectSelection
ElseIf DragS.State = dscMacroStateGivens Then
    ShowAllWithGivens
ElseIf DragS.State = dscMacroStateResults Then
    ShowAllWithResults
ElseIf DragS.State = dscDemo Then
    DrawSituation
Else
    ShowAll
End If
End Sub

Public Sub ShowSelectedAll(objList As ObjectList, Optional ByVal hDC As Long = -1)
Dim Z As Long

If hDC = -1 Then hDC = Paper.hDC
If setWallpaper <> "" Then DrawWallPaper hDC Else If nGradientPaper Then Gradient hDC, nPaperColor1, nPaperColor2, 0, 0, PaperScaleWidth, PaperScaleHeight, False ' Else PaperCls

If nShowGrid Then ShowGrid hDC
If nShowAxes Then ShowAxes hDC

With objList
    For Z = 1 To StaticGraphicCount
        If ObjectListFind(objList, gotSG, Z) = 0 Then DrawStaticGraphic hDC, Z Else DrawStaticGraphicSelected hDC, Z
    Next
    For Z = 1 To LabelCount
        If ObjectListFind(objList, gotLabel, Z) = 0 Then ShowLabel hDC, Z Else ShowSelectedLabel hDC, Z, False, False
    Next
    For Z = 1 To LocusCount
        If ObjectListFind(objList, gotLocus, Z) = 0 Then ShowLocus hDC, Z Else ShowLocusSelected hDC, Z
    Next
    For Z = 0 To FigureCount - 1
        If ObjectListFind(objList, gotFigure, Z) = 0 Then DrawFigure hDC, Z, False Else ShowSelectedFigure hDC, Z
    Next
    For Z = 1 To PointCount
        If ObjectListFind(objList, gotPoint, Z) = 0 Then ShowPoint hDC, Z Else ShowSelectedPoint hDC, Z
    Next
    For Z = 1 To ButtonCount
        If ObjectListFind(objList, gotButton, Z) = 0 Then ShowButton hDC, Z Else ShowButton hDC, Z, , , , True
    Next
End With

If setNoFlicker Then Paper.Refresh
End Sub

Public Sub HideAll(Optional ByVal hDC As Long = -1)
Dim Z As Long

If hDC = -1 Then hDC = Paper.hDC

If FigureCount > 0 Then
    For Z = 0 To FigureCount - 1
        HideFigure Z
    Next Z
End If
If PointCount > 0 Then
    For Z = 1 To PointCount
        ShowPoint hDC, Z, True
    Next Z
End If
If StaticGraphicCount > 0 Then
    For Z = 1 To StaticGraphicCount
        DrawStaticGraphic hDC, Z, False
    Next
End If
If LabelCount > 0 Then
    For Z = 1 To LabelCount
        ShowLabel hDC, Z, False
    Next
End If
If LocusCount > 0 Then
    For Z = 1 To LocusCount
        ShowLocus hDC, Z, False
    Next
End If
If ButtonCount > 0 Then
    For Z = 1 To ButtonCount
        ShowButton hDC, Z, False
    Next
End If
End Sub

Public Sub ShowPoint(ByVal hDC As Long, ByVal Index As Integer, Optional ByVal Hide As Boolean = False, Optional ByVal ShouldRefresh As Boolean = False, Optional ByVal ShowAnyway As Boolean = False, Optional ByVal ShowGrayed As Boolean = False, Optional ByVal DrawWidth As Long = 1)
On Local Error GoTo EH

Dim hBrush As Long, hOldBrush As Long
Dim hPen As Long, hOldPen As Long
Dim tLeft As Long, tTop As Long, tWidth As Long
Dim lpRect As RECT, Col As Long, sStr As String
Dim NameColor As Long

With BasePoint(Index)
    
    If Not .Visible Then Exit Sub
    If .Hide And Not ShowAnyway Then Exit Sub
    
    tLeft = .PhysicalX
    tTop = .PhysicalY
    tWidth = .PhysicalWidth \ 2
    
    '===========================================================
    '                               Draw point name
    '===========================================================
    If .ShowName Then
        sStr = .Name
        If .ShowCoordinates Then sStr = sStr & " (" & Format(.X, setFormatNumber) & "; " & Format(.Y, setFormatNumber) & ")"
        Col = .NameColor
        If ShowGrayed Then Col = GetGrayedColor(Col)
        
        If Not Hide Then
            SetTextColor hDC, Col
            SetTextAlign hDC, TA_BOTTOM Or TA_CENTER
            TextOut hDC, tLeft + .LabelOffsetX, tTop + .LabelOffsetY, sStr, Len(sStr)
        Else
            hBrush = CreateSolidBrush(nPaperColor1)
            lpRect.Left = tLeft + .LabelOffsetX - .LabelWidth \ 2 - 3
            lpRect.Top = tTop + .LabelOffsetY - .LabelHeight - 3
            lpRect.Right = lpRect.Left + .LabelWidth + 6
            lpRect.Bottom = lpRect.Top + .LabelHeight + 6
            FillRect hDC, lpRect, hBrush
            DeleteObject hBrush
        End If
    End If
    '===========================================================
    
    If Hide Then
        Col = nPaperColor1
    Else
        Col = .ForeColor
        If ShowGrayed Then Col = GetGrayedColor(Col)
    End If
    
    hPen = CreatePen(PS_SOLID, DrawWidth, Col)
    hOldPen = SelectObject(hDC, hPen)
    
    If .FillStyle = vbFSTransparent Then
        hBrush = GetStockObject(NULL_BRUSH)
    Else
        If Hide Then
            Col = nPaperColor1
        Else
            Col = .FillColor
            If ShowGrayed Then Col = GetGrayedColor(Col)
        End If
        hBrush = CreateSolidBrush(Col)
    End If
    hOldBrush = SelectObject(hDC, hBrush)
    
    If .Shape = vbShapeSquare Then
        Rectangle hDC, tLeft - tWidth, tTop - tWidth, tLeft + .PhysicalWidth - tWidth, tTop + .PhysicalWidth - tWidth
    Else
        Ellipse hDC, tLeft - tWidth, tTop - tWidth, tLeft + .PhysicalWidth - tWidth, tTop + .PhysicalWidth - tWidth
    End If
    
    SelectObject hDC, hOldBrush
    If .FillStyle <> vbFSTransparent Then DeleteObject hBrush
    SelectObject hDC, hOldPen
    DeleteObject hPen
    
    .Shown = False
End With

If ShouldRefresh Then Paper.Refresh
EH:
End Sub

Public Sub DrawFigure(ByVal hDC As Long, ByVal Index As Integer, Optional ByVal ShowPoints As Boolean = True, Optional ByVal ShowAnyway As Boolean = False, Optional ByVal ShowGrayed As Boolean = False)
Dim hPen As Long, hOldPen As Long, lpPoint As POINTAPI, Z As Long, Q As Long
Dim Col As Long, FillColor As Long
On Local Error Resume Next

With Figures(Index)
    .AlreadyHidden = False
    If Not .Visible Then Exit Sub
    If .Hide And Not ShowAnyway Then Exit Sub

    If ShowPoints Then
        For Z = 0 To .NumberOfPoints - 1
            If BasePoint(.Points(Z)).Visible = False Then
                If .FigureType <> dsIntersect And .FigureType <> dsPointOnFigure Then
                    For Q = 0 To .NumberOfPoints - 1
                        BasePoint(.Points(Q)).Shown = False
                    Next
                    Exit Sub
                End If
            End If
        Next Z
    End If
    
    Select Case .FigureType
        Case dsMiddlePoint, dsSimmPoint, dsSimmPointByLine, dsPointOnFigure, dsInvert, dsAnPoint
            If Not ShowPoints Then Exit Sub
            ShowPoint hDC, Figures(Index).Points(0), , , ShowAnyway, ShowGrayed
            BasePoint(Figures(Index).Points(0)).Shown = True
        
        Case dsIntersect
            If Not ShowPoints Then Exit Sub
            If ShowAnyway Then
                ShowPoint hDC, Figures(Index).Points(0), , , True, BasePoint(Figures(Index).Points(0)).Hide
                ShowPoint hDC, Figures(Index).Points(1), , , True, BasePoint(Figures(Index).Points(1)).Hide
            Else
                ShowPoint hDC, Figures(Index).Points(0)
                ShowPoint hDC, Figures(Index).Points(1)
            End If
            BasePoint(Figures(Index).Points(0)).Shown = True
            BasePoint(Figures(Index).Points(1)).Shown = True
        
        Case dsSegment, dsLine_2Points, dsBisector, dsRay, dsLine_PointAndParallelLine, dsLine_PointAndPerpendicularLine, dsAnLineGeneral, dsAnLineCanonic, dsAnLineNormal, dsAnLineNormalPoint
            If .DrawMode <> vbCopyPen Then SetROP2 hDC, .DrawMode
            Col = .ForeColor
            If ShowGrayed Then Col = GetGrayedColor(Col)
            hPen = CreatePen(.DrawStyle, .DrawWidth, Col)
            hOldPen = SelectObject(hDC, hPen)
            
            MoveToEx hDC, .AuxPoints(1).X, .AuxPoints(1).Y, lpPoint
            LineTo hDC, .AuxPoints(2).X, .AuxPoints(2).Y
            
            SelectObject hDC, hOldPen
            DeleteObject hPen
            If .DrawMode <> vbCopyPen Then SetROP2 hDC, vbCopyPen
        
        Case dsCircle_CenterAndCircumPoint, dsCircle_CenterAndTwoPoints
            If .DrawMode <> vbCopyPen Then SetROP2 hDC, .DrawMode
            Col = .ForeColor
            If ShowGrayed Then Col = GetGrayedColor(Col)
            FillColor = .FillColor
            If ShowGrayed Then FillColor = GetGrayedColor(FillColor)
            DrawCircle hDC, .AuxPoints(1).X, .AuxPoints(1).Y, .AuxInfo(1), , , Col, .DrawMode, .DrawStyle, .DrawWidth, FillColor, .FillStyle
            If .DrawMode <> vbCopyPen Then SetROP2 hDC, vbCopyPen
        
        Case dsAnCircle
            If .DrawMode <> vbCopyPen Then SetROP2 hDC, .DrawMode
            Col = .ForeColor
            If ShowGrayed Then Col = GetGrayedColor(Col)
            FillColor = .FillColor
            If ShowGrayed Then FillColor = GetGrayedColor(FillColor)
            DrawCircle hDC, .AuxPoints(1).X, .AuxPoints(1).Y, .AuxInfo(4), , , Col, .DrawMode, .DrawStyle, .DrawWidth, FillColor, .FillStyle
            If .DrawMode <> vbCopyPen Then SetROP2 hDC, vbCopyPen
        
        Case dsCircle_ArcCenterAndRadiusAndTwoPoints
            If .DrawMode <> vbCopyPen Then SetROP2 hDC, .DrawMode
            Col = .ForeColor
            If ShowGrayed Then Col = GetGrayedColor(Col)
            FillColor = .FillColor
            If ShowGrayed Then FillColor = GetGrayedColor(FillColor)
            DrawCircle hDC, .AuxPoints(1).X, .AuxPoints(1).Y, .AuxInfo(1), .AuxInfo(2), .AuxInfo(3), Col, .DrawMode, .DrawStyle, .DrawWidth, FillColor, .FillStyle
            If .DrawMode <> vbCopyPen Then SetROP2 hDC, vbCopyPen
            If ShowPoints Then
                ShowPoint hDC, Figures(Index).Points(5), , , ShowAnyway, ShowGrayed
                ShowPoint hDC, Figures(Index).Points(6), , , ShowAnyway, ShowGrayed
                BasePoint(Figures(Index).Points(5)).Shown = True
                BasePoint(Figures(Index).Points(6)).Shown = True
            End If

        Case dsMeasureDistance
            Col = .ForeColor
            If ShowGrayed Then Col = GetGrayedColor(Col)
            
            SetTextAlign hDC, TA_CENTER Or TA_BOTTOM
            SetTextColor hDC, Col
            PrintMeasure hDC, Format(.AuxInfo(1), setFormatDistance), .AuxPoints(3).X + .AuxPoints(6).X, .AuxPoints(3).Y + .AuxPoints(6).Y, .AuxInfo(2)

        Case dsMeasureAngle
            Col = .ForeColor
            If ShowGrayed Then Col = GetGrayedColor(Col)
            
            If .DrawStyle > 0 Then
                If .AuxInfo(2) < 10 Then .AuxInfo(2) = defAngleMarkRadius
                If Abs(.AuxInfo(1) - 90) < 0.5 Then
                    DrawRightAngleMark hDC, .AuxPoints(1).X, .AuxPoints(1).Y, .AuxPoints(2).X, .AuxPoints(2).Y, .AuxPoints(3).X, .AuxPoints(3).Y, Col, ((.DrawStyle - 1) Mod 3) + 1, .DrawWidth, .AuxInfo(2)
                Else
                    DrawAngleMark hDC, .AuxPoints(1).X, .AuxPoints(1).Y, .AuxPoints(2).X, .AuxPoints(2).Y, .AuxPoints(3).X, .AuxPoints(3).Y, Col, ((.DrawStyle - 1) Mod 3) + 1, .DrawWidth, .AuxInfo(2)
                End If
            End If
            
            If .DrawStyle < 4 Then
                SetTextAlign hDC, TA_CENTER Or TA_BOTTOM
                SetTextColor hDC, Col
                PrintMeasure hDC, Format(.AuxInfo(1), setFormatAngle) & "°", .AuxPoints(4).X + .AuxPoints(6).X, .AuxPoints(4).Y + .AuxPoints(6).Y, 0
            End If
            
    End Select
End With
End Sub
'
'Public Sub DrawFigure(ByVal Index As Integer)
'On Local Error Resume Next
'
'With Figures(Index)
'    .AlreadyHidden = False
'    If Not .Visible Or .Hide Then Exit Sub
'
'    For Z = 0 To .NumberOfPoints - 1
'        If BasePoint(.Points(Z)).Visible = False Then
'            If .FigureType <> dsIntersect And .FigureType <> dsPointOnFigure Then
'                For Q = 0 To .NumberOfPoints - 1
'                    BasePoint(.Points(Q)).Shown = False
'                Next
'                Exit Sub
'            End If
'        End If
'    Next Z
'
'    If .DrawMode <> vbCopyPen Then PaperDrawMode = .DrawMode: Paper.DrawMode = PaperDrawMode
'    If .DrawStyle <> 0 Then PaperDrawStyle = .DrawStyle: Paper.DrawStyle = PaperDrawStyle
'    If .DrawWidth <> 1 Then PaperDrawWidth = .DrawWidth: Paper.DrawWidth = PaperDrawWidth
'
'    Select Case .FigureType
'        Case dsSegment, dsLine_2Points, dsRay, dsLine_PointAndParallelLine, dsLine_PointAndPerpendicularLine, dsAnLineGeneral, dsAnLineCanonic, dsAnLineNormal, dsAnLineNormalPoint
'            Paper.Line (.AuxPoints(1).X, .AuxPoints(1).Y)-(.AuxPoints(2).X, .AuxPoints(2).Y), .ForeColor
'
'        Case dsCircle_CenterAndCircumPoint, dsCircle_CenterAndTwoPoints
'            DrawCircle .AuxPoints(1).X, .AuxPoints(1).Y, .AuxInfo(1), , , .ForeColor
'
'        Case dsAnCircle
'            DrawCircle .AuxPoints(1).X, .AuxPoints(1).Y, .AuxInfo(4), , , .ForeColor
'
'        Case dsCircle_ArcCenterAndRadiusAndTwoPoints
'            DrawCircle .AuxPoints(1).X, .AuxPoints(1).Y, .AuxInfo(1), .AuxInfo(2), .AuxInfo(3), .ForeColor
'            ShowPoint Paper.hDC, Figures(Index).Points(5)
'            ShowPoint Paper.hDC, Figures(Index).Points(6)
'            BasePoint(Figures(Index).Points(5)).Shown = True
'            BasePoint(Figures(Index).Points(6)).Shown = True
'
'        Case dsMiddlePoint, dsSimmPoint, dsSimmPointByLine, dsPointOnFigure, dsInvert
'            ShowPoint Paper.hDC, Figures(Index).Points(0)
'            BasePoint(Figures(Index).Points(0)).Shown = True
'
'        Case dsIntersect
'            ShowPoint Paper.hDC, Figures(Index).Points(0)
'            ShowPoint Paper.hDC, Figures(Index).Points(1)
'            BasePoint(Figures(Index).Points(0)).Shown = True
'            BasePoint(Figures(Index).Points(1)).Shown = True
'
'        Case dsMeasureDistance
'            Paper.ForeColor = .ForeColor
'            PrintMeasure Round(.AuxInfo(1), setDistancePrecision), .AuxPoints(3).X, .AuxPoints(3).Y, .AuxInfo(2)
'
'        Case dsMeasureAngle
'            Paper.ForeColor = .ForeColor
'            PrintMeasure Round(.AuxInfo(1), setAnglePrecision) & "°", .AuxPoints(4).X, .AuxPoints(4).Y, 0
'
'    End Select
'End With
'
'If PaperDrawMode <> vbCopyPen Then Paper.DrawMode = vbCopyPen
'If PaperDrawStyle <> vbSolid Then Paper.DrawStyle = vbSolid
'If PaperDrawWidth <> 1 Then Paper.DrawWidth = 1

Public Sub HideFigure(ByVal Index As Integer, Optional ByVal ShouldRecalc As Boolean = True, Optional HidePoints As Boolean = True)
Dim hPen As Long, hOldPen As Long, hDC As Long
Dim hBrush As Long, hOldBrush As Long, lpPoint As POINTAPI
On Local Error Resume Next

hDC = Paper.hDC

With Figures(Index)
    If .AlreadyHidden Then Exit Sub
    .AlreadyHidden = True
    
    Select Case .FigureType
        Case dsSegment, dsLine_2Points, dsRay, dsBisector, dsLine_PointAndParallelLine, dsLine_PointAndPerpendicularLine, dsAnLineGeneral, dsAnLineCanonic, dsAnLineNormal, dsAnLineNormalPoint
            If .DrawMode <> vbCopyPen Then SetROP2 Paper.hDC, .DrawMode
            'Paper.Line (.AuxPoints(1).X, .AuxPoints(1).Y)-(.AuxPoints(2).X, .AuxPoints(2).Y), nPaperColor1
            hPen = CreatePen(.DrawStyle, .DrawWidth, nPaperColor1)
            hOldPen = SelectObject(hDC, hPen)
            
            MoveToEx hDC, .AuxPoints(1).X, .AuxPoints(1).Y, lpPoint
            LineTo hDC, .AuxPoints(2).X, .AuxPoints(2).Y
            
            SelectObject hDC, hOldPen
            DeleteObject hPen
            If .DrawMode <> vbCopyPen Then SetROP2 hDC, vbCopyPen
        Case dsCircle_CenterAndCircumPoint, dsCircle_CenterAndTwoPoints
            DrawCircle hDC, .AuxPoints(1).X, .AuxPoints(1).Y, .AuxInfo(1), , , nPaperColor1
        Case dsAnCircle
            DrawCircle hDC, .AuxPoints(1).X, .AuxPoints(1).Y, .AuxInfo(4), , , nPaperColor1
        Case dsCircle_ArcCenterAndRadiusAndTwoPoints
            DrawCircle hDC, .AuxPoints(1).X, .AuxPoints(1).Y, .AuxInfo(1), .AuxInfo(2), .AuxInfo(3), nPaperColor1
            If HidePoints Then
                ShowPoint hDC, Figures(Index).Points(5), True
                ShowPoint hDC, Figures(Index).Points(6), True
            End If
        Case dsMiddlePoint, dsSimmPoint, dsSimmPointByLine, dsInvert, dsAnPoint, dsPointOnFigure
            If HidePoints Then ShowPoint hDC, Figures(Index).Points(0), True
        Case dsIntersect
            If HidePoints Then
                ShowPoint Paper.hDC, Figures(Index).Points(0), True
                ShowPoint Paper.hDC, Figures(Index).Points(1), True
            End If
        Case dsMeasureDistance
            Paper.ForeColor = Paper.BackColor
            SetBkMode Paper.hDC, OPAQUE
            PrintMeasure Paper.hDC, Format(.AuxInfo(1), setFormatDistance), .AuxPoints(3).X + .AuxPoints(6).X, .AuxPoints(3).Y + .AuxPoints(6).Y, .AuxInfo(2)
            SetBkMode Paper.hDC, Transparent
        Case dsMeasureAngle
            Paper.ForeColor = Paper.BackColor
            SetBkMode Paper.hDC, OPAQUE
            PrintMeasure Paper.hDC, Format(.AuxInfo(1), setFormatAngle) & "°", .AuxPoints(4).X + .AuxPoints(6).X, .AuxPoints(4).Y + .AuxPoints(6).Y, 0
            SetBkMode Paper.hDC, Transparent
    End Select
End With
If ShouldRecalc Then RecalcAuxInfo Index
End Sub

Public Sub ShowLocus(ByVal hDC As Long, ByVal Index As Long, Optional ByVal bShow As Boolean = True, Optional ByVal ShouldRefresh As Boolean = False, Optional ByVal ShowAnyway As Boolean = False, Optional ByVal ShowGrayed As Boolean = False)
Dim hPen As Long, hOldPen As Long, Col As Long, Z As Long
Dim hBrush As Long, nWidth As Long, nWidth1 As Long, nWidth2 As Long, lpRect As RECT

If Index < 1 Or Index > LocusCount Then Exit Sub

With Locuses(Index)
    If Not .Visible Then Exit Sub
    If .Hide And Not ShowAnyway Then Exit Sub
    If .LocusPointCount <= 1 Then Exit Sub
    
    If bShow Then
        Col = .ForeColor
        If ShowGrayed Then Col = GetGrayedColor(Col)
    Else
        Col = nPaperColor1
    End If
    
    If Locuses(Index).Type = 0 Then
        hPen = CreatePen(PS_SOLID, .DrawWidth, Col)
        hOldPen = SelectObject(hDC, hPen)
        
        If .LocusNumber <= 1 Then
            Polyline hDC, .LocusPixels(1), .LocusPointCount
        Else
            PolyPolyline hDC, .LocusPixels(1), .LocusNumbers(1), .LocusNumber
        End If
        
        SelectObject hDC, hOldPen
        DeleteObject hPen
    Else
        If nWidth = 1 Then
            For Z = 1 To .LocusPointCount
                SetPixelV hDC, .LocusPixels(Z).X, .LocusPixels(Z).Y, Col
            Next
        Else
            hBrush = CreateSolidBrush(Col)
            nWidth = .DrawWidth
            nWidth2 = nWidth \ 2
            nWidth1 = nWidth - nWidth2
            For Z = 1 To .LocusPointCount
                lpRect.Left = .LocusPixels(Z).X - nWidth2
                lpRect.Top = .LocusPixels(Z).Y - nWidth2
                lpRect.Right = .LocusPixels(Z).X + nWidth1
                lpRect.Bottom = .LocusPixels(Z).Y + nWidth1
                FillRect hDC, lpRect, hBrush
            Next
            DeleteObject hBrush
        End If
    End If
    
    If ShouldRefresh Then Paper.Refresh
End With
End Sub

Public Sub ShowLocusSelected(ByVal hDC As Long, ByVal Index As Long, Optional ByVal bShow As Boolean = True, Optional ByVal ShouldRefresh As Boolean = False)
Dim hPen As Long, hOldPen As Long, Col As Long, Z As Long, Q As Long, SCol As Long, DCol As Long
Dim hBrush As Long, nWidth As Long, nWidth2 As Long, lpRect As RECT
Dim DrawColorR(1 To 2) As Integer
Dim DrawColorG(1 To 2) As Integer
Dim DrawColorB(1 To 2) As Integer
Dim Coeff As Single

With Locuses(Index)
    If Not .Visible Then Exit Sub
    If .Hide Then Exit Sub
    
    If .Visible And .LocusPointCount > 1 Then
        OffsetObjectPixels .LocusPixels, Shadow, Shadow
        SCol = RGB(128, 128, 128)
        DCol = EnsureRGB(nPaperColor1)
        DrawColorR(1) = SCol And 255
        DrawColorG(1) = (SCol And 65535) \ 256
        DrawColorB(1) = (SCol \ 65536)
        DrawColorR(2) = (DCol And 255) - DrawColorR(1)
        DrawColorG(2) = ((DCol And 65535) \ 256) - DrawColorG(1)
        DrawColorB(2) = (DCol \ 65536) - DrawColorB(1)
        
        For Q = Shadow To 1 Step -1
            If .Type = 0 Then
                Coeff = Q / Shadow
                hPen = CreatePen(PS_SOLID, .DrawWidth, RGB(DrawColorR(1) + DrawColorR(2) * Coeff, DrawColorG(1) + DrawColorG(2) * Coeff, DrawColorB(1) + DrawColorB(2) * Coeff))
                hOldPen = SelectObject(hDC, hPen)
                
                If .LocusNumber <= 1 Then
                    Polyline hDC, .LocusPixels(1), .LocusPointCount
                Else
                    PolyPolyline hDC, .LocusPixels(1), .LocusNumbers(1), .LocusNumber
                End If
                
                SelectObject hDC, hOldPen
                DeleteObject hPen
            Else
                Coeff = Q / Shadow
                Col = RGB(DrawColorR(1) + DrawColorR(2) * Coeff, DrawColorG(1) + DrawColorG(2) * Coeff, DrawColorB(1) + DrawColorB(2) * Coeff)
                hBrush = CreateSolidBrush(Col)
                nWidth = .DrawWidth
                nWidth2 = nWidth \ 2
                For Z = 1 To .LocusPointCount
                    If nWidth = 1 Then
                        SetPixelV hDC, .LocusPixels(Z).X, .LocusPixels(Z).Y, Col
                    Else
                        lpRect.Left = .LocusPixels(Z).X - nWidth2
                        lpRect.Top = .LocusPixels(Z).Y - nWidth2
                        lpRect.Right = .LocusPixels(Z).X + nWidth2
                        lpRect.Bottom = .LocusPixels(Z).Y + nWidth2
                        FillRect hDC, lpRect, hBrush
                    End If
                Next
                DeleteObject hBrush
            End If
            
            OffsetObjectPixels .LocusPixels, -1, -1
        Next Q
        
        If ShouldRefresh Then Paper.Refresh
    End If
End With
End Sub

Public Sub DrawStaticGraphic(ByVal hDC As Long, ByVal Index As Long, Optional ByVal bShow As Boolean = True, Optional ByVal ShouldRefresh As Boolean = False, Optional ByVal ShowGrayed As Boolean = False, Optional ByVal ShowAnyway As Boolean = False)
Dim hBrush As Long, hOldBrush As Long
Dim hPen As Long, hOldPen As Long
Dim FillColor As Long, ForeColor As Long
Dim lpPoint As POINTAPI

If Index = 0 Then Exit Sub

With StaticGraphics(Index)
    If Not .Visible Then Exit Sub
    If .Hide And Not ShowAnyway Then Exit Sub
    Select Case .Type
        
        '===============================================
        '                                   POLYGON
        '===============================================
        Case sgPolygon
            If bShow Then
                FillColor = .FillColor
                If ShowGrayed Then FillColor = GetGrayedColor(FillColor)
                If .FillStyle = -1 Then
                    hBrush = CreateSolidBrush(FillColor)
                Else
                    hBrush = CreateHatchBrush(.FillStyle, .FillColor)
                End If
                ForeColor = .ForeColor
            Else
                hBrush = CreateSolidBrush(Paper.BackColor)
                ForeColor = Paper.BackColor
            End If
            
            If ShowGrayed Then ForeColor = GetGrayedColor(ForeColor)
            'hPen = CreatePen(.DrawStyle, .DrawWidth, ForeColor)
            hPen = CreatePen(PS_NULL, 1, 0)
            hOldPen = SelectObject(hDC, hPen)
            hOldBrush = SelectObject(hDC, hBrush)
            
            If .DrawMode <> 13 Then SetROP2 hDC, .DrawMode
            Polygon hDC, .ObjectPixels(1), .NumberOfPoints
            If .DrawMode <> 13 Then SetROP2 hDC, 13
            
            SelectObject hDC, hOldPen
            SelectObject hDC, hOldBrush
            DeleteObject hPen
            DeleteObject hBrush
            
        '===============================================
        '                                   BEZIER
        '===============================================
        Case sgBezier
            If bShow Then
                ForeColor = .ForeColor
            Else
                ForeColor = Paper.BackColor
            End If
            If ShowGrayed Then ForeColor = GetGrayedColor(ForeColor)
            hPen = CreatePen(.DrawStyle, .DrawWidth, ForeColor)
            hOldPen = SelectObject(hDC, hPen)
            
            If .DrawMode <> 13 Then SetROP2 hDC, .DrawMode
            PolyBezier hDC, .ObjectPixels(1), .NumberOfPoints
            If .DrawMode <> 13 Then SetROP2 hDC, 13
            
            SelectObject hDC, hOldPen
            DeleteObject hPen
        
        '===============================================
        '                                   VECTOR
        '===============================================
        Case sgVector
            If bShow Then
                FillColor = .FillColor
                If ShowGrayed Then FillColor = GetGrayedColor(FillColor)
                If .FillStyle = -1 Then
                    hBrush = CreateSolidBrush(.FillColor)
                Else
                    hBrush = CreateHatchBrush(.FillStyle, .FillColor)
                End If
                ForeColor = .ForeColor
            Else
                hBrush = CreateSolidBrush(Paper.BackColor)
                ForeColor = Paper.BackColor
            End If
            If ShowGrayed Then ForeColor = GetGrayedColor(ForeColor)
            hPen = CreatePen(.DrawStyle, .DrawWidth, ForeColor)
            hOldPen = SelectObject(hDC, hPen)
            hOldBrush = SelectObject(hDC, hBrush)
            
            If .DrawMode <> 13 Then SetROP2 hDC, .DrawMode
            'DrawArrow hDC, .ObjectPoints(1).X, .ObjectPoints(1).Y, .ObjectPoints(2).X, .ObjectPoints(2).Y
            Polygon hDC, .ObjectPixels(1), 3
            MoveToEx hDC, .ObjectPixels(1).X, .ObjectPixels(1).Y, lpPoint
            LineTo hDC, .ObjectPixels(1).X, .ObjectPixels(1).Y
            If .DrawMode <> 13 Then SetROP2 hDC, 13
            
            SelectObject hDC, hOldPen
            SelectObject hDC, hOldBrush
            DeleteObject hPen
            DeleteObject hBrush
    End Select
End With
If ShouldRefresh Then Paper.Refresh
End Sub

Public Sub DrawStaticGraphicSelected(ByVal hDC As Long, ByVal Index As Long, Optional ByVal bShow As Boolean = True, Optional ByVal ShouldRefresh As Boolean = False)
Dim hBrush As Long, hOldBrush As Long
Dim hPen As Long, hOldPen As Long
Dim Z As Long, Q As Long, SCol As Long, DCol As Long
Dim DrawColorR(1 To 2) As Integer
Dim DrawColorG(1 To 2) As Integer
Dim DrawColorB(1 To 2) As Integer
Dim Coeff As Single

If Index = 0 Then Exit Sub
With StaticGraphics(Index)
    If Not .Visible Then Exit Sub
    Select Case .Type
        Case sgPolygon
            '===============================================
            ' Begin by drawing the shadow
            '===============================================
            hPen = CreatePen(PS_NULL, 1, 0)
            hOldPen = SelectObject(hDC, hPen)
            
            SetROP2 hDC, 9
            OffsetObjectPixels .ObjectPixels, Shadow, Shadow
            
            SCol = RGB(128, 128, 128)
            DCol = EnsureRGB(nPaperColor1)
            DrawColorR(1) = SCol And 255
            DrawColorG(1) = (SCol And 65535) \ 256
            DrawColorB(1) = (SCol \ 65536)
            DrawColorR(2) = (DCol And 255) - DrawColorR(1)
            DrawColorG(2) = ((DCol And 65535) \ 256) - DrawColorG(1)
            DrawColorB(2) = (DCol \ 65536) - DrawColorB(1)
            
            For Z = Shadow To 1 Step -1
                Coeff = Z / Shadow
                hBrush = CreateSolidBrush(RGB(DrawColorR(1) + DrawColorR(2) * Coeff, DrawColorG(1) + DrawColorG(2) * Coeff, DrawColorB(1) + DrawColorB(2) * Coeff))
                hOldBrush = SelectObject(hDC, hBrush)
                Polygon hDC, .ObjectPixels(1), .NumberOfPoints
                SelectObject hDC, hOldBrush
                DeleteObject hBrush
                
                OffsetObjectPixels .ObjectPixels, -1, -1
            Next
            SelectObject hDC, hOldPen
            DeleteObject hPen
            
            SetROP2 hDC, 13
            
            '===============================================
            ' Now draw the polygon itself
            '===============================================
            hPen = CreatePen(PS_NULL, 1, .ForeColor)
            hOldPen = SelectObject(hDC, hPen)
            
            '===============================================
            ' Filled base
            '===============================================
            hBrush = CreateSolidBrush(.FillColor)
            hOldBrush = SelectObject(hDC, hBrush)
            Polygon hDC, .ObjectPixels(1), .NumberOfPoints
            SelectObject hDC, hOldBrush
            DeleteObject hBrush
            
            '===============================================
            ' And cross interior
            '===============================================
            hBrush = CreateHatchBrush(4, DarkenColor(.FillColor, 0.8))
            hOldBrush = SelectObject(hDC, hBrush)
            Polygon hDC, .ObjectPixels(1), .NumberOfPoints
            SelectObject hDC, hOldBrush
            DeleteObject hBrush
            
            SelectObject hDC, hOldPen
            DeleteObject hPen

        
        Case sgBezier
            SetROP2 hDC, 9
            OffsetObjectPixels .ObjectPixels, Shadow, Shadow
            
            SCol = RGB(192, 192, 192)
            DCol = EnsureRGB(nPaperColor1)
            DrawColorR(1) = SCol And 255
            DrawColorG(1) = (SCol And 65535) \ 256
            DrawColorB(1) = (SCol \ 65536)
            DrawColorR(2) = (DCol And 255) - DrawColorR(1)
            DrawColorG(2) = ((DCol And 65535) \ 256) - DrawColorG(1)
            DrawColorB(2) = (DCol \ 65536) - DrawColorB(1)
            
            For Z = Shadow To 1 Step -1
                Coeff = Z / Shadow
                hPen = CreatePen(PS_SOLID, 2, RGB(DrawColorR(1) + DrawColorR(2) * Coeff, DrawColorG(1) + DrawColorG(2) * Coeff, DrawColorB(1) + DrawColorB(2) * Coeff))
                hOldPen = SelectObject(hDC, hPen)
                PolyBezier hDC, .ObjectPixels(1), .NumberOfPoints
                SelectObject hDC, hOldPen
                DeleteObject hPen
                
                OffsetObjectPixels .ObjectPixels, -1, -1
            Next
            
            hPen = CreatePen(.DrawStyle, .DrawWidth, IIf(bShow, .ForeColor, Paper.BackColor))
            hOldPen = SelectObject(hDC, hPen)
            If .DrawMode <> 13 Then SetROP2 hDC, .DrawMode
            PolyBezier hDC, .ObjectPixels(1), .NumberOfPoints
            SelectObject hDC, hOldPen
            DeleteObject hPen
            SetROP2 hDC, 13
            
        Case sgVector
            If bShow Then
                If .FillStyle = -1 Then
                    hBrush = CreateSolidBrush(.FillColor)
                Else
                    hBrush = CreateHatchBrush(.FillStyle, .FillColor)
                End If
            Else
                hBrush = CreateSolidBrush(Paper.BackColor)
            End If
            hPen = CreatePen(.DrawStyle, .DrawWidth, IIf(bShow, .ForeColor, Paper.BackColor))
            hOldPen = SelectObject(hDC, hPen)
            hOldBrush = SelectObject(hDC, hBrush)
            
            If .DrawMode <> 13 Then SetROP2 hDC, .DrawMode
            'DrawArrow hDC, .ObjectPoints(1).X, .ObjectPoints(1).Y, .ObjectPoints(2).X, .ObjectPoints(2).Y
            Polygon hDC, .ObjectPixels(1), 3
            If .DrawMode <> 13 Then SetROP2 hDC, 13
            
            SelectObject hDC, hOldPen
            SelectObject hDC, hOldBrush
            DeleteObject hPen
            DeleteObject hBrush
    End Select
End With
If ShouldRefresh Then Paper.Refresh
End Sub

Public Sub ShowAxes(ByVal hDC As Long, Optional ByVal ShowGrayed As Boolean = False)
Const ArrowLen As Long = 8
Const ArrowWidth As Long = 2 * 2
Const AxesFontSize = 7
Const tS As String = "Arial"

Dim Col As Long
Dim hBrush As Long, hPen As Long, hOldBrush As Long, hOldPen As Long, lpPoint As POINTAPI
Dim hFont As Long, hOldFont As Long
Dim tAPI(1 To 3) As POINTAPI, LF As LOGFONT
Dim tX As Double, tY As Double, TI As Long, X As Long, Y As Long, Z As Long

Col = nAxesColor
If ShowGrayed Then Col = GetGrayedColor(Col)

hBrush = CreateSolidBrush(Col)
hPen = CreatePen(PS_SOLID, 1, Col)
hOldBrush = SelectObject(hDC, hBrush)
hOldPen = SelectObject(hDC, hPen)

If Sgn(CanvasBorders.P1.X) <> Sgn(CanvasBorders.P2.X) Then
    MoveToEx hDC, OriginX, 0, lpPoint
    LineTo hDC, OriginX, PaperScaleHeight
    tAPI(1).Y = ArrowLen
    tAPI(2).Y = ArrowLen
    tAPI(3).Y = 0
    tAPI(1).X = OriginX - ArrowWidth \ 2
    tAPI(2).X = OriginX + ArrowWidth \ 2
    tAPI(3).X = OriginX
    Polygon hDC, tAPI(1), 3
End If

If Sgn(CanvasBorders.P1.Y) <> Sgn(CanvasBorders.P2.Y) Then
    MoveToEx hDC, 0, OriginY, lpPoint
    LineTo hDC, PaperScaleWidth, OriginY
    tAPI(1).X = PaperScaleWidth - ArrowLen - 1
    tAPI(2).X = PaperScaleWidth - ArrowLen - 1
    tAPI(3).X = PaperScaleWidth - 1
    tAPI(1).Y = OriginY - ArrowWidth \ 2
    tAPI(2).Y = OriginY + ArrowWidth \ 2
    tAPI(3).Y = OriginY
    Polygon hDC, tAPI(1), 3
End If

SelectObject hDC, hOldBrush
SelectObject hDC, hOldPen
DeleteObject hBrush
DeleteObject hPen

If setShowAxesMarks Then
    LF.lfWidth = 0
    LF.lfEscapement = 0
    LF.lfOrientation = 0
    LF.lfWeight = 400
    LF.lfItalic = 0
    LF.lfUnderline = 0
    LF.lfStrikeOut = 0
    LF.lfCharSet = 0
    LF.lfOutPrecision = 0
    LF.lfClipPrecision = 0
    LF.lfQuality = 2
    LF.lfPitchAndFamily = 0
    For Z = 1 To Len(tS)
        LF.lfFaceName(Z - 1) = Asc(Mid$(tS, Z, 1))
    Next
    LF.lfFaceName(Len(tS)) = 0
    'LF.lfFaceName = tS & vbNullChar
    LF.lfHeight = AxesFontSize * -20 / Screen.TwipsPerPixelY
    
    hFont = CreateFontIndirect(LF)
    hOldFont = SelectObject(hDC, hFont)
    SetTextColor hDC, Col
    
    If Sgn(CanvasBorders.P1.Y) <> Sgn(CanvasBorders.P2.Y) Then
        SetTextAlign hDC, TA_TOPCENTER
        TI = -21
        For X = Round(CanvasBorders.P1.X) To Round(CanvasBorders.P2.X)
            tX = X
            tY = 0
            ToPhysical tX, tY
            If tX > TI + 20 And X <> 0 Then
                TextOut hDC, tX, tY, CStr(X), Len(CStr(X))
                TI = tX
            End If
        Next
    End If
    
    If Sgn(CanvasBorders.P1.X) <> Sgn(CanvasBorders.P2.X) Then
        SetTextAlign hDC, TA_TOPRIGHT
        TI = 100000
        For Y = Round(CanvasBorders.P1.Y) To Round(CanvasBorders.P2.Y)
            tX = 0
            tY = Y
            ToPhysical tX, tY
            If tY < TI - 20 Then
                TextOut hDC, tX, tY, CStr(Y), Len(CStr(Y))
                TI = tY
            End If
        Next
    End If
    
    SetTextAlign hDC, TA_BOTTOMCENTER
    SelectObject hDC, hOldFont
    DeleteObject hFont
End If
End Sub

Public Sub ShowGrid(ByVal hDC As Long, Optional ByVal ShowGrayed As Boolean = False)
Dim hPen As Long, hOldPen As Long, Col As Long

Col = EnsureRGB(nGridColor)
If ShowGrayed Then Col = GetGrayedColor(Col)

hPen = CreatePen(PS_SOLID, 1, nGridColor)
hOldPen = SelectObject(hDC, hPen)
PolyPolyline hDC, Pts(1), PtNums(1), UBoundPtNums
SelectObject hDC, hOldPen
DeleteObject hPen
End Sub

Public Sub PaperCls(Optional ByVal hDC As Long = -1)
Dim hBrush As Long, lpRect As RECT
If hDC = -1 Then hDC = Paper.hDC
hBrush = CreateSolidBrush(nPaperColor1)
lpRect.Right = PaperScaleWidth
lpRect.Bottom = PaperScaleHeight
FillRect hDC, lpRect, hBrush
DeleteObject hBrush
End Sub

Public Sub ShowLabel(ByVal hDC As Long, ByVal Index As Long, Optional ByVal bShow As Boolean = True, Optional ByVal ShouldRefresh As Boolean = False, Optional ByVal ShowAnyway As Boolean = False, Optional ByVal ShowGrayed As Boolean = False)
Const FontABCOffset = 10
Dim lpRect As RECT, X As Double, Y As Double
Dim hFont As Long, hOldFont As Long, hBrush As Long, Col As Long

If Index = 0 Then Exit Sub
If TextLabels(Index).Hide And Not ShowAnyway Then Exit Sub

With TextLabels(Index)
    If bShow Then
        X = .Position.Left
        Y = .Position.Top
        lpRect.Left = X - 2
        lpRect.Top = Y
        lpRect.Right = .Position.Right + X
        lpRect.Bottom = .Position.Bottom + Y
        
        If .Shadow Then DrawSelectionShadow hDC, lpRect, .ForeColor
'        If Not .Transparent Then
'            hBrush = CreateSolidBrush(.BackColor)
'            FillRect hDC, lpRect, hBrush
'            DeleteObject hBrush
'        End If
        lpRect.Left = X
        
        hFont = CreateFont(.FontSize * -20 / Screen.TwipsPerPixelY, _
            0, 0, 0, IIf(.FontBold, 700, 400), -CLng(.FontItalic), -CLng(.FontUnderline), 0, .Charset, _
            0, 0, 2, 0, .FontName & vbNullChar)
        hOldFont = SelectObject(hDC, hFont)
        SetTextAlign hDC, TA_LEFTTOP
        Col = .ForeColor
        If ShowGrayed Then Col = GetGrayedColor(Col)
        SetTextColor hDC, Col
        
        DrawText hDC, .DisplayName, .LenDisplayName, lpRect, DT_NOPREFIX Or DT_NOCLIP Or DT_EXPANDTABS 'Or DT_EXTERNALLEADING
        SelectObject hDC, hOldFont
        DeleteObject hFont
        'SetTextAlign hDC, TA_BOTTOM Or TA_CENTER
    Else
        hBrush = CreateSolidBrush(setcolPaper1)
        lpRect.Left = .Position.Left - FontABCOffset
        lpRect.Top = .Position.Top - FontABCOffset
        lpRect.Right = .Position.Right + .Position.Left + FontABCOffset
        lpRect.Bottom = .Position.Bottom + .Position.Top + FontABCOffset
        FillRect hDC, lpRect, hBrush
        DeleteObject hBrush
    End If
End With

If ShouldRefresh Then Paper.Refresh
End Sub

'===========================================================================

Public Sub ShowButton(ByVal hDC As Long, ByVal Index As Long, Optional ByVal bShow As Boolean = True, Optional ByVal ShouldRefresh As Boolean = False, Optional ByVal Pushed As Boolean = False, Optional ByVal Selected As Boolean = False, Optional ByVal ShowAnyway As Boolean = False, Optional ByVal ShowGrayed As Boolean = False)
Const FontABCOffset = 10

Dim lpRect As RECT, X As Double, Y As Double
Dim hFont As Long, hOldFont As Long, hBrush As Long, hOldBrush As Long, Col As Long
Dim hPen As Long, hOldPen As Long

If Index = 0 Or Buttons(Index).Hide Then Exit Sub

With Buttons(Index)
    If bShow Then
        X = .Position.Left
        Y = .Position.Top
        lpRect.Left = X - FrameWidth - 2
        lpRect.Top = Y - FrameHeight
        lpRect.Right = .Position.Right + X + FrameWidth + 1
        lpRect.Bottom = .Position.Bottom + Y + FrameHeight + 1
        If (.Appearance And 1) Or Selected Then DrawSelectionShadow hDC, lpRect, .ForeColor
        
        Col = .BackColor
        If ShowGrayed Then Col = GetGrayedColor(Col)
        If CBool(.CurrentState) And .Type = butShowHide Then
            If setGradientFill Then
                'GetGrayedColor(Col)
                Gradient hDC, nPaperColor1, Col, lpRect.Left, lpRect.Top, lpRect.Right - lpRect.Left, lpRect.Bottom - lpRect.Top, False
            Else
                hBrush = CreateSolidBrush(GetSystemColor(SystemColorConstants.vbInfoBackground))
                FillRect hDC, lpRect, hBrush
                DeleteObject hBrush
            End If
        Else
            hBrush = CreateSolidBrush(Col)
            FillRect hDC, lpRect, hBrush
            DeleteObject hBrush
        End If
        
        If ShowGrayed Then
            hBrush = CreateSolidBrush(GetGrayedColor(vbButtonShadow))
            FrameRect hDC, lpRect, hBrush
            DeleteObject hBrush
        Else
            lpRect.Bottom = lpRect.Bottom + 1
            DrawEdge hDC, lpRect, IIf(.Pushed, EDGE_SUNKEN, EDGE_RAISED), BF_RECT
        End If
        
        lpRect.Left = X
        lpRect.Top = lpRect.Top + FrameHeight
        lpRect.Bottom = lpRect.Bottom - FrameHeight - 1
        lpRect.Right = lpRect.Right - FrameWidth - 1
        If .Pushed Then OffsetRect lpRect, 1, 1
        
        hFont = CreateFont(.FontSize * -20 / Screen.TwipsPerPixelY, _
            0, 0, 0, IIf(.FontBold, 700, 400), -CLng(.FontItalic), -CLng(.FontUnderline), 0, .Charset, _
            0, 0, 2, 0, .FontName & vbNullChar)
        hOldFont = SelectObject(hDC, hFont)
        SetTextAlign hDC, TA_LEFTTOP
        Col = .ForeColor
        If ShowGrayed Then Col = GetGrayedColor(Col)
        If .Pushed And Not setGradientFill Then Col = GetSystemColor(SystemColorConstants.vbInfoText)
        SetTextColor hDC, Col
        
        DrawText hDC, .Caption, Len(.Caption), lpRect, DT_NOPREFIX Or DT_NOCLIP Or DT_EXPANDTABS 'Or DT_EXTERNALLEADING
        SelectObject hDC, hOldFont
        DeleteObject hFont
        
        'SetTextAlign hDC, TA_BOTTOM Or TA_CENTER
    Else
        hBrush = CreateSolidBrush(Paper.BackColor)
        lpRect.Left = .Position.Left - FontABCOffset
        lpRect.Top = .Position.Top - FontABCOffset
        lpRect.Right = .Position.Right + .Position.Left + FontABCOffset + 8
        lpRect.Bottom = .Position.Bottom + .Position.Top + FontABCOffset + 8
        FillRect hDC, lpRect, hBrush
        DeleteObject hBrush
    End If
End With

If ShouldRefresh Then Paper.Refresh
End Sub

Public Sub ShowSelectedPoint(ByVal hDC As Long, ByVal Index As Long, Optional ByVal ShouldRefresh As Boolean = False, Optional ByVal Hide As Boolean = False, Optional ByVal ShowGrayed As Boolean = False, Optional ByVal ShowAnyway As Boolean = False)
Dim Z As Long
For Z = 2 To 0 Step -1
    BasePoint(Index).PhysicalWidth = BasePoint(Index).PhysicalWidth + 2 * Z
    ShowPoint hDC, Index, Hide, , ShowAnyway, ShowGrayed, 2
    BasePoint(Index).PhysicalWidth = BasePoint(Index).PhysicalWidth - 2 * Z
Next
If ShouldRefresh Then Paper.Refresh
End Sub

Public Sub ShowSelectedFigure(ByVal hDC As Long, ByVal Index As Long, Optional ByVal ShouldRefresh As Boolean = False, Optional ByVal ShowAnyway As Boolean = False, Optional ByVal ShowGrayed As Boolean = False)
Dim Z As Long
Dim R1 As Long, G1 As Long, B1 As Long
Dim R2 As Long, G2 As Long, B2 As Long
Dim C1 As Long, C2 As Long
Dim OldDrawWidth As Long, OldForeColor As Long, OldDrawMode As Long, OldDrawStyle As Long
Dim K As Double
Dim SelColor As Long: SelColor = RGB(128, 128, 128)

With Figures(Index)
    If IsVisual(Index) Then
        OldDrawWidth = .DrawWidth
        OldForeColor = .ForeColor
        OldDrawMode = .DrawMode
        OldDrawStyle = .DrawStyle
        
        C1 = GetGrayedColor(.ForeColor) 'GetGrayedColor 'SelColor
        C2 = EnsureRGB(nPaperColor1)
'        FillRGB C1, R1, G1, B1
'        FillRGB C2, R2, G2, B2
'
'        .DrawMode = IIf(IsColorDark(nPaperColor1), 15, 9)
'
'        For Z = SelectionThickness To 1 Step -1
'            .DrawWidth = OldDrawWidth + Z
'            K = (Z / (SelectionThickness + 1)) 'Sqr(Sqr
'            .ForeColor = RGB(R1 + K * (R2 - R1), G1 + K * (G2 - G1), B1 + K * (B2 - B1))
'            DrawFigure hDC, Index, False
'        Next
'

        .DrawWidth = OldDrawWidth + SelectionThickness
        .ForeColor = C1
        '.DrawMode = IIf(IsColorDark(nPaperColor1), 15, 9)
        DrawFigure hDC, Index, False, ShowAnyway, ShowGrayed
        
        .DrawWidth = OldDrawWidth
        .ForeColor = OldForeColor
        
        .DrawStyle = OldDrawStyle
        .DrawMode = OldDrawMode
        DrawFigure hDC, Index, False, ShowAnyway, ShowGrayed
        
        .DrawStyle = vbDot
        .DrawMode = vbInvert
        DrawFigure hDC, Index, False, ShowAnyway, ShowGrayed
        .DrawStyle = OldDrawStyle
        .DrawMode = OldDrawMode
        
'        Figures(Index).ForeColor = Figures(Index).ForeColor Xor XorConst
'        Figures(Index).DrawWidth = Figures(Index).DrawWidth + SelectionThickness
'        If Hide Then HideFigure Index, False Else DrawFigure hDC, Index
'        Figures(Index).ForeColor = Figures(Index).ForeColor Xor XorConst
'        Figures(Index).DrawWidth = Figures(Index).DrawWidth - SelectionThickness
'        If Hide Then HideFigure Index, False Else DrawFigure hDC, Index
    Else
        If Figures(Index).FigureType = dsMeasureDistance Or Figures(Index).FigureType = dsMeasureAngle Then
            SetBkMode hDC, OPAQUE
            SetBkColor hDC, GetGrayedColor(Figures(Index).ForeColor)
            DrawFigure hDC, Index
            SetBkMode Paper.hDC, Transparent
        Else
            For Z = 0 To Figures(Index).NumberOfPoints - 1
                If IsChildPointPos(Figures(Index), Z) Then
                    ShowSelectedPoint hDC, Figures(Index).Points(Z)
                End If
            Next
        End If
    End If
End With

If ShouldRefresh Then Paper.Refresh
End Sub

Public Sub ShowSelectedLabel(ByVal hDC As Long, ByVal Index As Long, Optional ByVal ShouldRefresh As Boolean = True, Optional ByVal Hide As Boolean = False)
Dim X As Double, Y As Double, lpRect As RECT, hBrush As Long

With TextLabels(Index)
    If Not Hide Then
        X = .Position.Left
        Y = .Position.Top
        lpRect.Left = X - 2
        lpRect.Top = Y
        lpRect.Right = .Position.Right + X
        lpRect.Bottom = .Position.Bottom + Y
        DrawSelectionShadow hDC, lpRect, .ForeColor
        
        'hBrush = CreateSolidBrush(GetSysColor(vbButtonFace + SysColorTranslationBase))
        'FillRect hDC, lpRect, hBrush
        'DeleteObject hBrush
    End If
End With
ShowLabel hDC, Index, Not Hide, ShouldRefresh
End Sub

Public Sub DrawArrow(ByVal hDC As Long, ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double)
Const K = 1 / 3
Dim pPoints(1 To 3) As POINTAPI
Dim pPointsExact(0 To 3) As OnePoint
Dim dDist As Double
Dim ArrLength As Double

pPointsExact(1).X = X2
pPointsExact(1).Y = Y2
ArrLength = ArrowLength
ToLogicalLength ArrLength
dDist = Distance(X1, Y1, X2, Y2)
If dDist = 0 Then Exit Sub
dDist = dDist / ArrLength
pPointsExact(0).X = X2 + (X1 - X2) / dDist
pPointsExact(0).Y = Y2 + (Y1 - Y2) / dDist
pPointsExact(2).X = pPointsExact(0).X + (pPointsExact(1).Y - pPointsExact(0).Y) * K
pPointsExact(2).Y = pPointsExact(0).Y + (pPointsExact(0).X - pPointsExact(1).X) * K
pPointsExact(3).X = pPointsExact(0).X + (pPointsExact(0).Y - pPointsExact(1).Y) * K
pPointsExact(3).Y = pPointsExact(0).Y + (pPointsExact(1).X - pPointsExact(0).X) * K

ToPhysical pPointsExact(1).X, pPointsExact(1).Y
ToPhysical pPointsExact(2).X, pPointsExact(2).Y
ToPhysical pPointsExact(3).X, pPointsExact(3).Y

pPoints(1).X = pPointsExact(2).X
pPoints(1).Y = pPointsExact(2).Y
pPoints(2).X = pPointsExact(1).X
pPoints(2).Y = pPointsExact(1).Y
pPoints(3).X = pPointsExact(3).X
pPoints(3).Y = pPointsExact(3).Y

Polygon hDC, pPoints(1), 3
End Sub

Public Sub DrawPolygon(P() As POINTAPI, Optional ByVal Color As Long = EmptyVar, Optional ByVal FillColor As Long = EmptyVar, Optional ByVal DrawMode As Integer = 13, Optional ByVal ShouldRefresh As Boolean = True)
Dim hPen As Long, hOldPen As Long
Dim hBrush As Long, hOldBrush As Long

If Color = EmptyVar Then Color = setdefcolFigure
If FillColor = EmptyVar Then FillColor = setdefcolFigureFill

If DrawMode <> 13 Then SetROP2 Paper.hDC, DrawMode

hPen = CreatePen(PS_SOLID, 1, Color)
hOldPen = SelectObject(Paper.hDC, hPen)
hBrush = CreateSolidBrush(FillColor)
hOldBrush = SelectObject(Paper.hDC, hBrush)

Polygon Paper.hDC, P(1), UBound(P) - LBound(P) + 1

SelectObject Paper.hDC, hOldBrush
DeleteObject hBrush
SelectObject Paper.hDC, hOldPen
DeleteObject hPen

If DrawMode <> 13 Then SetROP2 Paper.hDC, 13

If ShouldRefresh Then Paper.Refresh
End Sub

Public Sub DrawLine(ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double, Optional ByVal Color As Long = EmptyVar, Optional ByVal DrawMode As Integer = 13, Optional ByVal ShouldRefresh As Boolean = True)
Dim hPen As Long, hOldPen As Long, lpPoint As POINTAPI

If Color = EmptyVar Then Color = setdefcolFigure

ToPhysical X1, Y1
ToPhysical X2, Y2

If DrawMode <> 13 Then SetROP2 Paper.hDC, DrawMode

'Paper.Line (X1, Y1)-(X2, Y2), Color
hPen = CreatePen(PS_SOLID, 1, Color)
hOldPen = SelectObject(Paper.hDC, hPen)
MoveToEx Paper.hDC, X1, Y1, lpPoint
LineTo Paper.hDC, X2, Y2
SelectObject Paper.hDC, hOldPen
DeleteObject hPen

If DrawMode <> 13 Then SetROP2 Paper.hDC, 13

If ShouldRefresh Then Paper.Refresh
End Sub

Public Sub DrawCircle(ByVal hDC As Long, ByVal XC As Double, ByVal YC As Double, ByVal Radius As Double, Optional ByVal A1 As Double = 0, Optional ByVal A2 As Double = 0, Optional ByVal Color As Long = EmptyVar, Optional ByVal DrawMode As Integer = 13, Optional ByVal DrawStyle As Long = 0, Optional ByVal DrawWidth As Long = 1, Optional ByVal FillColor As Long = colPolygonFillColor, Optional ByVal FillStyle As Long = 6)
Dim A1X As Double, A2X As Double, A1Y As Double, A2Y As Double
Dim hPen As Long, hOldPen As Long, hBrush As Long, hOldBrush As Long

If Color = EmptyVar Then Color = setdefcolFigure
If DrawMode <> 13 Then SetROP2 hDC, DrawMode

hPen = CreatePen(DrawStyle, DrawWidth, Color)
hOldPen = SelectObject(hDC, hPen)

If FillStyle = -1 Then
    hBrush = CreateSolidBrush(FillColor)
ElseIf FillStyle = 6 Then
    hBrush = GetStockObject(NULL_BRUSH)
Else
    hBrush = CreateHatchBrush(FillStyle, FillColor)
End If
hOldBrush = SelectObject(hDC, hBrush)

If A1 = 0 And A2 = 0 Then
    Ellipse hDC, XC - Radius, YC - Radius, XC + Radius, YC + Radius
Else
    A1X = XC + 2 * Radius * Cos(A1)
    A1Y = YC + 2 * Radius * Sin(A1)
    A2X = XC + 2 * Radius * Cos(A2)
    A2Y = YC + 2 * Radius * Sin(A2)
    
    If FillStyle = 6 Then
        Arc hDC, XC - Radius, YC - Radius, XC + Radius, YC + Radius, A1X, A1Y, A2X, A2Y
    Else
        Pie hDC, XC - Radius, YC - Radius, XC + Radius, YC + Radius, A1X, A1Y, A2X, A2Y
    End If
    
End If

SelectObject hDC, hOldBrush
If FillStyle <> 6 Then DeleteObject hBrush

SelectObject hDC, hOldPen
DeleteObject hPen
If DrawMode <> 13 Then SetROP2 hDC, vbCopyPen
End Sub

Public Sub DrawPoint(ByVal X As Double, ByVal Y As Double, Optional ByVal Color As Long = -1, Optional ByVal DrawMode As Long = 13, Optional ByVal Shape As Integer = defPointShape, Optional ByVal Size As Integer = 0, Optional ByVal DrawWidth As Integer)
If Size = 0 Then Size = defPointSize
If Color = -1 Then Color = setdefcolPoint
If DrawMode <> 13 Then Paper.DrawMode = DrawMode
Paper.ForeColor = Color

ToPhysical X, Y
Size = Size / 2
X = Int(X)
Y = Int(Y)

If Shape = vbShapeSquare Then Rectangle Paper.hDC, X - Size, Y - Size, X + Size + 1, Y + Size + 1 Else Ellipse Paper.hDC, X - Size, Y - Size, X + Size + 1, Y + Size + 1

If DrawMode <> 13 Then Paper.DrawMode = vbCopyPen
End Sub

Public Function DrawLinearObjectList(ByVal hDC As Long, L As LinearObjectList, Optional ByVal UptilTo As Long = 0, Optional ByVal ShouldClear As Boolean = True, Optional ByVal ShouldRefresh As Boolean = True)
Dim Z As Long

If UptilTo = 0 Then UptilTo = L.Current

If ShouldClear Then PaperCls hDC
If setWallpaper <> "" Then DrawWallPaper hDC Else If nGradientPaper Then Gradient hDC, nPaperColor1, nPaperColor2, 0, 0, PaperScaleWidth, PaperScaleHeight, False ' Else PaperCls

If nShowGrid Then ShowGrid hDC
If nShowAxes Then ShowAxes hDC

For Z = 1 To UptilTo
    If L.Items(Z).Type = gotSG Then
        With StaticGraphics(L.Items(Z).Index)
            DrawStaticGraphic hDC, L.Items(Z).Index, , , .Hide, True
        End With
    End If
Next

For Z = 1 To UptilTo
    If L.Items(Z).Type = gotLabel Then
        ShowLabel hDC, L.Items(Z).Index, , , True, TextLabels(L.Items(Z).Index).Hide
    End If
Next

For Z = 1 To UptilTo
    If L.Items(Z).Type = gotFigure Then
        If Figures(L.Items(Z).Index).FigureType = dsDynamicLocus Then
            ShowLocus hDC, BasePoint(Figures(L.Items(Z).Index).Points(0)).Locus, , , True, Locuses(BasePoint(Figures(L.Items(Z).Index).Points(0)).Locus).Hide
        End If
    End If
Next

For Z = 1 To UptilTo
    If L.Items(Z).Type = gotFigure Then
        If IsVisual(Figures(L.Items(Z).Index).FigureType) Then
            DrawFigure hDC, L.Items(Z).Index, , True, Figures(L.Items(Z).Index).Hide
        End If
    End If
Next

For Z = 1 To UptilTo
    If L.Items(Z).Type = gotFigure Then
        If Not IsVisual(Figures(L.Items(Z).Index).FigureType) Then
            DrawFigure hDC, L.Items(Z).Index, , True, Figures(L.Items(Z).Index).Hide
        End If
    End If
Next

For Z = 1 To UptilTo
    If L.Items(Z).Type = gotPoint Then
        ShowPoint hDC, L.Items(Z).Index, , , True, BasePoint(L.Items(Z).Index).Hide
    End If
Next

For Z = 1 To UptilTo
    If L.Items(Z).Type = gotButton Then
        ShowButton hDC, L.Items(Z).Index, , , , , True, Buttons(L.Items(Z).Index).Hide
    End If
Next

If ShouldRefresh Then Paper.Refresh

End Function

Public Sub PrintMeasure(ByVal hDC As Long, ByVal szS As String, Optional ByVal X As Long = EmptyVar, Optional ByVal Y As Long = EmptyVar, Optional ByVal Ang As Single = 0)
Dim LF As LOGFONT, i As Long, NF As Long, lpPoint As POINTAPI, Q As Long

If X = EmptyVar Or Y = EmptyVar Then
    GetCurrentPositionEx hDC, lpPoint
    X = lpPoint.X
    Y = lpPoint.Y
End If

LF.lfWidth = 0
LF.lfEscapement = CLng(Ang * 10)
LF.lfOrientation = LF.lfEscapement
LF.lfWeight = 400
LF.lfItalic = 0
LF.lfUnderline = 0
LF.lfStrikeOut = 0
LF.lfCharSet = 0
LF.lfOutPrecision = 0
LF.lfClipPrecision = 0
LF.lfQuality = 2
LF.lfPitchAndFamily = 0
'LF.lfFaceName = ByteFontName
For Q = 0 To 31
    LF.lfFaceName(Q) = ByteFontName(Q)
Next
LF.lfHeight = defSLabelFontSize * -20 / Screen.TwipsPerPixelY

i = CreateFontIndirect(LF)
NF = SelectObject(hDC, i)

TextOut hDC, X, Y, szS, Len(szS)

SelectObject hDC, NF
DeleteObject i
End Sub

Public Sub RenderHighQualityDynamicLoci()
Dim Z As Long
If setLocusDetailsHigh < setLocusDetails + 10 Then Exit Sub
For Z = 0 To FigureCount - 1
    If Figures(Z).FigureType = dsDynamicLocus Then
        RecalcDynamicLocus Z, True
    End If
Next
PaperCls
ShowAll
End Sub

Public Sub RenderHighQuality()
On Local Error Resume Next
RenderHighQualityDynamicLoci
'AntiAliasHDC Paper.hDC, 0, 0, PaperScaleWidth - 1, PaperScaleHeight - 1
Paper.Refresh
End Sub
