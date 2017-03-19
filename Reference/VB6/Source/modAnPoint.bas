Attribute VB_Name = "modAnPoint"
Option Explicit

Public Sub AddAnPoint(ByVal XS As String, ByVal YS As String)
On Local Error GoTo EH:
Dim XTree As Tree, YTree As Tree, NP As Long, nPoint As Long, Z As Long, Q As Long, i As Long

If XS = "" Or YS = "" Then Exit Sub
If IsNumeric(XS) And IsNumeric(YS) Then
    AddBasePoint CDbl(XS), CDbl(YS)
    BasePoint(PointCount).ShowName = True
    BasePoint(PointCount).ShowCoordinates = True
    Exit Sub
End If

XTree = BuildTree(XS)
If WasThereAnErrorEvaluatingLastExpression <> eetEverythingOK Then
    If IsSevere(WasThereAnErrorEvaluatingLastExpression) Or XTree.Erroneous Then
        MsgBox GetString(ResError) & " (" & GetString(ResFigureBase + 2 * dsAnPoint) & ")" & vbCrLf & ERR.Description & vbCrLf & GetString(ResEvalErrorBase + 2 * WasThereAnErrorEvaluatingLastExpression), vbExclamation
        GoTo EH
    End If
    WasThereAnErrorEvaluatingLastExpression = eetEverythingOK
End If
YTree = BuildTree(YS)
If WasThereAnErrorEvaluatingLastExpression <> eetEverythingOK Then
    If IsSevere(WasThereAnErrorEvaluatingLastExpression) Or XTree.Erroneous Then
        MsgBox GetString(ResError) & " (" & GetString(ResFigureBase + 2 * dsAnPoint) & ")" & vbCrLf & ERR.Description & vbCrLf & GetString(ResEvalErrorBase + 2 * WasThereAnErrorEvaluatingLastExpression), vbExclamation
        GoTo EH
    End If
    WasThereAnErrorEvaluatingLastExpression = eetEverythingOK
End If

RedimPreserveFigures 0, FigureCount
AddFigureAux FigureCount, dsAnPoint

NP = 1 + XTree.NumberOfPoints + YTree.NumberOfPoints
Figures(FigureCount).NumberOfPoints = NP
ReDim Figures(FigureCount).Points(0 To NP - 1)
Figures(FigureCount).NumberOfChildren = 0

AddBasePoint XTree.Branches(1).CurrentValue, YTree.Branches(1).CurrentValue, , dsAnPoint, False, FigureCount, False, False

BasePoint(PointCount).ShowName = True
BasePoint(PointCount).ShowCoordinates = True

Figures(FigureCount).XTree = XTree
Figures(FigureCount).YTree = YTree
Figures(FigureCount).XS = RestoreExpressionFromTree(XTree)
Figures(FigureCount).YS = RestoreExpressionFromTree(YTree)

Figures(FigureCount).Points(0) = PointCount

For Z = 1 To XTree.NumberOfPoints
    Figures(FigureCount).Points(Z) = XTree.Points(Z)
Next
For Z = 1 To YTree.NumberOfPoints
    Figures(FigureCount).Points(Z + XTree.NumberOfPoints) = YTree.Points(Z)
Next

Z = 2
Do While Z <= Figures(FigureCount).NumberOfPoints - 1
    For Q = 1 To Z - 1
        If Figures(FigureCount).Points(Q) = Figures(FigureCount).Points(Z) Then
            If Z < Figures(FigureCount).NumberOfPoints - 1 Then
                For i = Z To Figures(FigureCount).NumberOfPoints - 2
                    Figures(FigureCount).Points(i) = Figures(FigureCount).Points(i + 1)
                Next
            End If
            Figures(FigureCount).NumberOfPoints = Figures(FigureCount).NumberOfPoints - 1
            ReDim Preserve Figures(FigureCount).Points(0 To Figures(FigureCount).NumberOfPoints - 1)
            Z = Z - 1
        End If
    Next Q
    Z = Z + 1
Loop

Figures(FigureCount).Description = GetObjectDescription(gotFigure, FigureCount)
BasePoint(PointCount).Description = GetObjectDescription(gotPoint, PointCount)

FigureCount = FigureCount + 1
RecalcAuxInfo FigureCount - 1

PaperCls
ShowAll
Exit Sub
EH:
End Sub

Public Sub OffsetAnPointDependencies(ByVal OldValue As Long, ByVal NewValue As Long)
Dim Z As Long

For Z = 0 To FigureCount - 1
    If Figures(Z).FigureType = dsAnPoint Then
        ReplacePointInTree Figures(Z).XTree, OldValue, NewValue
        ReplacePointInTree Figures(Z).YTree, OldValue, NewValue
    End If
Next
End Sub

Public Sub OffsetAnPointFigureDependencies(ByVal Figure1 As Long, ByVal OldValue As Long, ByVal NewValue As Long)
Dim Z As Long
If Figures(Figure1).FigureType = dsAnPoint Then
    For Z = 1 To Figures(Figure1).AuxInfo(1)
        If Figures(Figure1).AuxArray(Z) = OldValue Then Figures(Figure1).AuxArray(Z) = NewValue
    Next
End If
End Sub
