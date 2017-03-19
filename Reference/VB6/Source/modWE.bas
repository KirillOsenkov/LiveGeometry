Attribute VB_Name = "modWE"
Option Explicit

Public Sub WEInputNew()
On Local Error GoTo EH
Dim InputStr As String, EqvIns As Long
Dim tName As String, tExpression As String, nNewWE As Long

InputStr = Trim(InputBox(GetString(ResEnterExpression), , "", , , App.HelpFile, ResHlp_Interface_Expressions))
If InputStr = "" Then Exit Sub
EqvIns = InStr(InputStr, "=")

If EqvIns <> 0 Then
    tName = Left(InputStr, EqvIns - 1)
    tExpression = Right(InputStr, Len(InputStr) - EqvIns)
Else
    tName = InputStr
    tExpression = InputStr
End If

nNewWE = WEAdd(tName, tExpression)

EH:
End Sub

Public Function WEAdd(ByRef sName As String, ByVal sValue As String, Optional ByVal ShouldRecord As Boolean = True) As Long
Dim pAction As Action, OldWECount As Long
pAction.Type = actAddWE
pAction.pWE = WECount + 1

OldWECount = WECount
WECount = WECount + 1
ReDim Preserve WatchExpressions(1 To WECount)

Do While GetWatchExpressionByName(sName) <> 0
    sName = sName & Format(WECount)
Loop

WatchExpressions(WECount).Name = sName
WatchExpressions(WECount).Expression = sValue
WatchExpressions(WECount).WatchTree = BuildTree(sValue)

If WatchExpressions(WECount).WatchTree.Erroneous Then
    If IsSevere(WatchExpressions(WECount).WatchTree.Error) Then
        WECount = OldWECount
        If WECount > 0 Then ReDim Preserve WatchExpressions(1 To WECount)
        MsgBox GetString(ResError) & ": " & vbCrLf & GetString(ResEvalErrorBase + WasThereAnErrorEvaluatingLastExpression * 2 - 2) & "."
        WasThereAnErrorEvaluatingLastExpression = eetEverythingOK
        WEAdd = 0
        Exit Function
    Else
        WatchExpressions(WECount).Value = EmptyVar
        WatchExpressions(WECount).WatchTree.Error = eetEverythingOK
        WatchExpressions(WECount).WatchTree.Erroneous = False
        FormMain.ValueTable1.WEInsert sName, GetString(ResError) & ": " & GetString(ResEvalErrorBase + WasThereAnErrorEvaluatingLastExpression * 2 - 2)
    End If
Else
    WatchExpressions(WECount).Value = WatchExpressions(WECount).WatchTree.Branches(1).CurrentValue
    FormMain.ValueTable1.WEInsert sName, Round(WatchExpressions(WECount).Value, setNumberPrecision)
End If

RebuildWatchTrees

WEAdd = WECount
If ShouldRecord Then RecordAction pAction
Exit Function

EH:
ReportError "Ошибка при добавлении динамического выражения:" & vbCrLf & ERR.Description
WECount = OldWECount
If WECount > 0 Then ReDim Preserve WatchExpressions(1 To WECount)
WEAdd = 0
End Function

Public Function WERemove(ByVal Index As Long)
On Local Error GoTo EH:
Dim pAction As Action, Z As Long
If WECount = 0 Then Exit Function

pAction.Type = actRemoveWE
MakeStructureSnapshot pAction

Z = FigureCount - 1
Do While Z >= 0
    If Figures(Z).FigureType = dsAnPoint Then
        If TreeDependsOnWE(Index, Figures(Z).XTree) Or TreeDependsOnWE(Index, Figures(Z).YTree) Then
            DeleteFigure Z, False, False
        End If
    End If
    Z = Z - 1
Loop

Z = Index + 1
Do While Z <= WECount
    If TreeDependsOnWE(Index, WatchExpressions(Z).WatchTree) Then
        WERemove Z
    Else
        Z = Z + 1
    End If
Loop

If Index < WECount Then
    For Z = Index To WECount - 1
        WatchExpressions(Z) = WatchExpressions(Z + 1)
        OffsetWEDependencies Z + 1, Z
    Next
End If
WECount = WECount - 1
If WECount > 0 Then
    ReDim Preserve WatchExpressions(1 To WECount)
Else
    ReDim WatchExpressions(1 To 1)
End If
FormMain.ValueTable1.RemoveExpression Index, True
RecordAction pAction
PaperCls
ShowAll
EH:
End Function

Public Function WEEdit(ByVal Index As Long)

End Function

Public Function WEClear()

End Function

Public Function WERefresh()

End Function

Public Function AddWatchExpression(ByRef exName As String, ByVal exExpression As String, Optional ByVal ShouldRecord As Boolean = True) As Boolean
Dim pAction As Action
pAction.Type = actAddWE
pAction.pWE = WECount + 1

WECount = WECount + 1
ReDim Preserve WatchExpressions(1 To WECount)

Do While GetWatchExpressionByName(exName) <> 0
    exName = exName & Format(WECount)
Loop

WatchExpressions(WECount).Name = exName
WatchExpressions(WECount).Expression = exExpression
WatchExpressions(WECount).WatchTree = BuildTree(exExpression)
If WasThereAnErrorEvaluatingLastExpression <> eetEverythingOK Then
    WECount = WECount - 1
    If WECount > 0 Then ReDim Preserve WatchExpressions(1 To WECount)
    MsgBox GetString(ResError) & ": " & vbCrLf & GetString(ResEvalErrorBase + WasThereAnErrorEvaluatingLastExpression * 2 - 2) & "."
    WasThereAnErrorEvaluatingLastExpression = eetEverythingOK
    AddWatchExpression = False
    Exit Function
End If
WatchExpressions(WECount).Value = WatchExpressions(WECount).WatchTree.Branches(1).CurrentValue

RebuildWatchTrees


AddWatchExpression = True
If ShouldRecord Then RecordAction pAction
Exit Function

EH:
WECount = WECount - 1
If WECount > 0 Then ReDim Preserve WatchExpressions(1 To WECount)
'?????
ReportError "Ошибка при добавлении динамического выражения:" & vbCrLf & ERR.Description
End Function

Public Sub RemoveWatchExpression(ByVal Index As Long, Optional ByVal ShouldUpdateExpressions As Boolean = True, Optional ByVal ShouldRefresh As Boolean = True)
On Local Error GoTo EH:
Dim pAction As Action, Z As Long
If WECount = 0 Then Exit Sub

pAction.Type = actRemoveWE
MakeStructureSnapshot pAction

Z = FigureCount - 1
Do While Z >= 0
    If Figures(Z).FigureType = dsAnPoint Then
        If TreeDependsOnWE(Index, Figures(Z).XTree) Or TreeDependsOnWE(Index, Figures(Z).YTree) Then
            DeleteFigure Z, False, False
        End If
    End If
    Z = Z - 1
Loop

Z = Index + 1
Do While Z <= WECount
    If TreeDependsOnWE(Index, WatchExpressions(Z).WatchTree) Then
        RemoveWatchExpression Z, False
    Else
        Z = Z + 1
    End If
Loop

If Index < WECount Then
    For Z = Index To WECount - 1
        WatchExpressions(Z) = WatchExpressions(Z + 1)
        OffsetWEDependencies Z + 1, Z
    Next
End If
WECount = WECount - 1
If WECount > 0 Then
    ReDim Preserve WatchExpressions(1 To WECount)
Else
    ReDim WatchExpressions(1 To 1)
End If
FormMain.ValueTable1.RemoveExpression Index, ShouldUpdateExpressions
If ShouldUpdateExpressions Then RecordAction pAction
If ShouldRefresh Then
    PaperCls
    ShowAll
End If
EH:
End Sub

Public Function EditWatchExpression(ByVal Index As Long, ByVal exName As String, ByVal exExpression As String) As Boolean
On Local Error GoTo EH:
Dim Tree1 As Tree
If exName = "" Or exExpression = "" Then Exit Function
Tree1 = BuildTree(exExpression)
If WasThereAnErrorEvaluatingLastExpression <> eetEverythingOK Then
    MsgBox GetString(ResError) & ": " & vbCrLf & GetString(ResEvalErrorBase + WasThereAnErrorEvaluatingLastExpression * 2 - 2) & "."
    WasThereAnErrorEvaluatingLastExpression = eetEverythingOK
    EditWatchExpression = False
    Exit Function
End If
WatchExpressions(Index).WatchTree = Tree1
WatchExpressions(Index).Name = exName
WatchExpressions(Index).Expression = exExpression
RebuildWatchTrees
RecalcAllAuxInfo
PaperCls
ShowAll
EditWatchExpression = True
EH:
End Function

Public Sub RebuildWatchTrees()
Dim Z As Long

For Z = 1 To WECount
    WatchExpressions(Z).WatchTree = BuildTree(WatchExpressions(Z).Expression)
    If WasThereAnErrorEvaluatingLastExpression <> eetEverythingOK Then
        MsgBox GetString(ResError) & ": " & vbCrLf & GetString(ResEvalErrorBase + WasThereAnErrorEvaluatingLastExpression * 2 - 2) & "."
        WasThereAnErrorEvaluatingLastExpression = eetEverythingOK
        Exit Sub
    End If
    WatchExpressions(Z).Value = WatchExpressions(Z).WatchTree.Branches(1).CurrentValue
Next
End Sub

Public Function GetWatchExpressionByName(ByVal WName As String) As Long
Dim Z As Long
For Z = 1 To WECount
    If WName = WatchExpressions(Z).Name Then GetWatchExpressionByName = Z: Exit Function
Next
End Function

Public Sub OffsetWEDependencies(Optional ByVal Offset As Long = 1, Optional ByVal Bound As Long = 0)
Dim Z As Long, Q As Long
For Q = 1 To WECount
    ReplaceWEInTree WatchExpressions(Q).WatchTree, Offset, Bound
Next
For Z = 0 To FigureCount - 1
    If Figures(Z).FigureType = dsAnPoint Then
        ReplaceWEInTree Figures(Z).XTree, Offset, Bound
        ReplaceWEInTree Figures(Z).YTree, Offset, Bound
    End If
Next
End Sub

Public Sub OffsetWEPointDependencies(Optional ByVal Offset As Long = 1, Optional ByVal Bound As Long = 0)
Dim Q As Long
For Q = 1 To WECount
    ReplacePointInTree WatchExpressions(Q).WatchTree, Offset, Bound
Next
End Sub

