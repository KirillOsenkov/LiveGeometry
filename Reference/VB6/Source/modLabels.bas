Attribute VB_Name = "modLabels"
Option Explicit

Public Sub MoveLabel(ByVal Index As Integer, ByVal X As Double, ByVal Y As Double)
TextLabels(Index).LogicalPosition.P1.X = X
TextLabels(Index).LogicalPosition.P1.Y = Y
RecalcLabel Index
End Sub

Public Function AddTextLabel(Optional ByVal sCaption As String, Optional ByVal Left As Double, Optional ByVal Top As Double) As Boolean
Dim SWidth As Double, sHeight As Double

If LabelCount >= MaxLabelCount Then Exit Function

LabelCount = LabelCount + 1
ReDim Preserve TextLabels(1 To LabelCount)
With TextLabels(LabelCount)
    .Caption = sCaption
    .Charset = setdefLabelCharset
    .DisplayName = sCaption
    .LenDisplayName = Len(sCaption)
    .BackColor = setdefcolLabelBackColor
    .ForeColor = setdefcolLabelColor
    .Transparent = setdefLabelTransparent
    .Visible = True
    .FontName = setdefLabelFont
    .FontBold = setdefLabelBold
    .FontItalic = setdefLabelItalic
    .FontSize = setdefLabelFontSize
    .FontUnderline = setdefLabelUnderline
    .Description = GetObjectDescription(gotLabel, LabelCount)
    .InDemo = True
    .LogicalPosition.P1.X = Left
    .LogicalPosition.P1.Y = Top
End With
ParseLabel LabelCount
GetLabelSize LabelCount

AddTextLabel = True
End Function

Public Function GetLabelFromPoint(ByVal X As Double, ByVal Y As Double) As Long
Dim Z As Long
For Z = LabelCount To 1 Step -1
    If PointInRectangle(X, Y, TextLabels(Z).LogicalPosition.P1.X, TextLabels(Z).LogicalPosition.P1.Y, TextLabels(Z).LogicalPosition.P2.X, TextLabels(Z).LogicalPosition.P2.Y) And Not TextLabels(Z).Hide Then GetLabelFromPoint = Z: Exit Function
Next
End Function

Public Function IsLabel(ByVal Label1 As Long) As Boolean
If Label1 > 0 And Label1 <= LabelCount Then IsLabel = True
End Function

Public Function GetLabelSize(ByVal Index As Long, Optional ByVal Update As Boolean = True) As TwoNumbers
On Local Error Resume Next
Dim SWidth As Double, sHeight As Double

With TextLabels(Index)
    If .FontName <> setdefPointFontName Then Paper.FontName = .FontName
    If .Charset <> Paper.Font.Charset Then Paper.Font.Charset = .Charset
    If .FontSize <> setdefPointFontSize Then Paper.FontSize = .FontSize
    If .FontBold <> setdefPointFontBold Then Paper.FontBold = .FontBold
    If .FontItalic <> setdefPointFontItalic Then Paper.FontItalic = .FontItalic
    If .FontUnderline <> setdefPointFontUnderline Then Paper.FontUnderline = .FontUnderline
    If Not .Transparent Then SetBkMode Paper.hDC, OPAQUE
    
    SWidth = Paper.TextWidth(.DisplayName)
    sHeight = Paper.TextHeight(.DisplayName)
    
    GetLabelSize.n1 = SWidth
    GetLabelSize.n2 = sHeight
    If Update Then
        .Position.Right = SWidth
        .Position.Bottom = sHeight
        RecalcLabel Index
    End If
    
    If .FontName <> setdefPointFontName Then Paper.FontName = setdefPointFontName
    If Paper.Font.Charset <> setdefPointFontCharset Then Paper.Font.Charset = setdefPointFontCharset
    If .FontSize <> setdefPointFontSize Then Paper.FontSize = setdefPointFontSize
    If .FontBold <> setdefPointFontBold Then Paper.FontBold = setdefPointFontBold
    If .FontItalic <> setdefPointFontItalic Then Paper.FontItalic = setdefPointFontItalic
    If .FontUnderline <> setdefPointFontUnderline Then Paper.FontUnderline = setdefPointFontUnderline
    If Not .Transparent Then SetBkMode Paper.hDC, Transparent
End With

'Dim lpSize As Size
'
'With TextLabels(Index)
'    lpSize = GetTextSize(.Caption, .FontName, .FontSize, .FontBold, .FontItalic, .FontUnderline)
'    GetLabelSize.N1 = lpSize.CX
'    GetLabelSize.N2 = lpSize.CY
'    If Update Then
'        .Position.Right = lpSize.CX
'        .Position.Bottom = lpSize.CY
'        RecalcLabel Index
'    End If
'End With
'Dim oldFontName As String, oldFontSize As Integer, oldFontBold As Boolean, oldFontItalic As Boolean, oldFontUnderline As Boolean, OldFontTransparent As Boolean
'Dim sWidth As Double, sHeight As Double
'
'AdjustLabelCharset Index
'
'With TextLabels(Index)
'    If .FontName <> defSLabelFontName Then
'        oldFontName = .FontName
'        Paper.FontName = oldFontName
'        If .Charset <> 0 Then Paper.Font.Charset = .Charset
'    End If
'    If .FontSize <> defSLabelFontSize Then oldFontSize = .FontSize: Paper.FontSize = oldFontSize
'    If .FontBold Then oldFontBold = True: Paper.FontBold = True
'    If .FontItalic Then oldFontItalic = True: Paper.FontItalic = True
'    If .FontUnderline Then oldFontUnderline = True: Paper.FontUnderline = True
'    If Not .Transparent Then SetBkMode Paper.hDC, OPAQUE: OldFontTransparent = True
'    sWidth = Paper.TextWidth(.DisplayName)
'    sHeight = Paper.TextHeight(.DisplayName)
'    GetLabelSize.N1 = sWidth
'    GetLabelSize.N2 = sHeight
'    If Update Then
'        .Position.Right = sWidth
'        .Position.Bottom = sHeight
'        RecalcLabel Index
'    End If
'End With
'
'If oldFontName <> defSLabelFontName Then Paper.FontName = defSLabelFontName: Paper.Font.Charset = DefaultFontCharset
'If oldFontSize <> defSLabelFontSize Then Paper.FontSize = defSLabelFontSize
'If oldFontBold Then Paper.FontBold = False
'If oldFontItalic Then Paper.FontItalic = False
'If OldFontTransparent Then SetBkMode Paper.hDC, Transparent
End Function

Public Sub DeleteLabel(ByVal Index As Long, Optional ByVal ShouldRecord As Boolean = True)
Dim pAction As Action, Z As Long
If Not IsLabel(Index) Then Exit Sub

If ShouldRecord Then
    pAction.Type = actRemoveLabel
    MakeStructureSnapshot pAction
    ReDim pAction.AuxInfo(1 To 1)
    pAction.AuxInfo(1) = Index
    RecordAction pAction
End If

DeleteFromDependentButtons Index, gotLabel, False

PaperCls
If Index < LabelCount Then
    For Z = Index To LabelCount - 1
        TextLabels(Z) = TextLabels(Z + 1)
        OffsetButtonObjectDependencies Z + 1, Z, gotLabel
    Next
End If

LabelCount = LabelCount - 1
If LabelCount > 0 Then ReDim Preserve TextLabels(1 To LabelCount)
ShowAll
Paper.Refresh
End Sub

Public Sub RecalcLabel(ByVal Index As Long)
Dim tX As Double, tY As Double
With TextLabels(Index)
    tX = .LogicalPosition.P1.X
    tY = .LogicalPosition.P1.Y
    ToPhysical tX, tY
    .Position.Left = tX
    .Position.Top = tY
    
    .LogicalPosition.P2.X = .Position.Right
    .LogicalPosition.P2.Y = .Position.Bottom
    ToLogicalLength .LogicalPosition.P2.X
    ToLogicalLength .LogicalPosition.P2.Y
    .LogicalPosition.P2.X = .LogicalPosition.P1.X + .LogicalPosition.P2.X
    .LogicalPosition.P2.Y = .LogicalPosition.P1.Y - .LogicalPosition.P2.Y
End With
End Sub

Public Sub RecalcLabels()
Dim Z As Long
For Z = 1 To LabelCount
    RecalcLabel Z
Next
End Sub

Public Function ParseLabel(ByVal Index As Long)
On Local Error Resume Next
Dim Z As Long
Dim S As String, tS As String, tB As Boolean, CCCP As Long, Z2 As Long, GreatTextSum As String
ReDim TextLabels(Index).CompiledCaptionParts(1 To 1)

S = TextLabels(Index).Caption
Z = InStr(S, "[")
TextLabels(Index).Dynamic = False
If Z = 0 Then
    TextLabels(Index).CompiledCaptionParts(1).Type = DynamicLabelType.StaticString
    TextLabels(Index).CompiledCaptionParts(1).StaticText = S
    TextLabels(Index).DisplayName = S
    TextLabels(Index).CCCP = 1
    Exit Function
End If

Z2 = InStr(Z, S, "]")
If Z2 = 0 Then
    TextLabels(Index).CompiledCaptionParts(1).Type = StaticString
    TextLabels(Index).CompiledCaptionParts(1).StaticText = S
    TextLabels(Index).DisplayName = S
    TextLabels(Index).CCCP = 1
    Exit Function
End If

If Z > 1 Then
    CCCP = 1
    TextLabels(Index).CompiledCaptionParts(CCCP).Type = StaticString
    TextLabels(Index).CompiledCaptionParts(CCCP).StaticText = Left(S, Z - 1)
    GreatTextSum = GreatTextSum & TextLabels(Index).CompiledCaptionParts(CCCP).StaticText
    TextLabels(Index).CCCP = CCCP
End If

tS = Mid(S, Z + 1, Z2 - Z - 1)
If tS <> "" Then
    CCCP = CCCP + 1
    ReDim Preserve TextLabels(Index).CompiledCaptionParts(1 To CCCP)
    tB = TextLabels(Index).Dynamic
    TextLabels(Index).Dynamic = True
    TextLabels(Index).CompiledCaptionParts(CCCP).Type = DynamicString
    TextLabels(Index).CompiledCaptionParts(CCCP).DynamicTree = BuildTree(tS)
    TextLabels(Index).CompiledCaptionParts(CCCP).StaticText = Format(TextLabels(Index).CompiledCaptionParts(CCCP).DynamicTree.Branches(1).CurrentValue, setFormatNumber)
    If IsSevere(WasThereAnErrorEvaluatingLastExpression) Then
        WasThereAnErrorEvaluatingLastExpression = eetEverythingOK
        TextLabels(Index).CompiledCaptionParts(CCCP).Type = StaticString
        TextLabels(Index).CompiledCaptionParts(CCCP).StaticText = "[" & GetString(ResError) & ": " & tS & "]"
        TextLabels(Index).Dynamic = tB
    ElseIf WasThereAnErrorEvaluatingLastExpression <> eetEverythingOK Then
        WasThereAnErrorEvaluatingLastExpression = eetEverythingOK
        TextLabels(Index).CompiledCaptionParts(CCCP).StaticText = "[" & GetString(ResError) & "]"
    End If
    GreatTextSum = GreatTextSum & TextLabels(Index).CompiledCaptionParts(CCCP).StaticText
    TextLabels(Index).CCCP = CCCP
End If

Do
    Z = InStr(Z2, S, "[")
    If Z = 0 Then
        tS = Right(S, Len(S) - Z2)
        If tS <> "" Then
            CCCP = CCCP + 1
            ReDim Preserve TextLabels(Index).CompiledCaptionParts(1 To CCCP)
            TextLabels(Index).CompiledCaptionParts(CCCP).Type = StaticString
            TextLabels(Index).CompiledCaptionParts(CCCP).StaticText = tS
            GreatTextSum = GreatTextSum & TextLabels(Index).CompiledCaptionParts(CCCP).StaticText
        End If
        Exit Do
    End If
    If Z > Z2 + 1 Then
        CCCP = CCCP + 1
        ReDim Preserve TextLabels(Index).CompiledCaptionParts(1 To CCCP)
        TextLabels(Index).CompiledCaptionParts(CCCP).Type = StaticString
        TextLabels(Index).CompiledCaptionParts(CCCP).StaticText = Mid(S, Z2 + 1, Z - Z2 - 1)
        GreatTextSum = GreatTextSum & TextLabels(Index).CompiledCaptionParts(CCCP).StaticText
    End If
    Z2 = InStr(Z, S, "]")
    If Z2 = 0 Then
        tS = Right(S, Len(S) - Z)
        If tS <> "" Then
            CCCP = CCCP + 1
            ReDim Preserve TextLabels(Index).CompiledCaptionParts(1 To CCCP)
            TextLabels(Index).CompiledCaptionParts(CCCP).Type = StaticString
            TextLabels(Index).CompiledCaptionParts(CCCP).StaticText = tS
            GreatTextSum = GreatTextSum & TextLabels(Index).CompiledCaptionParts(CCCP).StaticText
        End If
        Exit Do
    End If
    tS = Mid(S, Z + 1, Z2 - Z - 1)
    If tS <> "" Then
        CCCP = CCCP + 1
        ReDim Preserve TextLabels(Index).CompiledCaptionParts(1 To CCCP)
        tB = TextLabels(Index).Dynamic
        TextLabels(Index).Dynamic = True
        TextLabels(Index).CompiledCaptionParts(CCCP).Type = DynamicString
        TextLabels(Index).CompiledCaptionParts(CCCP).DynamicTree = BuildTree(tS)
        TextLabels(Index).CompiledCaptionParts(CCCP).StaticText = Format(TextLabels(Index).CompiledCaptionParts(CCCP).DynamicTree.Branches(1).CurrentValue, setFormatNumber)
        If IsSevere(WasThereAnErrorEvaluatingLastExpression) Then
            WasThereAnErrorEvaluatingLastExpression = eetEverythingOK
            TextLabels(Index).CompiledCaptionParts(CCCP).Type = StaticString
            TextLabels(Index).CompiledCaptionParts(CCCP).StaticText = "[" & GetString(ResError) & ": " & tS & "]"
            TextLabels(Index).Dynamic = tB
        ElseIf WasThereAnErrorEvaluatingLastExpression <> eetEverythingOK Then
            WasThereAnErrorEvaluatingLastExpression = eetEverythingOK
            TextLabels(Index).CompiledCaptionParts(CCCP).StaticText = "[" & GetString(ResError) & "]"
        End If
        GreatTextSum = GreatTextSum & TextLabels(Index).CompiledCaptionParts(CCCP).StaticText
    End If
Loop

TextLabels(Index).CCCP = CCCP
TextLabels(Index).DisplayName = GreatTextSum
TextLabels(Index).LenDisplayName = Len(GreatTextSum)
End Function

Public Sub RecalculateDynamicLabel(ByVal Index As Long)
On Local Error Resume Next
Dim GreatTextSum As String, Z As Long

With TextLabels(Index)
    For Z = 1 To .CCCP
        If .CompiledCaptionParts(Z).Type = DynamicString Then .CompiledCaptionParts(Z).StaticText = Format(RecalculateTree(.CompiledCaptionParts(Z).DynamicTree), setFormatNumber)
        If WasThereAnErrorEvaluatingLastExpression <> eetEverythingOK Then
            .CompiledCaptionParts(Z).StaticText = "[" & GetString(ResError) & ": " & GetString(ResEvalErrorBase + 2 * WasThereAnErrorEvaluatingLastExpression - 2) & "]"
            WasThereAnErrorEvaluatingLastExpression = eetEverythingOK
        End If
        GreatTextSum = GreatTextSum & .CompiledCaptionParts(Z).StaticText
    Next Z
    .DisplayName = GreatTextSum
    .LenDisplayName = Len(GreatTextSum)
End With
End Sub

Public Sub UpdateLabels(Optional ByVal ShouldHide As Boolean = True)
Dim Z As Long

For Z = 1 To LabelCount
    If TextLabels(Z).Dynamic Then
        'If ShouldHide Then ShowLabel Paper.hDC, Z, False
        RecalculateDynamicLabel Z
        GetLabelSize Z
    End If
Next
End Sub

Public Sub AdjustLabelCharset(ByVal Label1 As Long)
Dim sStr As String, Z As Long
sStr = TextLabels(Label1).DisplayName
TextLabels(Label1).Charset = 0
For Z = 1 To Len(sStr)
    If Asc(Mid(sStr, Z, 1)) > 127 Then TextLabels(Label1).Charset = DefaultFontCharset: Exit Sub
Next
End Sub

Public Function LabelMemoryConsumption(pLabel As TextLabel) As Long
On Local Error Resume Next
Dim MemSum As Long, Z As Long
MemSum = MemSum + Len(pLabel) + Len(pLabel.Caption)
If UBound(pLabel.CompiledCaptionParts) >= LBound(pLabel.CompiledCaptionParts) Then
    For Z = LBound(pLabel.CompiledCaptionParts) To UBound(pLabel.CompiledCaptionParts)
        MemSum = MemSum + DynamicTreeMemoryConsumption(pLabel.CompiledCaptionParts(Z).DynamicTree)
    Next
End If
LabelMemoryConsumption = MemSum
End Function

Public Sub AddLabel()
ActiveLabel = 0
frmLabelProps.Show
End Sub

Public Sub AddLabelJob( _
ByVal Caption As String, _
Optional ByVal Color As Long = 0, _
Optional ByVal TempFontBold As Boolean = False, _
Optional ByVal TempFontItalic As Boolean = False, _
Optional ByVal TempFontUnderline As Boolean = False, _
Optional ByVal TempFontSize As Integer = 12, _
Optional ByVal TempFontName As String = "Arial", _
Optional ByVal TempFontCharset As Long = 0, _
Optional ByVal bFix As Boolean = False)

Dim pAction As Action

If Not AddTextLabel(Caption) Then Exit Sub
With TextLabels(LabelCount)
    .Caption = Caption
    .DisplayName = .Caption
    .LenDisplayName = Len(.Caption)
    .ForeColor = Color
    .FontBold = TempFontBold
    .FontItalic = TempFontItalic
    .FontUnderline = TempFontUnderline
    .FontSize = TempFontSize
    .FontName = TempFontName
    .Charset = TempFontCharset
    .Fixed = bFix
End With
MoveLabel LabelCount, (CanvasBorders.P1.X + CanvasBorders.P2.X) / 2, (CanvasBorders.P1.Y + CanvasBorders.P2.Y) / 2

pAction.Type = actAddLabel
pAction.pLabel = LabelCount
RecordAction pAction

ParseLabel LabelCount
GetLabelSize LabelCount
PaperCls
ShowAll
End Sub
