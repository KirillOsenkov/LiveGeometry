Attribute VB_Name = "modButtons"
Option Explicit

Public Function AddButton(Optional ByVal BType As ButtonType, Optional ByVal sCaption As String, Optional ByVal Left As Double, Optional ByVal Top As Double) As Boolean
Dim SWidth As Double, sHeight As Double, pAction As Action

If ButtonCount >= MaxButtonCount Then Exit Function

ButtonCount = ButtonCount + 1
ReDim Preserve Buttons(1 To ButtonCount)
With Buttons(ButtonCount)
    .Caption = sCaption
    .Charset = setdefLabelCharset
    .BackColor = GetSysColor(vbButtonFace + SysColorTranslationBase)
    .ForeColor = GetSysColor(vbButtonText + SysColorTranslationBase)
    .Visible = True
    .FontName = setdefLabelFont
    .FontBold = setdefLabelBold
    .FontItalic = setdefLabelItalic
    .FontSize = setdefLabelFontSize
    .FontUnderline = setdefLabelUnderline
    .LogicalPosition.P1.X = Left
    .LogicalPosition.P1.Y = Top
    .RemindToSaveFile = True
    .Type = BType
    .Description = GetObjectDescription(gotButton, ButtonCount)
    .InDemo = True
End With

GetButtonSize ButtonCount
pAction.Type = actAddButton
pAction.pButton = ButtonCount
RecordAction pAction

AddButton = True
End Function

Public Function IsButton(ByVal Index As Long) As Boolean
If Index < 1 Or Index > ButtonCount Then IsButton = False Else IsButton = True
End Function

Public Sub DeleteButton(ByVal Index As Long, Optional ByVal ShouldRecord As Boolean = True)
Dim pAction As Action, Z As Long
If Not IsButton(Index) Then Exit Sub

If ShouldRecord Then
    pAction.Type = actRemoveButton
    MakeStructureSnapshot pAction
    ReDim pAction.AuxInfo(1 To 1)
    pAction.AuxInfo(1) = Index
    RecordAction pAction
End If

If Buttons(Index).Type = butShowHide Then
    ObjectListShowHideAll Buttons(Index).ObjectListAux, True, False
End If

DeleteFromDependentButtons Index, gotButton, False

If Index < ButtonCount Then
    For Z = Index To ButtonCount - 1
        Buttons(Z) = Buttons(Z + 1)
        OffsetButtonObjectDependencies Z + 1, Z, gotButton
    Next
End If

ButtonCount = ButtonCount - 1
If ButtonCount > 0 Then ReDim Preserve Buttons(1 To ButtonCount)
PaperCls
ShowAll
Paper.Refresh
End Sub

Public Sub DeleteFromDependentButtons(ByVal Index As Long, ByVal ObjType As GeometryObjectType, Optional ByVal ShouldRecord As Boolean = True)
Dim Z As Long, Removed As Boolean
If ButtonCount = 0 Then Exit Sub

Z = 1
Do While Z <= ButtonCount
    Removed = ObjectListRemove(Buttons(Z).ObjectListAux, ObjType, Index) <> 0
    If Buttons(Z).ObjectListAux.TotalCount = 0 And Removed Then DeleteButton Z, ShouldRecord Else Z = Z + 1
Loop
End Sub

Public Sub ButtonPushed(ByVal Index As Long)
On Local Error GoTo EH
Dim pAction As Action, S As String, A As String, D As String, F As String

If Index = 0 Then Exit Sub
Select Case Buttons(Index).Type
Case butShowHide
    RecordGenericAction ResUndoButtonPush
    'pAction.Type = actApplyButton
    'MakeStructureSnapshot pAction
    'RecordAction pAction
    
    Buttons(Index).CurrentState = CLng(Not CBool(Buttons(Index).CurrentState))
    ObjectListShowHideAll Buttons(Index).ObjectListAux, CBool(Buttons(Index).CurrentState)
Case butMsgBox
    MsgBox Buttons(Index).Message, vbInformation
Case butPlaySound
    If Dir(Buttons(Index).Path) <> "" Then PlaySound Buttons(Index).Path, SND_ASYNC Or SND_NODEFAULT Else Beep
Case butLaunchFile
    A = RetrieveDir(DrawingName)
    If A = "" Then A = ProgramPath
    
    S = Buttons(Index).Path
    D = RetrieveDir(S)
    F = RetrieveName(S)
    If D = "" Then
        D = A
        S = D & S
    End If
    
    If Not IsValidPath(D) Then S = GetAbsolutePath(S, A)
    
    If Dir(S) <> "" Then
        If Right(UCase(S), 3) = extFIG Then
            If IsDirty And Buttons(Index).RemindToSaveFile Then
                Select Case MsgBox(GetString(ResSave) & " " & DrawingName & "?", vbYesNoCancel + vbQuestion)
                    Case vbYes
                        MenuCommand ResSave
                    Case vbNo
                        'do nothing
                    Case vbCancel
                        Exit Sub
                End Select
            End If
            DrawingName = S
            AddMRUItem DrawingName
            i_ShowStatus GetString(ResWorkingPleaseWait)
            FormMain.Enabled = False
            Screen.MousePointer = vbHourglass
            If Not FormMain.Fullscreen Then FormMain.Caption = GetString(ResCaption) + " - " + RetrieveName(DrawingName)
            DoEvents
            OpenFile DrawingName, True, True, False
            i_ShowStatus
            Screen.MousePointer = vbDefault
            FormMain.Enabled = True
            If FormMain.Enabled And FormMain.Visible Then FormMain.SetFocus
        Else
            ShellExecute 0, "open", S, "", "", 1
        End If
    Else
        MsgBox GetString(ResMsgCannotOpenFile), vbExclamation
    End If
End Select
EH:
End Sub

Public Sub AutorunButtons()
Dim Z As Long
For Z = 1 To ButtonCount
    If Buttons(Z).Type = butShowHide Then ObjectListShowHideAll Buttons(Z).ObjectListAux, Buttons(Z).InitiallyVisible, False
Next
End Sub

Public Function GetButtonFromPoint(ByVal X As Double, ByVal Y As Double) As Long
Dim Z As Long, lpRect As RECT
ToPhysical X, Y
For Z = ButtonCount To 1 Step -1
    lpRect = Buttons(Z).Position
    lpRect.Left = Buttons(Z).Position.Left - FrameWidth - 2
    lpRect.Top = Buttons(Z).Position.Top - FrameHeight
    lpRect.Right = Buttons(Z).Position.Left + Buttons(Z).Position.Right + FrameWidth + 1
    lpRect.Bottom = Buttons(Z).Position.Top + Buttons(Z).Position.Bottom + FrameHeight + 1
    If PointInRectangle(X, Y, lpRect.Left, lpRect.Top, lpRect.Right, lpRect.Bottom) Then GetButtonFromPoint = Z: Exit Function
Next
End Function

Public Sub MoveButton(ByVal Index As Integer, ByVal X As Double, ByVal Y As Double)
Buttons(Index).LogicalPosition.P1.X = X
Buttons(Index).LogicalPosition.P1.Y = Y
RecalcButton Index
End Sub

Public Sub OffsetButtonObjectDependencies(ByVal OldValue As Long, ByVal NewValue As Long, ByVal ObjType As GeometryObjectType)
Dim Z As Long, tZ As Long

For Z = 1 To ButtonCount
    ObjectListReplace Buttons(Z).ObjectListAux, ObjType, OldValue, NewValue
Next
End Sub

Public Sub RecalcButton(ByVal Index As Long)
Dim tX As Double, tY As Double
With Buttons(Index)
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

Public Sub RecalcButtons()
Dim Z As Long
For Z = 1 To ButtonCount
    RecalcButton Z
Next
End Sub

Public Function GetButtonSize(ByVal Index As Long, Optional ByVal Update As Boolean = True) As TwoNumbers
On Local Error Resume Next
Dim SWidth As Double, sHeight As Double

With Buttons(Index)
    If .FontName <> setdefPointFontName Then Paper.FontName = .FontName
    If .Charset <> Paper.Font.Charset Then Paper.Font.Charset = .Charset
    If .FontSize <> setdefPointFontSize Then Paper.FontSize = .FontSize
    If .FontBold <> setdefPointFontBold Then Paper.FontBold = .FontBold
    If .FontItalic <> setdefPointFontItalic Then Paper.FontItalic = .FontItalic
    If .FontUnderline <> setdefPointFontUnderline Then Paper.FontUnderline = .FontUnderline
    
    SWidth = Paper.TextWidth(.Caption)
    sHeight = Paper.TextHeight(.Caption)
    
    GetButtonSize.n1 = SWidth
    GetButtonSize.n2 = sHeight
    If Update Then
        .Position.Right = SWidth
        .Position.Bottom = sHeight
        RecalcButton Index
    End If
    
    If .FontName <> setdefPointFontName Then Paper.FontName = setdefPointFontName
    If Paper.Font.Charset <> setdefPointFontCharset Then Paper.Font.Charset = setdefPointFontCharset
    If .FontSize <> setdefPointFontSize Then Paper.FontSize = setdefPointFontSize
    If .FontBold <> setdefPointFontBold Then Paper.FontBold = setdefPointFontBold
    If .FontItalic <> setdefPointFontItalic Then Paper.FontItalic = setdefPointFontItalic
    If .FontUnderline <> setdefPointFontUnderline Then Paper.FontUnderline = setdefPointFontUnderline
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

