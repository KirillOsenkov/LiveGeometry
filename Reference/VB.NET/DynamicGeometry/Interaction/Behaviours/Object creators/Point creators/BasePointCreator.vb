Imports System.Windows.Forms
Imports GuiLabs.Canvas.Events

Friend Class BasePointCreator

	Inherits FigureCreator

	Public Sub New(ByVal ParentDoc As DGDocument)
		MyBase.New(ParentDoc)
	End Sub

	Public Overrides Function CreateFigure() As DynamicGeometry.IFigure
        Return Nothing
	End Function

	Public Overrides Sub InitializeFigureType()
		NewType = DGObject.GetFigureType(GetType(DGBasePoint))
	End Sub

	Public Overrides Sub Reset()
		MyBase.Reset()
	End Sub

	'====================================================================

    Public Overrides Sub OnMouseDown(ByVal e As MouseEventArgsWithKeys)
        If e.IsRightButtonPressed Then
            AbortAndSetDefaultTool()
            Return
        End If

		Dim NewPoint As IDGPoint = New DGBasePoint(Doc, CInt(e.X), CInt(e.Y))
        Doc.ActionManager.AddFigure(NewPoint)
        Doc.RaiseNeedRedraw()
    End Sub

    Public Overrides Sub OnMouseMove(ByVal e As MouseEventArgsWithKeys)

    End Sub

	'====================================================================

	'TODO: TEST: stress: 1000 points
    Public Overrides Sub OnDoubleClick(ByVal e As MouseEventArgsWithKeys)
        Dim i As Integer
        For i = 1 To 1000
            ' Dim NewPoint As IDGPoint = New DGBasePoint(CInt(Rnd() * Me.Doc.View.Canvas.ClientSize.Width), CInt(Rnd() * Me.Doc.View.Canvas.ClientSize.Height), Doc.Paper.ActiveCoordinateSystem)
            'Doc.ActionManager.AddFigure(NewPoint)
        Next
        Doc.RaiseNeedRedraw()
    End Sub

End Class