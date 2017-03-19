Friend Class CircleByRadiusCreator
    Inherits FigureCreator

	Public Sub New(ByVal ParentDoc As DGDocument)
		MyBase.New(ParentDoc)
	End Sub

	Public Overrides Sub InitializeFigureType()
		NewType = DGObject.GetFigureType(GetType(DGCircleByRadius))
	End Sub

	' instantiates the object being created by this tool
	Public Overrides Function CreateFigure() As IFigure
		Dim NewFigure As IFigure = New DGCircleByRadius(Me.Doc, DirectCast(NewParents(0), IDGPoint), DirectCast(NewParents(1), IDGPoint), DirectCast(NewParents(2), IDGPoint))
		Return NewFigure
    End Function

    Public Overrides Sub Draw(ByVal Renderer As GuiLabs.Canvas.Renderer.IRenderer)
        If CurrentParent = 1 Then
            Dim startPoint As DGPoint = DirectCast(NewParents(0), DGPoint)
            Dim endPoint As IDGPoint = TempPoint

            If startPoint IsNot Nothing And endPoint IsNot Nothing Then
                Renderer.DrawOperations.DrawLine(startPoint.Coordinates, endPoint.Coordinates, AppearanceFactory.Instance.DefaultLineAppearance.LineStyle)
            End If
        End If

        MyBase.Draw(Renderer)
    End Sub
End Class
