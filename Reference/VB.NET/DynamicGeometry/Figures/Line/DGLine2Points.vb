<FigureType( _
FigureTypeStrings.Line2Points, _
FigureCategoryStrings.Line, _
New String() { _
FigureCategoryStrings.Point, _
FigureCategoryStrings.Point} _
)> _
Friend Class DGLine2Points

	Inherits DGLine

	Public Sub New(ByVal ContainerDoc As DGDocument, ByVal NewP1 As IDGPoint, ByVal NewP2 As IDGPoint)
		MyBase.New(ContainerDoc, NewP1, NewP2)
		Recalculate()
	End Sub

	Public Overrides Function IsPointOver(ByVal x As Integer, ByVal y As Integer) As Boolean

	End Function

	Public Overrides Sub Recalculate()
		Dim bounds As MathTwoPoints = Document.Paper.ActiveCoordinateSystem.Viewport
		'bounds.p1.Add(3)
		'bounds.p2.Add(-3)
		Dim start As MathPoint = P1.Coordinates.GetMathPoint
		Dim finish As MathPoint = P2.Coordinates.GetMathPoint
		Dim result As MathTwoPoints = DGMath.GetLineFromSegment(start, finish, bounds)
		physicalP1.SetLogical(result.p1)
		physicalP2.SetLogical(result.p2)
		Document.Paper.ActiveCoordinateSystem.UpdatePhysicalFromLogical(physicalP1)
		Document.Paper.ActiveCoordinateSystem.UpdatePhysicalFromLogical(physicalP2)
	End Sub

	Private physicalP1 As CartesianPoint = New CartesianPoint()
	Private physicalP2 As CartesianPoint = New CartesianPoint()

	Public Overrides Sub Draw(ByVal CurrentRenderer As GuiLabs.Canvas.Renderer.IRenderer)
		CurrentRenderer.DrawOperations.DrawLine(physicalP1, physicalP2, Me.Style.LineStyle)
	End Sub

End Class
