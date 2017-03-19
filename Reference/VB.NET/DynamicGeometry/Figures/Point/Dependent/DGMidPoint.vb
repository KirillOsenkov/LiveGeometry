Imports GuiLabs.Canvas.DrawStyle

<FigureType( _
FigureTypeStrings.MidPoint, _
FigureCategoryStrings.Point, _
New String() { _
FigureCategoryStrings.Point, _
FigureCategoryStrings.Point} _
)> _
Friend Class DGMidPoint

	Inherits DGPoint

	Private Shared font As GuiLabs.Canvas.DrawStyle.IFontStyleInfo

	Public Sub New(ByVal ContainerDoc As DGDocument, ByVal NewP1 As IDGPoint, ByVal NewP2 As IDGPoint, ByVal frameOfReference As ICoordinateSystem)
		MyBase.New(ContainerDoc, CType(ContainerDoc.Paper.ActiveCoordinateSystem, CoordinateSystem))
		If font Is Nothing Then
			font = GuiLabs.Canvas.Renderer.RendererSingleton.StyleFactory.ProduceNewFontStyleInfo("Verdana", 11, Drawing.FontStyle.Regular)
		End If

		Parents.Add(NewP1)
		Parents.Add(NewP2)
		Me.FrameOfReference = DirectCast(frameOfReference, CoordinateSystem)
	End Sub

	Public Property P1() As IDGPoint
		Get
			Return DirectCast(Parents(0), IDGPoint)
		End Get
		Set(ByVal Value As IDGPoint)
			Parents(0) = Value
		End Set
	End Property

	Public Property P2() As IDGPoint
		Get
			Return DirectCast(Parents(1), IDGPoint)
		End Get
		Set(ByVal Value As IDGPoint)
			Parents(1) = Value
		End Set
	End Property

	Public Overrides Sub Recalculate()
		SetLogicalCoordinates( _
		  (P1.Coordinates.UnitsX + P2.Coordinates.UnitsX) / 2, _
		  (P1.Coordinates.UnitsY + P2.Coordinates.UnitsY) / 2)
	End Sub

	'Public Overrides Sub Draw(ByVal CurrentRenderer As GuiLabs.Canvas.Renderer.IRenderer)
	'	MyBase.Draw(CurrentRenderer)
	'	CurrentRenderer.DrawOperations.DrawString(Coordinates.ToLogicalString(), New GuiLabs.Canvas.Rect(Coordinates.X, Coordinates.Y, 200, 100), font)
	'End Sub

End Class
