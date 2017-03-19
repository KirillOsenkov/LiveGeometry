Imports GuiLabs.Canvas
Imports GuiLabs.Canvas.Renderer

<FigureType( _
FigureTypeStrings.Circle, _
FigureCategoryStrings.Circle, _
New String() { _
FigureCategoryStrings.Point, _
FigureCategoryStrings.Point} _
)> _
Friend Class DGCircle
	Inherits DGObject
	Implements IDGCircle

	Public Sub New(ByVal ContainerDoc As DGDocument, ByVal NewP1 As IDGPoint, ByVal NewP2 As IDGPoint)
		MyBase.New(ContainerDoc)
		Parents.Add(NewP1)
		Parents.Add(NewP2)
		Me.Style = AppearanceFactory.Instance.FindLineAppearance("Line")
	End Sub

	Private mStyle As ILineAppearance
	Public Property Style() As ILineAppearance
		Get
			Return mStyle
		End Get
		Set(ByVal Value As ILineAppearance)
			mStyle = Value
		End Set
	End Property

	Public Function IsPointOver(ByVal x As Integer, ByVal y As Integer) As Boolean Implements DynamicGeometry.IDGCircle.IsPointOver
		Return False
	End Function

	Public Sub Draw(ByVal UseRenderer As IRenderer) Implements DynamicGeometry.IDGCircle.Draw

		Dim t As CoordinateSystem = DirectCast(Center, DGPoint).FrameOfReference

		Dim x0d As Double = Center.Coordinates.UnitsX
		Dim y0d As Double = Center.Coordinates.UnitsY
		Dim rd As Double = Radius
		Dim x1d As Double = x0d - rd
		Dim y1d As Double = y0d + rd

		t.ToPhysical(x1d, y1d)

		Dim x0 As Integer = CInt(x1d)
		Dim y0 As Integer = CInt(y1d)
		Dim r As Long = t.UnitsToPixels(2 * rd)


		Dim Rect As Rect = New Rect(x0, y0, CInt(r) + 1, CInt(r) + 1)

		UseRenderer.DrawOperations.DrawEllipse(Rect, Style.LineStyle)
	End Sub

	Public Function PixelRadius() As Integer
		Return DGMath.PixelDistance(P1.Coordinates, Center.Coordinates)
	End Function

	Public Property Radius() As Double Implements DynamicGeometry.IDGCircle.Radius
		Get
			Return DGMath.Distance(P1.Coordinates, Center.Coordinates)
		End Get
		Set(ByVal Value As Double)
			'do nothing
		End Set
	End Property

	Public Property P1() As IDGPoint
		Get
			Return DirectCast(Parents(1), IDGPoint)
		End Get
		Set(ByVal Value As IDGPoint)
			Parents(1) = Value
		End Set
	End Property

	Public Property Center() As DynamicGeometry.IDGPoint Implements DynamicGeometry.IDGCircle.Center
		Get
			Return DirectCast(Parents(0), IDGPoint)
		End Get
		Set(ByVal Value As DynamicGeometry.IDGPoint)
			Parents(0) = Value
		End Set
	End Property

End Class
