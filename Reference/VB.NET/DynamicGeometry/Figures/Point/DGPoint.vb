Imports GuiLabs.Canvas
Imports GuiLabs.Canvas.Renderer

Friend MustInherit Class DGPoint

    Inherits DGObject
    Implements IDGPoint

	Public Sub New(ByVal ContainerDoc As DGDocument, ByVal plane As CoordinateSystem)
		MyBase.New(ContainerDoc)
		Style = AppearanceFactory.Instance.FindPointAppearance("Free point")
		FrameOfReference = plane
	End Sub

#Region " Coordinates "

    Private mCoordinates As CartesianPoint = New CartesianPoint()
    Public ReadOnly Property Coordinates() As CartesianPoint Implements IDGPoint.Coordinates
        Get
            Return mCoordinates
        End Get
    End Property

	Private mFrameOfReference As CoordinateSystem
	Public Property FrameOfReference() As CoordinateSystem
		Get
			Return mFrameOfReference
		End Get
		Set(ByVal value As CoordinateSystem)
			mFrameOfReference = value
		End Set
	End Property

#End Region

#Region " Style "

    Private mStyle As IPointAppearance
    Public Property Style() As IPointAppearance Implements IDGPoint.Style
        Get
            Return mStyle
        End Get
        Set(ByVal Value As IPointAppearance)
            mStyle = Value
			Bounds.Size.Set(2 * Radius + 1)
		End Set
    End Property

#End Region

    Private mBounds As Rect = New Rect(0, 0, 7, 7)
	Public ReadOnly Property Bounds() As Rect
		Get
			Return mBounds
		End Get
		'Set(ByVal Value As Rect)
		'    mBounds = Value
		'End Set
	End Property

    Private ReadOnly Property Radius() As Integer
        Get
            Return Style.Width
		End Get
	End Property

	Public Overridable Sub Draw(ByVal CurrentRenderer As IRenderer) Implements DynamicGeometry.IDGObject.Draw
		CurrentRenderer.DrawOperations.DrawFilledEllipse( _
			Bounds, Style.LineStyle, Style.FillStyle)
	End Sub

	Public Sub SetPhysicalCoordinates(ByVal x As Integer, ByVal y As Integer) Implements IDGPoint.MoveTo
		Coordinates.Set(x, y)
		FrameOfReference.UpdateLogicalFromPhysical(Coordinates)
		UpdateBounds()
	End Sub

	Public Sub SetPhysicalCoordinates(ByVal newLocation As Point) Implements IDGPoint.MoveTo
		SetPhysicalCoordinates(newLocation.X, newLocation.Y)
	End Sub

	Public Sub SetLogicalCoordinates(ByVal x As Double, ByVal y As Double)
		Coordinates.SetLogical(x, y)
		FrameOfReference.UpdatePhysicalFromLogical(Coordinates)
		UpdateBounds()
	End Sub

	Public Sub UpdateBounds()
		Bounds.Location.Set(Coordinates.X - Radius, Coordinates.Y - Radius)
	End Sub

    Public Function IsPointOver(ByVal x As Integer, ByVal y As Integer) As Boolean Implements DynamicGeometry.IDGObject.IsPointOver
        Dim tx As Integer = Math.Abs(x - Coordinates.X)
        Dim ty As Integer = Math.Abs(y - Coordinates.Y)
        If tx > ty Then ty = tx

        Return ty <= Radius + CursorSensitivity
    End Function

End Class