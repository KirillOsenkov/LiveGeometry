Imports GuiLabs.Canvas

Friend Class PointPair

    Implements IPointPair

    Private point1 As Point
    Private point2 As Point

    Public Property p1() As Point Implements DynamicGeometry.IPointPair.p1
        Get
            Return point1
        End Get
        Set(ByVal Value As Point)
            point1 = Value
        End Set
    End Property

    Public Property p2() As Point Implements DynamicGeometry.IPointPair.p2
        Get
            Return point2
        End Get
        Set(ByVal Value As Point)
            point2 = Value
        End Set
    End Property

    Public Sub New(ByVal NewPoint1 As Point, ByVal NewPoint2 As Point)
        p1 = NewPoint1
        p2 = NewPoint2
    End Sub
End Class
