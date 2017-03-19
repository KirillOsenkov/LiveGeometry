Imports GuiLabs.Canvas.DrawStyle

Friend Class PointAppearance

    Implements IPointAppearance

    Private mFillStyle As IFillStyleInfo
    Public Property FillStyle() As IFillStyleInfo Implements DynamicGeometry.IPointAppearance.FillStyle
        Get
            Return mFillStyle
        End Get
        Set(ByVal Value As IFillStyleInfo)
            mFillStyle = Value
        End Set
    End Property

    Private mLineStyle As ILineStyleInfo
    Public Property LineStyle() As ILineStyleInfo Implements DynamicGeometry.IPointAppearance.LineStyle
        Get
            Return mLineStyle
        End Get
        Set(ByVal Value As ILineStyleInfo)
            mLineStyle = Value
        End Set
    End Property

    Private mShape As PointShapeType = PointShapeType.Circle
    Public Property Shape() As PointShapeType Implements IPointAppearance.Shape
        Get
            Return mShape
        End Get
        Set(ByVal Value As PointShapeType)
            mShape = Value
        End Set
    End Property

    Private mWidth As Integer = 3
    Public Property Width() As Integer Implements DynamicGeometry.IPointAppearance.Width
        Get
            Return mWidth
        End Get
        Set(ByVal Value As Integer)
            mWidth = Value
        End Set
    End Property
End Class
