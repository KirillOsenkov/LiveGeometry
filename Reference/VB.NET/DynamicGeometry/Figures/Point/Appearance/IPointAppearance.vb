Imports GuiLabs.Canvas.DrawStyle

Friend Enum PointShapeType
    Circle
    Square
    Diamond
End Enum

Friend Interface IPointAppearance

    Property LineStyle() As ILineStyleInfo
    Property FillStyle() As IFillStyleInfo

    Property Shape() As PointShapeType
    Property Width() As Integer

End Interface
