Imports GuiLabs.Canvas

Friend Interface IDGPoint

    Inherits IDGObject
    Inherits ISnap

    Property Style() As IPointAppearance
    'Property FrameOfReference() As ICoordinateTransformer

    Sub MoveTo(ByVal x As Integer, ByVal y As Integer)
    Sub MoveTo(ByVal newLocation As Point)

End Interface