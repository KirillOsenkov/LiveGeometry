Friend Interface IPlane

    Property Width() As Integer
    Property Height() As Integer
    Sub Resize(ByVal NewX As Integer, ByVal NewY As Integer)
    Event Resized(ByVal NewX As Integer, ByVal NewY As Integer)

	ReadOnly Property ActiveCoordinateSystem() As ICoordinateSystem

End Interface
