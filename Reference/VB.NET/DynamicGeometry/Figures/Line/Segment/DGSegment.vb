<FigureType( _
FigureTypeStrings.Segment, _
FigureCategoryStrings.Line, _
New String() { _
FigureCategoryStrings.Point, _
FigureCategoryStrings.Point} _
)> _
Friend Class DGSegment

	Inherits DGLine

	Public Sub New(ByVal ContainerDoc As DGDocument, ByVal NewP1 As IDGPoint, ByVal NewP2 As IDGPoint)
		MyBase.New(ContainerDoc, NewP1, NewP2)
		'Recalculate()
	End Sub

	Public Overrides Function IsPointOver(ByVal x As Integer, ByVal y As Integer) As Boolean

	End Function

End Class
