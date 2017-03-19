Friend Class SegmentCreator

	Inherits FigureCreator

	Public Sub New(ByVal ParentDoc As DGDocument)
		MyBase.New(ParentDoc)
	End Sub

	Public Overrides Sub InitializeFigureType()
		NewType = DGObject.GetFigureType(GetType(DGSegment))
	End Sub

	' instantiates the object being created by this tool
	Public Overrides Function CreateFigure() As IFigure
		Dim NewFigure As IFigure = New DGSegment(Doc, DirectCast(NewParents(0), IDGPoint), DirectCast(NewParents(1), IDGPoint))
		Return NewFigure
	End Function

End Class
