Friend Class MidPointCreator

	Inherits FigureCreator

	Public Sub New(ByVal ParentDoc As DGDocument)
		MyBase.New(ParentDoc)
	End Sub

	Public Overrides Sub InitializeFigureType()
		NewType = DGObject.GetFigureType(GetType(DGMidPoint))
	End Sub

	' instantiates the object being created by this tool
	Public Overrides Function CreateFigure() As IFigure
		Dim NewFigure As IFigure = New DGMidPoint(Doc, DirectCast(NewParents(0), IDGPoint), DirectCast(NewParents(1), IDGPoint), Doc.Paper.ActiveCoordinateSystem)
		Return NewFigure
	End Function

End Class
