' A "smart" List that also manages dependencies between Figures
Friend Class GeometryObjectList

	Protected mList As FigureList = New FigureList()

	'==============================================================================
	' Public Add* methods - adding new elements to the list
	'==============================================================================

	' Adds an already existing and fully pre-initialized figure to the list
	' All parents of this figure must point to the existing figures already in the list

	Public Sub AddFigure(ByVal ExistingFigure As IFigure)
		' AddFigure adds the figure to the list...
		mList.Add(ExistingFigure)

		' and notifies all parents of the ExistingFigure
		Dim Parent As IFigure
		For Each Parent In ExistingFigure.Parents
			' that they have a new child
			Parent.Children.Add(ExistingFigure)
		Next

		ExistingFigure.Recalculate()

	End Sub

	Public Sub RemoveFigure(ByVal ExistingFigure As IFigure)

		' first notify parents that this figure is leaving
		Dim Parent As IFigure
		For Each Parent In ExistingFigure.Parents
			Parent.Children.Remove(ExistingFigure)
		Next

		' and then delete it from the list
		mList.Remove(ExistingFigure)
	End Sub

	'==============================================================================
	' AddNew* methods - maybe we don't need them, because creating a figure should be done
	' outside this list. This list should only integrate already existing figures.
	' But these functions are convenient...
	'==============================================================================

	'Public Function AddNewPoint(ByVal x As Integer, ByVal y As Integer, ByVal FrameOfReference As ICoordinateSystem) As IDGPoint
	'	Dim NewPoint As IDGPoint = Factory.CreateBasePoint(x, y, FrameOfReference)
	'	AddFigure(NewPoint)
	'	Return NewPoint
	'End Function

	'Public Function AddNewSegment(ByVal EndPoint1 As IDGPoint, ByVal EndPoint2 As IDGPoint) As IDGLine
	'	Dim NewSegment As IDGLine = Factory.CreateSegment(EndPoint1, EndPoint2)
	'	AddFigure(NewSegment)
	'	Return NewSegment
	'End Function



End Class