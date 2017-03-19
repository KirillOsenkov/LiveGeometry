Imports System.Collections.Generic

Friend Class FigureListNavigator

	Private visitedSet As Dictionary(Of IFigure, Integer) = Nothing
	Private finalOrder As LinkedList(Of IFigure) = Nothing

	Public Function GetAllDependentsSorted(ByVal Root As IFigure) As IFigureList
		Return GetAllDependentsSorted(New IFigure() {Root})
	End Function

	Public Function GetAllDependentsSorted(ByVal Roots As IEnumerable(Of IFigure)) As IFigureList
		visitedSet = New Dictionary(Of IFigure, Integer)
		finalOrder = New LinkedList(Of IFigure)

		For Each root As IFigure In Roots
			AddAllDependents(root)
		Next

		Return New FigureList(finalOrder)
	End Function

	Private Sub AddAllDependents(ByVal Root As IFigure)
		If visitedSet.ContainsKey(Root) Then Return

		For Each child As IFigure In Root.Children
			AddAllDependents(child)
		Next

		visitedSet(Root) = 1

		finalOrder.AddFirst(Root)
	End Sub

End Class
