Imports System.Collections.Generic

' List of IFigure objects
Friend Class FigureList
	Inherits List(Of IFigure)
	Implements IFigureList

	Public Sub New()
		MyBase.New()
	End Sub

	Public Sub New(ByVal copyFrom As IEnumerable(Of IFigure))
		MyBase.New(copyFrom)
	End Sub

	Public Function FindFirstFigure(ByVal Category As String) As IFigure Implements IFigureList.FindFirstFigure
		Dim i As IFigure
		For Each i In Me
			If i.FigureType.Category = Category Then
				Return i
			End If
		Next
		Return Nothing
	End Function

	Public Sub Recalculate() Implements IFigureList.Recalculate
		For Each figure As IFigure In Me
			figure.Recalculate()
		Next
	End Sub

	'Private list As ArrayList = New ArrayList()

	'Default Public Property Item(ByVal Index As Integer) As IFigure Implements IFigureList.Item
	'	Get
	'		Return DirectCast(list.Item(Index), IFigure)
	'	End Get
	'	Set(ByVal Value As IFigure)
	'		list.Item(Index) = Value
	'	End Set
	'End Property

	'Public Sub Add(ByVal ExistingFigure As IFigure) Implements IFigureList.Add
	'	list.Add(ExistingFigure)
	'End Sub

	'Public Sub Remove(ByVal ExistingFigure As IFigure) Implements IFigureList.Remove
	'	list.Remove(ExistingFigure)
	'End Sub

	'Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
	'	Return list.GetEnumerator
	'End Function

	'Public Sub Clear() Implements IFigureList.Clear
	'	list.Clear()
	'End Sub

End Class
