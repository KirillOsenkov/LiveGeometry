Friend Class FigureCategoryList
	Implements System.Collections.IEnumerable

	Private list As ArrayList = New ArrayList()

	Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
		Return list.GetEnumerator
	End Function

	Public Sub Add(ByVal NewType As String)
		list.Add(NewType)
	End Sub

	Public Function Count() As Integer
		Return list.Count
	End Function

	Default Public Property Item(ByVal Index As Integer) As String
		Get
			Return DirectCast(list.Item(Index), String)
		End Get
		Set(ByVal Value As String)
			list.Item(Index) = Value
		End Set
	End Property

End Class
