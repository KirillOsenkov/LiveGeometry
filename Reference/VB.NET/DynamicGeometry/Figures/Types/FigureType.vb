Imports System.Collections.Generic

Friend Class FigureType
	Implements IFigureType

	Public Sub New(ByVal MyName As String)
		Name = MyName
	End Sub

	Public Sub New(ByVal MyName As String, ByVal MyCategory As String)
		Name = MyName
		Category = MyCategory
	End Sub

	Private mParents As IList(Of String)
	Public Property Parents() As IList(Of String) Implements IFigureType.Parents
		Get
			Return mParents
		End Get
		Set(ByVal Value As IList(Of String))
			mParents = Value
		End Set
	End Property

	Private mName As String
	Public Property Name() As String Implements IFigureType.Name
		Get
			Return mName
		End Get
		Set(ByVal Value As String)
			mName = Value
		End Set
	End Property

	Private mCategory As String
	Public Property Category() As String Implements IFigureType.Category
		Get
			Return mCategory
		End Get
		Set(ByVal Value As String)
			mCategory = Value
		End Set
	End Property

End Class
