Friend Class FigureTypeAttribute
	Inherits System.Attribute

	Public Sub New(ByVal FigureName As String, ByVal FigureCategory As String, ByVal CategoriesOfFigureParents As String())
		Name = FigureName
		Category = FigureCategory
		CategoriesOfParents = CategoriesOfFigureParents
	End Sub

	Private mName As String
	Public Property Name() As String
		Get
			Return mName
		End Get
		Set(ByVal Value As String)
			mName = Value
		End Set
	End Property

	Private mCategory As String
	Public Property Category() As String
		Get
			Return mCategory
		End Get
		Set(ByVal Value As String)
			mCategory = Value
		End Set
	End Property

	Private mCategoriesOfParents As String()
	Public Property CategoriesOfParents() As String()
		Get
			Return mCategoriesOfParents
		End Get
		Set(ByVal value As String())
			mCategoriesOfParents = value
		End Set
	End Property

End Class
