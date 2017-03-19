Imports System.Reflection

Friend Class FigureLibrary

	'Private mFigureTypes As FigureTypeList
	'Public Property FigureTypes() As FigureTypeList
	'	Get
	'		Return mFigureTypes
	'	End Get
	'	Set(ByVal Value As FigureTypeList)
	'		mFigureTypes = Value
	'	End Set
	'End Property

	'Private mFigureCategories As FigureCategoryList
	'Public Property FigureCategories() As FigureCategoryList
	'	Get
	'		Return mFigureCategories
	'	End Get
	'	Set(ByVal Value As FigureCategoryList)
	'		mFigureCategories = Value
	'	End Set
	'End Property

	Protected Sub New()
		'mFigureTypes = New FigureTypeList()
		'mFigureCategories = New FigureCategoryList()
		ScanAssembly(GetType(FigureLibrary).Assembly)
	End Sub

	Private Sub ScanAssembly(ByVal ExistingAssembly As System.Reflection.Assembly)
		Dim Classes As Type() = ExistingAssembly.GetTypes()
		Dim i As Type
		For Each i In Classes
			Dim FoundAttributes As Object() = i.GetCustomAttributes(GetType(FigureTypeAttribute), False)
			If Not FoundAttributes Is Nothing AndAlso FoundAttributes.Length > 0 Then

				For Each attr As Attribute In FoundAttributes
					If TypeOf (attr) Is FigureTypeAttribute Then
						FoundAttribute(i, DirectCast(attr, FigureTypeAttribute))
					End If
				Next
			End If
		Next
	End Sub

	Private Sub FoundAttribute(ByVal classType As System.Type, ByVal TypeAttribute As FigureTypeAttribute)

	End Sub

	Private Shared mInstance As FigureLibrary
	Public Shared ReadOnly Property Instance() As FigureLibrary
		Get
			If mInstance Is Nothing Then
				mInstance = New FigureLibrary()
			End If
			Return mInstance
		End Get
	End Property

End Class
