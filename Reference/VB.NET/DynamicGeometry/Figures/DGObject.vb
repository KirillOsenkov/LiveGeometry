Imports System.Collections.Generic
Imports GuiLabs.Canvas.Utils

Friend MustInherit Class DGObject

	Implements IFigure

	Public Sub New(ByVal ContainerDoc As DGDocument)
		Document = ContainerDoc
	End Sub

#Region "FigureType"

	Private Shared figureTypes As Dict(Of FigureType) = New Dict(Of FigureType)()

	Private mFigureType As IFigureType = Nothing
	Public ReadOnly Property FigureType() As IFigureType Implements IFigure.FigureType
		Get
			If mFigureType IsNot Nothing Then
				Return mFigureType
			End If

			mFigureType = GetFigureType(Me.GetType())
			Return mFigureType
		End Get
	End Property

	Public Shared Function GetFigureType(ByVal figureClass As System.Type) As IFigureType
		Dim attr As FigureTypeAttribute = GetFigureTypeAttribute(figureClass)

		If attr Is Nothing Then
			Throw New Exception("Figure defined incompletely (FigureTypeAttribute missing): " + figureClass.ToString())
		End If

		Dim result As FigureType = figureTypes(attr.Name)
		If result IsNot Nothing Then
			Return result
		End If

		result = New FigureType(attr.Name, attr.Category)
		result.Parents = attr.CategoriesOfParents

		figureTypes.Add(attr.Name, result)
		Return result
	End Function

	Public Shared Function GetFigureTypeAttribute(ByVal figureType As System.Type) As FigureTypeAttribute
		For Each attr As Attribute In figureType.GetCustomAttributes(True)
			If TypeOf (attr) Is FigureTypeAttribute Then
				Return DirectCast(attr, FigureTypeAttribute)
			End If
		Next
		Return Nothing
	End Function

#End Region

	Private WithEvents mDocument As DGDocument
	Public Property Document() As DGDocument
		Get
			Return mDocument
		End Get
		Set(ByVal value As DGDocument)
			mDocument = value
		End Set
	End Property

	Private mParents As IFigureList = New FigureList()
	Public Property Parents() As IFigureList Implements IFigure.Parents
		Get
			Return mParents
		End Get
		Set(ByVal Value As IFigureList)
			mParents = Value
		End Set
	End Property

	Private mChildren As IFigureList = New FigureList()
	Public Property Children() As IFigureList Implements IFigure.Children
		Get
			Return mChildren
		End Get
		Set(ByVal Value As IFigureList)
			mChildren = Value
		End Set
	End Property

	Public Overridable Sub Recalculate() Implements IFigure.Recalculate

	End Sub

End Class