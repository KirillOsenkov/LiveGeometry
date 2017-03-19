Namespace Actions

	Friend MustInherit Class DGAction
		Inherits AbstractAction

		'====================================================================
		' Properties
		'====================================================================

		Private mDoc As DGDocument
		Protected Property Doc() As DynamicGeometry.DGDocument
			Get
				Return mDoc
			End Get
			Set(ByVal Value As DynamicGeometry.DGDocument)
				mDoc = Value
			End Set
		End Property

	End Class
End Namespace
