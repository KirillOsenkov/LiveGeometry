Public Class CommandModeDrag
	Inherits CommandToolBase

	Protected Overrides Sub OnClickCore()
		ParentDocument.Behaviour = New Dragger(ParentDocument)
	End Sub

End Class
