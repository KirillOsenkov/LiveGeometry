Public Class CommandToolLine
	Inherits CommandToolBase

	Protected Overrides Sub OnClickCore()
		ParentDocument.Behaviour = New LineCreator(ParentDocument)
	End Sub

End Class
