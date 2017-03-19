Public Class CommandToolPoint
	Inherits CommandToolBase

	Protected Overrides Sub OnClickCore()
		ParentDocument.Behaviour = New BasePointCreator(ParentDocument)
	End Sub

End Class
