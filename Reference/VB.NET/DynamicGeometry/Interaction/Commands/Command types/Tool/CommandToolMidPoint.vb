Public Class CommandToolMidPoint
	Inherits CommandToolBase

	Protected Overrides Sub OnClickCore()
		ParentDocument.Behaviour = New MidPointCreator(ParentDocument)
	End Sub

End Class
