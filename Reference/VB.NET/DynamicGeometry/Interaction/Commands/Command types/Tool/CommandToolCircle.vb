Public Class CommandToolCircle

	Inherits CommandToolBase

	Protected Overrides Sub OnClickCore()
		ParentDocument.Behaviour = New CircleCreator(ParentDocument)
	End Sub

End Class