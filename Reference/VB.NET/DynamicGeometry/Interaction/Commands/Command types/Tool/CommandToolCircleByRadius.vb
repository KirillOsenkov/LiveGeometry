Public Class CommandToolCircleByRadius
	Inherits CommandToolBase

	Protected Overrides Sub OnClickCore()
		ParentDocument.Behaviour = New CircleByRadiusCreator(ParentDocument)
	End Sub

End Class
