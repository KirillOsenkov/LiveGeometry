Public Class CommandToolRay
	Inherits CommandToolBase

	Protected Overrides Sub OnClickCore()
		ParentDocument.Behaviour = New RayCreator(ParentDocument)
	End Sub

End Class
