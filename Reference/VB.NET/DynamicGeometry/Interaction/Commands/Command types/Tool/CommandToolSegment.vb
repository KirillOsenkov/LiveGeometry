Public Class CommandToolSegment
	Inherits CommandToolBase

	Protected Overrides Sub OnClickCore()
		ParentDocument.Behaviour = New SegmentCreator(ParentDocument)
	End Sub

End Class
