Imports System.Windows.Forms

Public Class DGToolbarBuilder

	Implements IToolbarBuilder

	Public Function BuildToolbar(ByVal AppControlFactory As ControlFactory) As System.Windows.Forms.ToolBarButton() Implements IToolbarBuilder.BuildToolbar
		Dim tbbGeometrySep1 As ToolBarButton = AppControlFactory.CreateToolBarSeparator()
		Dim tbbGeometrySep2 As ToolBarButton = AppControlFactory.CreateToolBarSeparator()
		Dim tbbGeometrySep3 As ToolBarButton = AppControlFactory.CreateToolBarSeparator()

		Dim tbbPointer As ToolBarButton = AppControlFactory.CreateToolBarButton(DynamicGeometry.CommandStrings.Mode.Drag)
		Dim tbbPoint As ToolBarButton = AppControlFactory.CreateToolBarButton(DynamicGeometry.CommandStrings.Tool.Point)
		Dim tbbSegment As ToolBarButton = AppControlFactory.CreateToolBarButton(DynamicGeometry.CommandStrings.Tool.Segment)
		Dim tbbRay As ToolBarButton = AppControlFactory.CreateToolBarButton(DynamicGeometry.CommandStrings.Tool.Ray)
		Dim tbbLine As ToolBarButton = AppControlFactory.CreateToolBarButton(DynamicGeometry.CommandStrings.Tool.Line)
		Dim tbbMidPoint As ToolBarButton = AppControlFactory.CreateToolBarButton(DynamicGeometry.CommandStrings.Tool.MidPoint)
		Dim tbbCircle As System.Windows.Forms.ToolBarButton = AppControlFactory.CreateToolBarButton(DynamicGeometry.CommandStrings.Tool.Circle)
		Dim tbbCircleByRadius As System.Windows.Forms.ToolBarButton = AppControlFactory.CreateToolBarButton(DynamicGeometry.CommandStrings.Tool.CircleByRadius)

		tbbPointer.ImageIndex = 0
		tbbPoint.ImageIndex = 1
		tbbSegment.ImageIndex = 2
		tbbRay.ImageIndex = 3
		tbbLine.ImageIndex = 4
		tbbMidPoint.ImageIndex = 11
		tbbCircle.ImageIndex = 8
		tbbCircleByRadius.ImageIndex = 9

		Dim tbbArray As System.Windows.Forms.ToolBarButton() = _
		{tbbPointer, _
		tbbPoint, _
		tbbGeometrySep1, _
		tbbSegment, _
		tbbRay, _
		tbbLine, _
		tbbGeometrySep2, _
		tbbCircle, _
		tbbCircleByRadius, _
		tbbGeometrySep3, _
		tbbMidPoint}

		Return tbbArray
	End Function

End Class
