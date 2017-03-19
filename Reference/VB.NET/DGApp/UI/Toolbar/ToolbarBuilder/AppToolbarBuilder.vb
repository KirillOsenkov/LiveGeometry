Public Class AppToolbarBuilder

	Implements IToolbarBuilder

	Public Function BuildToolbar(ByVal AppControlFactory As ControlFactory) As System.Windows.Forms.ToolBarButton() Implements IToolbarBuilder.BuildToolbar
		Dim tbbNew As System.Windows.Forms.ToolBarButton = AppControlFactory.CreateToolBarButton(CommandStrings.FileNew)
		Dim tbbUndo As System.Windows.Forms.ToolBarButton = AppControlFactory.CreateToolBarButton(DynamicGeometry.CommandStrings.Edit.Undo)
		Dim tbbRedo As System.Windows.Forms.ToolBarButton = AppControlFactory.CreateToolBarButton(DynamicGeometry.CommandStrings.Edit.Redo)

		tbbNew.ImageIndex = 0
		tbbUndo.ImageIndex = 3
		tbbRedo.ImageIndex = 4

		Dim tbbArray As System.Windows.Forms.ToolBarButton() = {tbbNew, tbbUndo, tbbRedo}
		Return tbbArray
	End Function


	'Public Function BuildToolbar(ByVal AppControlFactory As ControlFactory) As System.Windows.Forms.ToolBarButton() Implements IToolbarBuilder.BuildToolbar
	'	Dim tbbNew As System.Windows.Forms.ToolBarButton = AppControlFactory.CreateToolBarButton(CommandStrings.FileNew)
	'	Dim tbbUndo As System.Windows.Forms.ToolBarButton = AppControlFactory.CreateToolBarButton(DynamicGeometry.CommandStrings.Edit.Undo)
	'	Dim tbbRedo As System.Windows.Forms.ToolBarButton = AppControlFactory.CreateToolBarButton(DynamicGeometry.CommandStrings.Edit.Redo)

	'	tbbNew.ImageIndex = 0
	'	tbbUndo.ImageIndex = 3
	'	tbbRedo.ImageIndex = 4

	'	Dim tbbArray As System.Windows.Forms.ToolBarButton()
	'	ReDim tbbArray(11)
	'	' = {tbbNew, tbbUndo, tbbRedo}

	'	Dim i As Integer
	'	For i = 0 To 8
	'		tbbArray(i) = AppControlFactory.CreateToolBarButton(DynamicGeometry.CommandStrings.Edit.Undo)
	'		tbbArray(i).ImageIndex = 3
	'	Next

	'	tbbArray(9) = tbbNew
	'	tbbArray(10) = tbbUndo
	'	tbbArray(11) = tbbRedo

	'	Return tbbArray
	'End Function


End Class
