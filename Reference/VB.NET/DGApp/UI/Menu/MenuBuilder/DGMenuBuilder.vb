Public Class DGMenuBuilder
	Implements IMenuBuilder

	Public Function BuildMenu(ByVal Factory As ControlFactory) As System.Windows.Forms.MainMenu Implements DGApp.IMenuBuilder.BuildMenu
		Dim MainMenu1 As System.Windows.Forms.MainMenu = New System.Windows.Forms.MainMenu()

		Dim mnuEdit As System.Windows.Forms.MenuItem = New System.Windows.Forms.MenuItem()
		Dim mnuUndo As System.Windows.Forms.MenuItem = Factory.CreateMenuItem(DynamicGeometry.CommandStrings.Edit.Undo)
		Dim mnuRedo As System.Windows.Forms.MenuItem = Factory.CreateMenuItem(DynamicGeometry.CommandStrings.Edit.Redo)

		mnuEdit.Text = "Edit"
		mnuEdit.Index = 1

		mnuUndo.Text = "Undo"

		mnuRedo.Text = "Redo"

		mnuEdit.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {mnuUndo, mnuRedo})
		MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {mnuEdit})

		Return MainMenu1
	End Function
End Class
