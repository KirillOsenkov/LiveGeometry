Public Class AppMenuBuilder

	Implements IMenuBuilder

	Public Function BuildMenu(ByVal MIFactory As ControlFactory) As System.Windows.Forms.MainMenu Implements IMenuBuilder.BuildMenu

		Dim mnuMDIParent As System.Windows.Forms.MainMenu = New System.Windows.Forms.MainMenu()

		Dim mnuFile As System.Windows.Forms.MenuItem = MIFactory.CreateMenuItem(CommandStrings.File)
		Dim mnuFileNew As System.Windows.Forms.MenuItem = MIFactory.CreateMenuItem(CommandStrings.FileNew)
		Dim mnuFileClose As System.Windows.Forms.MenuItem = MIFactory.CreateMenuItem(CommandStrings.FileClose)
		Dim mnuFileSep1 As System.Windows.Forms.MenuItem = New System.Windows.Forms.MenuItem()
		Dim mnuFileExit As System.Windows.Forms.MenuItem = MIFactory.CreateMenuItem(CommandStrings.FileExit)

		Dim mnuWindow As System.Windows.Forms.MenuItem = MIFactory.CreateMenuItem(CommandStrings.Window)
		Dim mnuWindowCascade As System.Windows.Forms.MenuItem = MIFactory.CreateMenuItem(CommandStrings.WindowCascade)
		Dim mnuWindowTileHoriz As System.Windows.Forms.MenuItem = MIFactory.CreateMenuItem(CommandStrings.WindowTileHorizontal)
		Dim mnuWindowTileVertic As System.Windows.Forms.MenuItem = MIFactory.CreateMenuItem(CommandStrings.WindowTileVertical)
		Dim mnuWindowArrangeIcons As System.Windows.Forms.MenuItem = MIFactory.CreateMenuItem(CommandStrings.WindowArrange)

		mnuFile.Text = "&File"
		mnuFile.Index = 0

		mnuFileNew.Shortcut = System.Windows.Forms.Shortcut.CtrlN
		mnuFileNew.Text = "&New"
		mnuFileClose.Text = "&Close"
		mnuFileSep1.Text = "-"
		mnuFileExit.Shortcut = System.Windows.Forms.Shortcut.CtrlX
		mnuFileExit.Text = "&Exit"

		mnuWindow.MdiList = True
		mnuWindow.MergeOrder = 2
		mnuWindow.Text = "&Window"
		mnuWindow.Index = 1

		mnuWindowCascade.Text = "&Cascade"
		mnuWindowTileHoriz.Text = "Tile &Horizontally"
		mnuWindowTileVertic.Text = "Tile &Vertically"
		mnuWindowArrangeIcons.Text = "&Arrange icons"

		mnuFile.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {mnuFileNew, mnuFileClose, mnuFileSep1, mnuFileExit})
		mnuWindow.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {mnuWindowCascade, mnuWindowTileHoriz, mnuWindowTileVertic, mnuWindowArrangeIcons})
		mnuMDIParent.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {mnuFile, mnuWindow})

		Return mnuMDIParent

	End Function

End Class
