Imports DynamicGeometry

Public Class DGAppCommandFactory

	Inherits CommandFactory

	Friend Sub New(ByVal ExistingForm As frmParent)
		mParentForm = ExistingForm
	End Sub

	Private mParentForm As frmParent = Nothing
	Friend ReadOnly Property ParentForm() As frmParent
		Get
			Return mParentForm
		End Get
	End Property

	Protected Overrides Function CreateNewCommand(ByVal CommandName As String) As Command
		Dim NewCommand As DGAppCommand

		Select Case CommandName

			Case CommandStrings.File
				NewCommand = New CommandMainMenu()

			Case CommandStrings.FileNew
				NewCommand = New CommandFileNew()

			Case CommandStrings.FileClose
				NewCommand = New CommandFileClose()

			Case CommandStrings.FileExit
				NewCommand = New CommandFileExit()


			Case CommandStrings.Window
				NewCommand = New CommandMainMenu(False)

			Case CommandStrings.WindowCascade
				NewCommand = New CommandWindowLayout(MdiLayout.Cascade)

			Case CommandStrings.WindowTileHorizontal
				NewCommand = New CommandWindowLayout(MdiLayout.TileHorizontal)

			Case CommandStrings.WindowTileVertical
				NewCommand = New CommandWindowLayout(MdiLayout.TileVertical)

			Case CommandStrings.WindowArrange
				NewCommand = New CommandWindowLayout(MdiLayout.ArrangeIcons)

			Case Else
				NewCommand = New DGCommandProxy(CommandName)

		End Select

		NewCommand.ParentForm = ParentForm

		Return NewCommand
	End Function

End Class
