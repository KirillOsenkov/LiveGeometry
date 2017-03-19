Imports DynamicGeometry

Public Class ControlFactory

	Public Sub New(ByVal ExistingCmdFactory As ICommandFactory)
		mCmdFactory = ExistingCmdFactory
	End Sub

	Private mCmdFactory As ICommandFactory
	Public ReadOnly Property CmdFactory() As ICommandFactory
		Get
			Return mCmdFactory
		End Get
	End Property

	Public Function CreateMenuItem(ByVal CommandName As String) As System.Windows.Forms.MenuItem
		Dim NewCommand As Command = CmdFactory.GetCommand(CommandName)
		Dim NewMenuItem As DGMenuItem = New DGMenuItem(NewCommand)
		Return NewMenuItem
	End Function

	Public Function CreateToolBarButton(ByVal CommandName As String) As System.Windows.Forms.ToolBarButton
		Dim NewCommand As Command = CmdFactory.GetCommand(CommandName)
		Dim NewButton As DGToolBarButton = New DGToolBarButton(NewCommand)
		NewButton.Style = ToolBarButtonStyle.ToggleButton
		NewButton.ToolTipText = NewButton.UICommand.Tooltip
		Return NewButton
	End Function

	Public Function CreateToolBarSeparator() As System.Windows.Forms.ToolBarButton
		Dim NewButton As New System.Windows.Forms.ToolBarButton
		NewButton.Style = ToolBarButtonStyle.Separator
		Return NewButton
	End Function

End Class
