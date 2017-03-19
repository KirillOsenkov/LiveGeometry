Public Interface ICommandFactory

	'Function GetNewCommand(ByVal CommandName As String) As ICommand
	'Function GetCommandList(ByVal CommandName As String) As CommandList
	Function GetCommand(ByVal CommandName As String) As ICommand
	Sub UpdateCommands()

End Interface
