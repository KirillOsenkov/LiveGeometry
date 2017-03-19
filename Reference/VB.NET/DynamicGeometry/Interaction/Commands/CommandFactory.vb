Imports System.Collections.Generic
Imports GuiLabs.Canvas.Utils

Public MustInherit Class CommandFactory

	Implements ICommandFactory

	Public Function GetCommand(ByVal CommandName As String) As ICommand Implements ICommandFactory.GetCommand
		Dim NewCommand As ICommand = FindCommand(CommandName)

		If NewCommand Is Nothing Then
			NewCommand = CreateNewCommand(CommandName)
			ExistingCommands.Add(CommandName, NewCommand)
			NewCommand.Name = CommandName
		End If

		Return NewCommand
	End Function

	Public Sub UpdateCommands() Implements ICommandFactory.UpdateCommands
		Dim cmd As ICommand
		For Each cmd In Me.ExistingCommands.Values
			cmd.RaiseStateChanged()
		Next
	End Sub

	Protected MustOverride Function CreateNewCommand(ByVal CommandName As String) As Command

	Private mExistingCommands As Dict(Of ICommand) = New Dict(Of ICommand)
	Friend ReadOnly Property ExistingCommands() As Dict(Of ICommand)
		Get
			Return mExistingCommands
		End Get
	End Property

	Private Function FindCommand(ByVal CommandName As String) As ICommand
		Return ExistingCommands(CommandName)
	End Function

End Class
