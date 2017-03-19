Imports System.Windows.Forms
Imports DynamicGeometry

Public Class DGMenuItem
	Inherits MenuItem
	Implements ICommandCarrier

	Public Sub New(ByVal NewCommand As Command)
		MyBase.New()
		UICommand = NewCommand
		UpdateState()
	End Sub

	Public Sub New()
		MyBase.New()
	End Sub

	Private WithEvents mUICommand As ICommand
	Public Property UICommand() As ICommand Implements DGApp.ICommandCarrier.UICommand
		Get
			Return mUICommand
		End Get
		Set(ByVal Value As ICommand)
			mUICommand = Value
		End Set
	End Property

	Private Sub mUICommand_StateChanged() Handles mUICommand.StateChanged
		UpdateState()
	End Sub

	Public Sub UpdateState()
		Me.Visible = UICommand.Visible
		Me.Enabled = UICommand.Enabled
		Me.Checked = UICommand.Checked
	End Sub

	Private Sub DGMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Click
		UICommand.OnClick()
	End Sub

End Class
