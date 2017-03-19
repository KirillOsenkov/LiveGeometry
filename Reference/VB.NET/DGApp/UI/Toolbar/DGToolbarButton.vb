Imports System.Windows.Forms
Imports DynamicGeometry

Public Class DGToolBarButton

	Inherits System.Windows.Forms.ToolBarButton
	Implements ICommandCarrier

	Private WithEvents mUICommand As ICommand
	Public Property UICommand() As ICommand Implements DGApp.ICommandCarrier.UICommand
		Get
			Return mUICommand
		End Get
		Set(ByVal Value As ICommand)
			mUICommand = Value
		End Set
	End Property

	Private Sub New()
		MyBase.New()
	End Sub

	Public Sub New(ByVal NewCommand As Command)
		UICommand = NewCommand
		UpdateState()
	End Sub

	Private Sub mUICommand_StateChanged() Handles mUICommand.StateChanged
		UpdateState()
	End Sub

	Public Sub UpdateState()
		Me.Visible = UICommand.Visible
		Me.Enabled = UICommand.Enabled
		Me.Pushed = UICommand.Checked
	End Sub

End Class
