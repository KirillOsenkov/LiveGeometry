Imports DynamicGeometry

Public Class DGCommandProxy

	Inherits DGAppCommand

	Public Sub New(ByVal ProxyCommandName As String)
		Me.Name = ProxyCommandName
	End Sub

	'====================================================================

	Private WithEvents mContainedCommand As ICommand
	Public Property ContainedCommand() As ICommand
		Get
			Return mContainedCommand
		End Get
		Set(ByVal Value As ICommand)
			mContainedCommand = Value
		End Set
	End Property

	Private Sub DGAppCommand_StateChanged() Handles MyBase.StateChanged
		If Not Me.ParentForm.GetActiveDocument Is Nothing Then
			ContainedCommand = Me.ParentForm.GetActiveDocument.Drawing.CmdFactory.GetCommand(Name)
			Flag = True
			ContainedCommand.RaiseStateChanged()
			Flag = False
		Else
			ContainedCommand = Nothing
		End If
	End Sub

	Private Sub mContainedCommand_StateChanged() Handles mContainedCommand.StateChanged
		If Flag Then Return
		RaiseStateChanged()
	End Sub

	Private Flag As Boolean = False

	'====================================================================

	Public Overrides Sub OnClick()
		If ContainedCommand Is Nothing Then Return
		ContainedCommand.OnClick()
	End Sub

	'====================================================================

	Public Overrides Property Checked() As Boolean
		Get
			If ContainedCommand Is Nothing Then
				Return MyBase.Checked
			End If
			Return ContainedCommand.Checked
		End Get
		Set(ByVal Value As Boolean)
			If ContainedCommand Is Nothing Then Return
			ContainedCommand.Checked = Value
		End Set
	End Property

	Public Overrides Property Enabled() As Boolean
		Get
			If ContainedCommand Is Nothing Then
				Return False
			End If
			Return ContainedCommand.Enabled
		End Get
		Set(ByVal Value As Boolean)
			If ContainedCommand Is Nothing Then Return
			ContainedCommand.Enabled = Value
		End Set
	End Property

	Public Overrides Property StatusText() As String
		Get
			If ContainedCommand Is Nothing Then
				Return MyBase.StatusText
			End If
			Return ContainedCommand.StatusText
		End Get
		Set(ByVal Value As String)
			If ContainedCommand Is Nothing Then Return
			ContainedCommand.StatusText = Value
		End Set
	End Property

	Public Overrides Property Tooltip() As String
		Get
			If ContainedCommand Is Nothing Then
				Return MyBase.Tooltip
			End If
			Return ContainedCommand.Tooltip
		End Get
		Set(ByVal Value As String)
			If ContainedCommand Is Nothing Then Return
			ContainedCommand.Tooltip = Value
		End Set
	End Property

	Public Overrides Property Visible() As Boolean
		Get
			If ContainedCommand Is Nothing Then
				Return False
			End If
			Return ContainedCommand.Visible
		End Get
		Set(ByVal Value As Boolean)
			If ContainedCommand Is Nothing Then Return
			ContainedCommand.Visible = Value
		End Set
	End Property

End Class
