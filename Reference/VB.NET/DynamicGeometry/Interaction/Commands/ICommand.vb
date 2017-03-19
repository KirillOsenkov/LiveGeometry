
Public Interface ICommand

	Sub OnClick()	' Execute? Do?
	Sub RaiseStateChanged()

	Event StateChanged()

	Property Name() As String

	Property Enabled() As Boolean
	Property Visible() As Boolean
	Property Tooltip() As String
	Property StatusText() As String
	Property Checked() As Boolean

End Interface