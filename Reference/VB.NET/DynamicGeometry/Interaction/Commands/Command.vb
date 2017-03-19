Public MustInherit Class Command
	Implements ICommand

	Public Event StateChanged() Implements ICommand.StateChanged
	Public Sub RaiseStateChanged() Implements ICommand.RaiseStateChanged
		RaiseEvent StateChanged()
	End Sub

	Public Overridable Sub OnClick() Implements ICommand.OnClick

	End Sub

	Private mName As String
	Public Property Name() As String Implements ICommand.Name
		Get
			Return mName
		End Get
		Set(ByVal Value As String)
			mName = Value
		End Set
	End Property

	Private mChecked As Boolean = False
	Public Overridable Property Checked() As Boolean Implements ICommand.Checked
		Get
			Return mChecked
		End Get
		Set(ByVal Value As Boolean)
			If Value <> mChecked Then
				mChecked = Value
				RaiseStateChanged()
			End If
		End Set
	End Property

	Private mEnabled As Boolean = True
	Public Overridable Property Enabled() As Boolean Implements ICommand.Enabled
		Get
			Return mEnabled
		End Get
		Set(ByVal Value As Boolean)
			If Value <> mEnabled Then
				mEnabled = Value
				RaiseStateChanged()
			End If
		End Set
	End Property

	Private mVisible As Boolean = True
	Public Overridable Property Visible() As Boolean Implements ICommand.Visible
		Get
			Return mVisible
		End Get
		Set(ByVal Value As Boolean)
			If Value <> mVisible Then
				mVisible = Value
				RaiseStateChanged()
			End If
		End Set
	End Property

	Private mStatusText As String = ""
	Public Overridable Property StatusText() As String Implements ICommand.StatusText
		Get
			Return mStatusText
		End Get
		Set(ByVal Value As String)
			mStatusText = Value
		End Set
	End Property

	Private mTooltip As String = ""
	Public Overridable Property Tooltip() As String Implements ICommand.Tooltip
		Get
			Return mTooltip
		End Get
		Set(ByVal Value As String)
			mTooltip = Value
		End Set
	End Property

End Class
