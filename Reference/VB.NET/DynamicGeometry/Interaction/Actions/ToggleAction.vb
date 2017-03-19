Imports DynamicGeometry.Actions

Friend MustInherit Class ToggleAction
	Inherits DGAction

	Public Overrides Function CanExecute() As Boolean
		Return Me.ExecuteCount = 0
	End Function

	Public Overrides Function CanUnExecute() As Boolean
		Return Me.ExecuteCount = 1
	End Function

End Class
