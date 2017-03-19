Imports DynamicGeometry.Actions

Namespace Actions

	Friend MustInherit Class AbstractAction

		Implements IAction

		Protected MustOverride Sub ExecuteCore()
		Protected MustOverride Sub UnExecuteCore()

		'====================================================================

		Public Sub Execute() Implements IAction.Execute
			If Not CanExecute() Then Return
			ExecuteCore()
			ExecuteCount += 1
		End Sub

		Public Sub UnExecute() Implements IAction.UnExecute
			If Not CanUnExecute() Then Return
			UnExecuteCore()
			ExecuteCount -= 1
		End Sub

		Private mExecuteCount As Integer = 0
		Public Property ExecuteCount() As Integer Implements IAction.ExecuteCount
			Get
				Return mExecuteCount
			End Get
			Set(ByVal Value As Integer)
				mExecuteCount = Value
			End Set
		End Property

		Public Overridable Function CanExecute() As Boolean Implements IAction.CanExecute
			Return ExecuteCount <= 0
		End Function

		Public Overridable Function CanUnExecute() As Boolean Implements IAction.CanUnExecute
			Return Not CanExecute()
		End Function

	End Class
End Namespace
