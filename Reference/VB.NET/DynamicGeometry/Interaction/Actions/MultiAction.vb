Namespace Actions

	Friend Class MultiAction
		Inherits ToggleAction

		Private mInnerActions As ArrayList = New ArrayList()
		Public Property InnerActions() As ArrayList
			Get
				Return mInnerActions
			End Get
			Set(ByVal Value As ArrayList)
				mInnerActions = Value
			End Set
		End Property

		Protected Overrides Sub ExecuteCore()
			Dim i As Integer
			For i = 0 To InnerActions.Count - 1
				DirectCast(InnerActions.Item(i), IAction).Execute()
			Next
		End Sub

		Protected Overrides Sub UnExecuteCore()
			Dim i As Integer
			For i = InnerActions.Count - 1 To 0 Step -1
				DirectCast(InnerActions.Item(i), IAction).UnExecute()
			Next
		End Sub
	End Class

End Namespace
