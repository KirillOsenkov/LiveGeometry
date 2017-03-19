Namespace Actions

	Friend Interface IActionHistory

		Sub RecordAction(ByVal NewAction As IAction)
		Sub Clear()

		Sub MoveBack()
		Sub MoveForward()

		Function CanMoveBack() As Boolean
		Function CanMoveForward() As Boolean
		Property Length() As Integer

	End Interface

End Namespace