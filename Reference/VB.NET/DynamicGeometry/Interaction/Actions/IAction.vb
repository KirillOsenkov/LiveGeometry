Namespace Actions

	' Encapsulates changes of the Document state.
	' This object represents a branch in a directed activity graph,
	' where nodes are states and directed branches are actions, that change states.

	' A branch is directed: applying changes is the main direction and undoing changes is
	' reverse direction.

	' E.g. (0. Start state: empty) --->[AddPoint action]---> (1. One point in the drawing)

	' Concrete implementations encapsulate additional information needed both by Execute & UnExecute

	Friend Interface IAction

		' apply changes encapsulated by this object
		' ExecuteCount++
		Sub Execute()

		' returns true if an encapsulated action can be applied
		' for most Actions, CanExecute is true when ExecuteCount = 0 (not yet executed)
		' and false when ExecuteCount = 1 (already executed once)
		Function CanExecute() As Boolean

		' restore changes already applied by this object
		' ExecuteCount--
		Sub UnExecute()

		' The same for reverse direction
		Function CanUnExecute() As Boolean

		' indicates which changes were already done by this IAction object
		' originally 0.
		Property ExecuteCount() As Integer

	End Interface
End Namespace