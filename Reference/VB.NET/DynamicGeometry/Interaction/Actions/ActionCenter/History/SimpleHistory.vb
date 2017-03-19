Namespace Actions

    '====================================================================
    ' IActionHistory represents a recorded list of all actions undertaken by user.

    ' This class implements a usual, linear action sequence. You can move back and forth
    ' changing the state of the respective document. When you move forward, you execute
    ' a respective action, when you move backward, you Undo it (UnExecute).

    ' Implemented through a double linked-list of SimpleHistoryNode objects.
    '====================================================================

	Friend Class SimpleHistory

		Implements IActionHistory

		Public Sub New()
			Clear()
		End Sub

		Public Sub Clear() Implements IActionHistory.Clear
			CurrentState = New SimpleHistoryNode()
		End Sub

		'====================================================================
		' "Iterator" to navigate through the sequence
		'====================================================================

		Private CurrentState As SimpleHistoryNode

		'====================================================================
        'Adds a new action to the tail after current state. If 
        'there exist more actions after this, they're lost (Garbage Collected).

        'This is the only method of this class that actually modifies the linked-list.
		'====================================================================

		Public Sub RecordAction(ByVal NewUserAction As DynamicGeometry.Actions.IAction) Implements IActionHistory.RecordAction
			CurrentState.NextAction = NewUserAction
			CurrentState.NextNode = New SimpleHistoryNode(NewUserAction, CurrentState)
		End Sub

		'====================================================================
		' Navigation.
		'====================================================================

		Public Function CanMoveForward() As Boolean Implements IActionHistory.CanMoveForward
			Return Not (CurrentState.NextAction Is Nothing Or CurrentState.NextNode Is Nothing)
		End Function

		Public Function CanMoveBack() As Boolean Implements IActionHistory.CanMoveBack
			Return Not (CurrentState.PreviousAction Is Nothing Or CurrentState.PreviousNode Is Nothing)
		End Function

		Public Sub MoveForward() Implements IActionHistory.MoveForward
			If Not CanMoveForward() Then Return
			CurrentState.NextAction.Execute()
			CurrentState = CurrentState.NextNode
			Length += 1
		End Sub

		Public Sub MoveBack() Implements IActionHistory.MoveBack
			If Not CanMoveBack() Then Return
			CurrentState.PreviousAction.UnExecute()
			CurrentState = CurrentState.PreviousNode
			Length -= 1
		End Sub

		'====================================================================
		' The length of Undo buffer (number of undertaken actions).
		'====================================================================

		Private mLength As Integer = 0
		Public Property Length() As Integer Implements IActionHistory.Length
			Get
				Return mLength
			End Get
			Set(ByVal Value As Integer)
				mLength = Value
			End Set
		End Property

	End Class
End Namespace
