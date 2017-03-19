Namespace Actions

    '====================================================================
    ' Represents a node of the doubly linked-list SimpleHistory (StateX in the following diagram:)

    ' ((State0)) --- [Action0] --- ((State1)) --- [Action1] --- ((State2))

    ' StateX (e.g. State1) has a link to the previous State, previous Action,
    ' next State and next Action.
    ' As you move from State1 to State2, an Action1 is executed (Redo).
    ' As you move from State1 to State0, an Action0 is un-executed (Undo).
    '====================================================================

	Friend Class SimpleHistoryNode

		Public Sub New()

		End Sub

		Public Sub New(ByVal LastExistingAction As IAction, ByVal LastExistingState As SimpleHistoryNode)
			PreviousAction = LastExistingAction
			PreviousNode = LastExistingState
		End Sub

		'====================================================================

		Private mPreviousAction As IAction
		Public Property PreviousAction() As IAction
			Get
				Return mPreviousAction
			End Get
			Set(ByVal Value As IAction)
				mPreviousAction = Value
			End Set
		End Property

		Private mNextAction As IAction
		Public Property NextAction() As IAction
			Get
				Return mNextAction
			End Get
			Set(ByVal Value As IAction)
				mNextAction = Value
			End Set
		End Property

		'====================================================================

		Private mPreviousNode As SimpleHistoryNode
		Public Property PreviousNode() As SimpleHistoryNode
			Get
				Return mPreviousNode
			End Get
			Set(ByVal Value As SimpleHistoryNode)
				mPreviousNode = Value
			End Set
		End Property

		Private mNextNode As SimpleHistoryNode
		Public Property NextNode() As SimpleHistoryNode
			Get
				Return mNextNode
			End Get
			Set(ByVal Value As SimpleHistoryNode)
				mNextNode = Value
			End Set
		End Property

	End Class
End Namespace
