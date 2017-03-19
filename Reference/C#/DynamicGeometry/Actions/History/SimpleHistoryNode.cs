namespace GuiLabs.Utils.Actions
{
	/// <summary>
	/// Represents a node of the doubly linked-list SimpleHistory
	/// (StateX in the following diagram:)

	/// (State0) --- [Action0] --- (State1) --- [Action1] --- (State2)

	/// StateX (e.g. State1) has a link to the previous State, previous Action,
	/// next State and next Action.
	/// As you move from State1 to State2, an Action1 is executed (Redo).
	/// As you move from State1 to State0, an Action0 is un-executed (Undo).
	/// </summary>
	public class SimpleHistoryNode
	{
		public SimpleHistoryNode(IAction LastExistingAction, SimpleHistoryNode LastExistingState)
		{
			PreviousAction = LastExistingAction;
			PreviousNode = LastExistingState;
		}

		public SimpleHistoryNode()
		{

		}

		private IAction mPreviousAction;
		public IAction PreviousAction
		{
			get
			{
				return mPreviousAction;
			}
			set
			{
				mPreviousAction = value;
			}
		}

		private IAction mNextAction;
		public IAction NextAction
		{
			get
			{
				return mNextAction;
			}
			set
			{
				mNextAction = value;
			}
		}

		private SimpleHistoryNode mPreviousNode;
		public SimpleHistoryNode PreviousNode
		{
			get
			{
				return mPreviousNode;
			}
			set
			{
				mPreviousNode = value;
			}
		}

		private SimpleHistoryNode mNextNode;
		public SimpleHistoryNode NextNode
		{
			get
			{
				return mNextNode;
			}
			set
			{
				mNextNode = value;
			}
		}
	}
}
