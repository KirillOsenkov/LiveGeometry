using System;
using System.Collections.Generic;
using System.Text;
using GuiLabs.Utils;
using System.Collections.Specialized;

namespace GuiLabs.Utils.Actions
{
	public class ActionManager : INotifyCollectionChanged
	{
		public ActionManager()
		{
			History = new SimpleHistory();
		}

		#region Events

        public event NotifyCollectionChangedEventHandler CollectionChanged;
		protected void RaiseUndoBufferChanged(object sender, NotifyCollectionChangedEventArgs e)
		{
            if (CollectionChanged != null)
            {
                CollectionChanged(this, e);
            }
		}

		#endregion

		#region RecordAction

		#region Running

		private IAction mCurrentAction = null;
		public IAction CurrentAction
		{
			get
			{
				return mCurrentAction;
			}
			internal set
			{
				mCurrentAction = value;
			}
		}

		public bool ActionIsExecuting
		{
			get
			{
				return CurrentAction != null;
			}
		}

		#endregion

		private bool mExecuteImmediatelyWithoutRecording = false;
		public bool ExecuteImmediatelyWithoutRecording
		{
			get
			{
				return mExecuteImmediatelyWithoutRecording;
			}
			set
			{
				mExecuteImmediatelyWithoutRecording = value;
			}
		}

		public void RecordAction(IAction existingAction)
		{
            if (existingAction == null)
            {
                throw new ArgumentNullException(
                    "ActionManager.RecordAction: the existingAction argument is null");
            }
			CheckNotRunningBeforeRecording(existingAction);

			if (ExecuteImmediatelyWithoutRecording 
				&& existingAction.CanExecute())
			{
				existingAction.Execute();
				return;
			}

			ITransaction currentTransaction = RecordingTransaction;
			if (currentTransaction != null)
			{
				currentTransaction.AccumulatingAction.Add(existingAction);
			}
			else
			{
				RunActionDirectly(existingAction);
			}
		}

		private void CheckNotRunningBeforeRecording(IAction existingAction)
		{
			string existing = existingAction != null ? existingAction.ToString() : "";

			if (CurrentAction != null)
			{
				throw new InvalidOperationException
				(
					string.Format
					(
						  "ActionManager.RecordActionDirectly: the ActionManager is currently running "
						+ "or undoing an action ({0}), and this action (while being executed) attempted "
						+ "to recursively record another action ({1}), which is not allowed. "
						+ "You can examine the stack trace of this exception to see what the "
						+ "executing action did wrong and change this action not to influence the "
						+ "Undo stack during its execution. Checking if ActionManager.ActionIsExecuting == true "
						+ "before launching another transaction might help to avoid the problem. Thanks and sorry for the inconvenience.",
						CurrentAction.ToString(),
						existing
					)
				);
			}
		}

		private object recordActionLock = new object();
		private void RunActionDirectly(IAction actionToRun)
		{
			CheckNotRunningBeforeRecording(actionToRun);

			lock (recordActionLock)
			{
				CurrentAction = actionToRun;
				if (History.AppendAction(actionToRun))
				{
					History.MoveForward();
				}
				CurrentAction = null;
			}
		}

		#endregion

		#region Transactions

		private Stack<ITransaction> mTransactionStack = new Stack<ITransaction>();
		public Stack<ITransaction> TransactionStack
		{
			get
			{
				return mTransactionStack;
			}
			set
			{
				mTransactionStack = value;
			}
		}

		public ITransaction RecordingTransaction
		{
			get
			{
				if (TransactionStack.Count > 0)
				{
					return TransactionStack.Peek();
				}
				return null;
			}
		}

		public void OpenTransaction(ITransaction t)
		{
			TransactionStack.Push(t);
		}

		public void CommitTransaction()
		{
			if (TransactionStack.Count == 0)
			{
				throw new InvalidOperationException(
					"ActionManager.CommitTransaction was called"
					+ " when there is no open transaction (TransactionStack is empty)."
					+ " Please examine the stack trace of this exception to find code"
					+ " which called CommitTransaction one time too many."
					+ " Normally you don't call OpenTransaction and CommitTransaction directly,"
					+ " but use using(Transaction t = new Transaction(Root)) instead.");
			}

			ITransaction committing = TransactionStack.Pop();

			if (committing.AccumulatingAction.Count > 0)
			{
				RecordAction(committing.AccumulatingAction);
			}
		}

		#endregion

		#region Undo, Redo

		public void Undo()
		{
			if (!CanUndo)
			{
				return;
			}
			if (ActionIsExecuting)
			{
				throw new InvalidOperationException(string.Format("ActionManager is currently busy"
					+ " executing a transaction ({0}). This transaction has called Undo()"
					+ " which is not allowed until the transaction ends."
					+ " Please examine the stack trace of this exception to see"
					+ " what part of your code called Undo.", CurrentAction));
			}
			CurrentAction = History.CurrentState.PreviousAction;
			History.MoveBack();
			CurrentAction = null;
		}

		public void Redo()
		{
			if (!CanRedo)
			{
				return;
			}
			if (ActionIsExecuting)
			{
				throw new InvalidOperationException(string.Format("ActionManager is currently busy"
					+ " executing a transaction ({0}). This transaction has called Redo()"
					+ " which is not allowed until the transaction ends."
					+ " Please examine the stack trace of this exception to see"
					+ " what part of your code called Undo.", CurrentAction));
			}
			CurrentAction = History.CurrentState.NextAction;
			History.MoveForward();
			CurrentAction = null;
		}

		public bool CanUndo
		{
			get
			{
				return History.CanMoveBack;
			}
		}

		public bool CanRedo
		{
			get
			{
				return History.CanMoveForward;
			}
		}

		#endregion

		#region Buffer

//		private void OutputBufferContents()
//		{
//			StringBuilder s = new StringBuilder();
//			foreach (IAction action in this.History.EnumUndoableActions())
//			{
//				s.AppendLine(action.GetType().ToString());
//			}
//			System.Windows.Forms.MessageBox.Show(s.ToString());
//		}

		public void Clear()
		{
			History.Clear();
		}

		public IEnumerable<IAction> EnumUndoableActions()
		{
			return History.EnumUndoableActions();
		}

		private IActionHistory mHistory;
		internal IActionHistory History
		{
			get
			{
				return mHistory;
			}
			set
			{
				if (mHistory != null)
				{
					mHistory.CollectionChanged -= RaiseUndoBufferChanged;
				}
				mHistory = value;
				if (mHistory != null)
				{
					mHistory.CollectionChanged += RaiseUndoBufferChanged;
				}
			}
		}

		#endregion
    }
}
