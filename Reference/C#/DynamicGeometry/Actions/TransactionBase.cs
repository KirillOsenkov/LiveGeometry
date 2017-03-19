using System;
using GuiLabs.Utils.Actions;

namespace GuiLabs.Utils.Actions
{
	public class TransactionBase : ITransaction
	{
		#region ctors

		public TransactionBase(ActionManager am)
		{
			ActionManager = am;
			if (am != null)
			{
				am.OpenTransaction(this);
			}
		}

		public TransactionBase()
		{

		}

		#endregion

		#region MultiAction

		protected IMultiAction mAccumulatingAction;
		public IMultiAction AccumulatingAction
		{
			get
			{
				return mAccumulatingAction;
			}
		}

		#endregion

		#region Commit

		public void Commit()
		{
			if (ActionManager != null)
			{
				ActionManager.CommitTransaction();
			}
		}

		#endregion

		#region ActionManager

		private ActionManager mActionManager;
		public ActionManager ActionManager
		{
			get
			{
				return mActionManager;
			}
			private set
			{
				if (value == null)
				{
					//throw new InvalidOperationException(
					//    "Transaction.Root should never be null.");
				}

				mActionManager = value;
				//if (mActionManager != null)
				//{
				//    mAccumulatingAction = new MultiAction(mRoot);
				//}
			}
		}

		#endregion

		#region Dispose

		public void Dispose()
		{
			Commit();
		}

		#endregion
	}
}
