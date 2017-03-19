using System.Collections.Generic;
using System.Collections.Specialized;

namespace GuiLabs.Utils.Actions
{
	internal interface IActionHistory : IEnumerable<IAction>, INotifyCollectionChanged
	{
		/// <summary>
		/// Appends an action to the end of the Undo buffer.
		/// </summary>
		/// <param name="newAction">An action to append.</param>
		/// <returns>false if merged with previous, else true</returns>
		bool AppendAction(IAction newAction);
		void Clear();

		void MoveBack();
		void MoveForward();

		bool CanMoveBack { get; }
		bool CanMoveForward { get; }
		int Length { get;}

		SimpleHistoryNode CurrentState { get; }

		IEnumerable<IAction> EnumUndoableActions();
	}
}
