using System;

namespace GuiLabs.Utils.Actions
{
	public interface ITransaction : IDisposable
	{
		IMultiAction AccumulatingAction { get; }
	}

	public interface IAction
	{
		/// <summary>
		/// Apply changes encapsulated by this object.
		/// </summary>
		/// <remarks>
		/// ExecuteCount++
		/// </remarks>
		void Execute();

		/// <summary>
		/// Undo changes made by a previous Execute call.
		/// </summary>
		/// <remarks>
		/// ExecuteCount--
		/// </remarks>
		void UnExecute();

		/// <summary>
		/// For most Actions, CanExecute is true when ExecuteCount = 0 (not yet executed)
		/// and false when ExecuteCount = 1 (already executed once)
		/// </summary>
		/// <returns>true if an encapsulated action can be applied</returns>
		bool CanExecute();
		
		/// <returns>true if an action was already executed and can be undone</returns>
		bool CanUnExecute();

		bool TryToMerge(IAction FollowingAction);
		bool AllowToMergeWithPrevious { get; set; }
	}
}
