namespace GuiLabs.Canvas.Events
{
	/// <summary>
	/// Everything that can RECEIVE mouse events
	/// and react to them
	/// </summary>
	public interface IMouseHandler
	{
		IMouseHandler DefaultMouseHandler { get; set; }
		bool NextHandlerValid(IMouseHandler nextHandler);
		void OnClick(MouseEventArgsWithKeys e);
		void OnDoubleClick(MouseEventArgsWithKeys e);
		void OnMouseDown(MouseEventArgsWithKeys e);
		void OnMouseHover(MouseEventArgsWithKeys e);
		void OnMouseMove(MouseEventArgsWithKeys e);
		void OnMouseUp(MouseEventArgsWithKeys e);
		void OnMouseWheel(MouseEventArgsWithKeys e);
	}
}