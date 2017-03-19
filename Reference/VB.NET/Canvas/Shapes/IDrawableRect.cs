using GuiLabs.Canvas;
using GuiLabs.Canvas.Renderer;

namespace GuiLabs.Canvas.Shapes
{
	public delegate void NeedRedrawHandler(IDrawableRect ShapeToRedraw);

	/// <summary>
	/// Alles was auf Canvas gezeichnet werden kann,
	/// muss diese Schnittstelle implementieren.
	/// </summary>
	public interface IDrawableRect : IDrawable
	{
		/// <summary>
		/// Benachrichtigung: "man muss mich neu zeichnen"
		/// </summary>
		event NeedRedrawHandler NeedRedraw;
		void RaiseNeedRedraw(); 
		void RaiseNeedRedraw(IDrawableRect ShapeToRedraw);

		void MoveTo(int x, int y);
		void Move(int deltaX, int deltaY);
		void MoveTo(Point point);

		/// <summary>
		/// Rechteckige Grenzen dieses Objekts
		/// </summary>
		Rect Bounds
		{
			get;
			set;
		}
	}
}
