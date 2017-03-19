using GuiLabs.Canvas.Events;
using GuiLabs.Canvas.Utils;

namespace GuiLabs.Canvas.Shapes
{
	/// <summary>
	/// Ein visuelles Objekt, das gezeichnet werden kann
	/// und das auf Maus und Tastatur reagieren kann
	/// </summary>
	public interface IShape : IDrawableRect, IMouseHandler, IKeyHandler
	{
		event SizeChangedHandler SizeChanged;
		IShape HitTest(int x, int y);
		bool Visible { get; set; }
		bool Enabled { get; set; }
	}
}
