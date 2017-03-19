using GuiLabs.Canvas.Shapes;
using GuiLabs.Canvas.Renderer;
using GuiLabs.Canvas.Events;
using GuiLabs.Canvas.Utils;

namespace GuiLabs.Canvas
{
	public interface IDrawWindow // : IMouseEvents
	{
		event RepaintHandler Repaint;
		event System.EventHandler Resize;
		event System.EventHandler GotFocus;

		event System.Windows.Forms.KeyEventHandler KeyDown;
		event System.Windows.Forms.KeyPressEventHandler KeyPress;
		event System.Windows.Forms.KeyEventHandler KeyUp;

		void Redraw();
		void Redraw(Rect ToRedraw);
		void Redraw(IDrawableRect ShapeToRedraw);

		Rect Bounds
		{
			get;
		}
	}
}
