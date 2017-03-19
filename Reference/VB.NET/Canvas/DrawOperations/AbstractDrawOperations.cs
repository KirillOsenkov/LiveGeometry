using GuiLabs.Canvas.DrawStyle;
using GuiLabs.Canvas.Renderer;

namespace GuiLabs.Canvas.DrawOperations
{
	public abstract class AbstractDrawOperations
	{
		#region Misc functions

		public enum EdgeType
		{
			Single,
			Raised,
			Sunken,
			Etched
		}

		// TODO: Kirill: add a function to draw raised and sunken edges,
		// similar to the Win32 API function DrawEdge
		public void DrawEdge()
		{
		}

		#endregion

		public abstract void DrawLine(int x1, int y1, int x2, int y2, ILineStyleInfo theStyle);

		public void DrawCaret(int x, int y, int height)
		{
			DrawLine(x, y, x, y + height, RendererSingleton.MyCaret.CaretStyle);
		}

		#region DrawCaret

		public virtual void DrawCaret(Caret caret)
		{
			// if (caret.Visible)
				DrawLine(
					caret.Bounds.Location.X,
					caret.Bounds.Location.Y,
					caret.Bounds.Location.X,
					caret.Bounds.Location.Y + caret.Bounds.Size.Y,
					caret.CaretStyle);
		}

		#endregion
	}
}