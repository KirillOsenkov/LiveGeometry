using System;
using System.Windows.Forms;
using System.Drawing;
using GuiLabs.Canvas.DrawOperations;
using GuiLabs.Canvas.Utils;

namespace GuiLabs.Canvas.Renderer
{
	public interface IRenderer : IDisposable
	{
		IDrawOperations DrawOperations
		{
			get;
			set;
		}

		Size ClientSize
		{
			get;
			set;
		}
		Color BackColor
		{
			get;
			set;
		}

		void Clear();
		void Clear(Rect Area);

		void RenderBuffer(Control DestinationControl, Rect ToRedraw);
		void RenderBuffer(Control DestinationControl, Rectangle r);
	}
}