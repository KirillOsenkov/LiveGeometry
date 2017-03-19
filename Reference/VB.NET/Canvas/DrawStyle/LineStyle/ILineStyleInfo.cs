using System;

namespace GuiLabs.Canvas.DrawStyle
{
	public interface ILineStyleInfo : IDisposable
	{
		System.Drawing.Color ForeColor
		{
			get;
			set;
		}
		int Width
		{
			get;
			set;
		}
	}
}
