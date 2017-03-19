using System;

namespace GuiLabs.Canvas.DrawStyle
{
	public interface IFillStyleInfo : IDisposable
	{
		System.Drawing.Color FillColor
		{
			get;
			set;
		}
	}
}
