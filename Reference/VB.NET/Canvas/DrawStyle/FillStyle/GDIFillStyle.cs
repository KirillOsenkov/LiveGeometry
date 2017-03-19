using System;
using System.Drawing;

namespace GuiLabs.Canvas.DrawStyle
{
	internal class GDIFillStyle : IFillStyleInfo
	{
		public GDIFillStyle(Color fillColor)
		{
			this.FillColor = fillColor;
		}

		private Color mFillColor = new Color();
		public Color FillColor
		{
			get
			{
				return mFillColor;
			}
			set
			{
				mFillColor = value;
				mWin32FillColor = System.Drawing.ColorTranslator.ToWin32(value);
			}
		}

		private int mWin32FillColor = 0;
		public int Win32FillColor
		{
			get
			{
				return mWin32FillColor;
			}
			set
			{
				mWin32FillColor = value;
				this.FillColor = System.Drawing.ColorTranslator.FromWin32(value);
			}
		}

		public void Dispose()
		{
		}
	}
}
