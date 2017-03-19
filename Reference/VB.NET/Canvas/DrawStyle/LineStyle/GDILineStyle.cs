using System;

namespace GuiLabs.Canvas.DrawStyle
{
	internal class GDILineStyle : ILineStyleInfo
	{
		public GDILineStyle(System.Drawing.Color foreColor)
		{
			this.ForeColor = foreColor;
		}

		public GDILineStyle(System.Drawing.Color foreColor, int width)
		{
			this.Width = width;
			this.ForeColor = foreColor;
		}

		private System.Drawing.Color mForeColor = System.Drawing.Color.Black;
		public System.Drawing.Color ForeColor
		{
			get
			{
				return mForeColor;
			}
			set
			{
				mForeColor = value;
				mWin32ForeColor = System.Drawing.ColorTranslator.ToWin32(value);
			}
		}

		private int mWin32ForeColor = 0;
		public int Win32ForeColor
		{
			get
			{
				return mWin32ForeColor;
			}
			set
			{
				mWin32ForeColor = value;
				this.ForeColor = System.Drawing.ColorTranslator.FromWin32(mWin32ForeColor);
			}
		}

		private int mWidth = 1;
		public int Width
		{
			get
			{
				return mWidth;
			}
			set
			{
				mWidth = value;
			}
		}

		public void Dispose()
		{
		}
	}
}