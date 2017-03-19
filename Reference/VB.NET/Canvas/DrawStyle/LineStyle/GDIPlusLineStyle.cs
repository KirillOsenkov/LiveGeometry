using System;

namespace GuiLabs.Canvas.DrawStyle
{
	internal class GDIPlusLineStyle : ILineStyleInfo
	{
		public GDIPlusLineStyle(System.Drawing.Color ForeColor)
		{
			Pen = new System.Drawing.Pen(ForeColor);
			Width = 1;
		}

		public GDIPlusLineStyle(System.Drawing.Color ForeColor, int Width)
		{
			Pen = new System.Drawing.Pen(ForeColor, Width);
		}

		private System.Drawing.Pen mPen;
		public System.Drawing.Pen Pen
		{
			get
			{
				return mPen;
			}
			set
			{
				mPen = value;
			}
		}

		public System.Drawing.Color ForeColor
		{
			get
			{
				return mPen.Color;
			}
			set
			{
				mPen.Color = value;
			}
		}

		public int Width
		{
			get
			{
				return (int)mPen.Width;
			}
			set
			{
				mPen.Width = value;
			}
		}

		public void Dispose()
		{
			if (mPen != null)
			{
				mPen.Dispose();
				mPen = null;
			}
		}

		~GDIPlusLineStyle()
		{
			Dispose();
		}
	}
}