using System;
using System.Drawing;

namespace GuiLabs.Canvas.DrawStyle
{
	internal class GDIPlusFillStyle : IFillStyleInfo
	{
		public GDIPlusFillStyle(Color FillColor)
		{
			Brush = new SolidBrush(FillColor);
		}

		private SolidBrush mBrush;
		public SolidBrush Brush
		{
			get
			{
				return mBrush;
			}
			set
			{
				mBrush = value;
			}
		}

		public Color FillColor
		{
			get
			{
				return mBrush.Color;
			}
			set
			{
				mBrush.Color = value;
			}
		}

		public void Dispose()
		{
			if(mBrush != null)
			{
				mBrush.Dispose();
				mBrush = null;
			}
		}

		~GDIPlusFillStyle()
		{
			Dispose();
		}
	}
}
