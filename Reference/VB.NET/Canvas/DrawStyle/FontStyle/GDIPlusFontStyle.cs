#region Using directives

using System;
using System.Collections.Generic;
using System.Text;

#endregion

namespace GuiLabs.Canvas.DrawStyle
{
	internal class GDIPlusFontStyle : IFontStyleInfo
	{
		public GDIPlusFontStyle(String FamilyName, float size)
		{
			Font = new GDIPlusFontWrapper(FamilyName, size);
		}

		public GDIPlusFontStyle(String FamilyName, float size, System.Drawing.FontStyle style)
		{
			Font = new GDIPlusFontWrapper(FamilyName, size, style);
		}

		private IFontInfo mFont;
		public IFontInfo Font
		{
			get
			{
				return mFont;
			}
			set
			{
				mFont = value;
			}
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
			}
		}
	}
}
