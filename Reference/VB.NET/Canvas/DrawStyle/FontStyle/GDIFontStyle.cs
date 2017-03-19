#region Using directives

using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;

#endregion

namespace GuiLabs.Canvas.DrawStyle
{
    internal class GDIFontStyle: IFontStyleInfo
	{
		#region Color

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

		#endregion

		#region Font

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

		#endregion
	}
}