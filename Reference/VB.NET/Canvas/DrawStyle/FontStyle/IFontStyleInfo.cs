#region Using directives

using System;
using System.Collections.Generic;
using System.Text;

#endregion

namespace GuiLabs.Canvas.DrawStyle
{
	public interface IFontStyleInfo
	{
		IFontInfo Font
		{
			get;
			set;
		}
		System.Drawing.Color ForeColor
		{
			get;
			set;
		}
	}
}
