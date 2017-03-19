using System;
using System.Drawing;
using System.Collections;
using System.Text;

namespace GuiLabs.Canvas.DrawStyle
{
	internal class GDIDrawInfoFactory : IDrawInfoFactory
	{
		public GDIDrawInfoFactory()
		{
		}

		public ILineStyleInfo ProduceNewLineStyleInfo(Color theColor, int theWidth)
		{
			ILineStyleInfo NewLineStyle = new GDILineStyle(theColor, theWidth);
			return NewLineStyle;
		}

		public IFillStyleInfo ProduceNewFillStyleInfo(Color FillColor)
		{
			IFillStyleInfo NewFillStyle = new GDIFillStyle(FillColor);
			return NewFillStyle;
		}

		public IFontStyleInfo ProduceNewFontStyleInfo(string FamilyName, float size, System.Drawing.FontStyle theStyle)
		{
			IFontInfo NewFont = new GDIFont(FamilyName, size, theStyle);
			IFontStyleInfo NewFontStyle = new GDIFontStyle();
			NewFontStyle.Font = NewFont;
			return NewFontStyle;
		}

		public IPicture ProduceNewPicture(System.Drawing.Image image)
		{
			IPicture picture = new Picture(image);
			return picture;
		}

		public IPicture ProduceNewTransparentPicture(
			System.Drawing.Image image,
			System.Drawing.Color transparentColor
		)
		{
			Picture picture = new Picture(image);
			picture.Transparent = true;
			picture.TransparentColor = transparentColor;
			return picture;
		}

	}
}
