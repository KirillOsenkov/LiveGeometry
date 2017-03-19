using System;
using System.Drawing;

namespace GuiLabs.Canvas.DrawStyle
{
	internal class GDIPlusDrawInfoFactory : IDrawInfoFactory
	{
		public GDIPlusDrawInfoFactory()
		{
		}

		public ILineStyleInfo ProduceNewLineStyleInfo(Color theColor, int theWidth)
		{
			ILineStyleInfo NewLineStyle = new GDIPlusLineStyle(theColor, theWidth);
			return NewLineStyle;
		}

		public IFillStyleInfo ProduceNewFillStyleInfo(Color FillColor)
		{
			IFillStyleInfo NewFillStyle = new GDIPlusFillStyle(FillColor);
			return NewFillStyle;
		}

		public IFontStyleInfo ProduceNewFontStyleInfo(string FamilyName, float size, System.Drawing.FontStyle theStyle)
		{
			IFontStyleInfo NewFontStyle = new GDIPlusFontStyle(FamilyName, size, theStyle);
			return NewFontStyle;
		}

		public IPicture ProduceNewPicture(System.Drawing.Image image)
		{
			GDIPlusPicture picture = new GDIPlusPicture(image);
			return picture;
		}

		public IPicture ProduceNewTransparentPicture(
			System.Drawing.Image image,
			System.Drawing.Color transparentColor
		)
		{
			return ProduceNewPicture(image);
		}
	}
}
