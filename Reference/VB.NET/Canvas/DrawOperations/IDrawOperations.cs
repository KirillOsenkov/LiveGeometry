using System.Collections.Generic;
using GuiLabs.Canvas.Utils;
using GuiLabs.Canvas.Renderer;
using GuiLabs.Canvas.DrawStyle;

namespace GuiLabs.Canvas.DrawOperations
{
	public interface IDrawOperations
	{
		IDrawInfoFactory Factory
		{
			get;
			set;
		}
		void DrawLine(Point p1, Point p2, ILineStyleInfo theStyle);
		void DrawLine(int x1, int y1, int x2, int y2, ILineStyleInfo theStyle);

		void DrawRectangle(Rect theRect, ILineStyleInfo theStyle);
		void DrawEllipse(Rect theRect, ILineStyleInfo theStyle);
		void FillRectangle(Rect theRect, IFillStyleInfo theStyle);
		void FillRectangle(Rect theRect, System.Drawing.Color fillColor);
		void DrawFilledRectangle(Rect theRect, ILineStyleInfo Line, IFillStyleInfo Fill);
		void DrawFilledEllipse(Rect theRect, ILineStyleInfo Line, IFillStyleInfo Fill);
		void FillPolygon(IList<Point> Points, ILineStyleInfo LineStyle, IFillStyleInfo FillStyle);
		void GradientFillRectangle(Rect theRect, System.Drawing.Color Color1, System.Drawing.Color Color2, System.Drawing.Drawing2D.LinearGradientMode GradientType);
		void DrawString(string Text, Rect theRect, IFontStyleInfo theFont);
		Point StringSize(string Text, IFontInfo theFont);
        void DrawStringWithSelection(Rect Block, int StartSelPos, int CaretPosition, string Text, IFontStyleInfo FontStyle);
		void DrawCaret(Caret Car);
		void DrawCaret(int x, int y, int height);
		void DrawImage(IPicture picture, Point point);
	}
}
