using System;
using System.Drawing;
using System.Collections.Generic;
using Graphics = System.Drawing.Graphics;
using GuiLabs.Canvas;
using GuiLabs.Canvas.DrawStyle;
using GuiLabs.Canvas.Renderer;

namespace GuiLabs.Canvas.DrawOperations
{
	internal class GDIPlusDrawOperations : AbstractDrawOperations, IDrawOperations
	{
		private Graphics mGraphics;

		public GDIPlusDrawOperations(Graphics InitialGraphics)
		{
			mGraphics = InitialGraphics;
		}

		// =====================================================================

		public void DrawRectangle(Rect theRect, ILineStyleInfo theStyle)
		{
			mGraphics.DrawRectangle(((GDIPlusLineStyle)theStyle).Pen, theRect.GetRectangle());
		}

		public void DrawEllipse(Rect theRect, ILineStyleInfo theStyle)
		{
			mGraphics.DrawEllipse(((GDIPlusLineStyle)theStyle).Pen, theRect.GetRectangle());
		}

		public void FillRectangle(Rect theRect, IFillStyleInfo theStyle)
		{
			mGraphics.FillRectangle(((GDIPlusFillStyle)theStyle).Brush, theRect.GetRectangle());
		}

		public void FillRectangle(Rect theRect, System.Drawing.Color fillColor)
		{
			mGraphics.FillRectangle(new System.Drawing.SolidBrush(fillColor), theRect.GetRectangle());
		}

		public void FillEllipse(Rect theRect, System.Drawing.Color fillColor)
		{
			mGraphics.FillEllipse(new System.Drawing.SolidBrush(fillColor), theRect.GetRectangle());
		}

		public void DrawFilledRectangle(Rect theRect, ILineStyleInfo Line, IFillStyleInfo Fill)
		{
			FillRectangle(theRect, Fill);
			DrawRectangle(theRect, Line);
		}

		public void DrawFilledEllipse(Rect theRect, ILineStyleInfo Line, IFillStyleInfo Fill)
		{
			FillEllipse(theRect, Fill.FillColor);
			DrawEllipse(theRect, Line);
		}

		#region Gradients

		public void GradientFillRectangle(Rect theRect, System.Drawing.Color Color1, System.Drawing.Color Color2, System.Drawing.Drawing2D.LinearGradientMode GradientType)
		{
			System.Drawing.Rectangle R = theRect.GetRectangle();
			System.Drawing.Brush b = new System.Drawing.Drawing2D.LinearGradientBrush(R, Color1, Color2, GradientType);
			mGraphics.FillRectangle(b, R);
		}

		#endregion

		public void FillPolygon(IList<GuiLabs.Canvas.Point> Points, ILineStyleInfo LineStyle, IFillStyleInfo FillStyle)
		{
			GDIPlusLineStyle Line = (GDIPlusLineStyle)LineStyle;
			GDIPlusFillStyle Fill = (GDIPlusFillStyle)FillStyle;

			System.Drawing.Point[] P = new System.Drawing.Point[Points.Count];

			for (int i = 0; i < Points.Count; i++)
			{
				P[i].X = Points[i].X;
				P[i].Y = Points[i].Y;
			}

			mGraphics.FillPolygon(Fill.Brush, P);
			mGraphics.DrawPolygon(Line.Pen, P);
		}

		public void DrawLine(GuiLabs.Canvas.Point p1, GuiLabs.Canvas.Point p2, ILineStyleInfo theStyle)
		{
			mGraphics.DrawLine(((GDIPlusLineStyle)theStyle).Pen, p1.X, p1.Y, p2.X, p2.Y);
		}

		public override void DrawLine(int x1, int y1, int x2, int y2, ILineStyleInfo theStyle)
		{
			mGraphics.DrawLine(((GDIPlusLineStyle)theStyle).Pen, x1, y1, x2, y2);
		}

		public void DrawImage(IPicture picture, Point p)
		{
			GDIPlusPicture pict = picture as GDIPlusPicture;
			mGraphics.DrawImage(pict.Image, new System.Drawing.Point(p.X, p.Y));
		}

		private SolidBrush textBrush = new SolidBrush(Color.Black);
		private PointF textLocation = new PointF();
		public void DrawString(string Text, Rect theRect, IFontStyleInfo theFont)
		{
			if (theFont.ForeColor != textBrush.Color)
			{
				textBrush.Color = theFont.ForeColor;
			}
			theRect.Location.FillPoint(ref textLocation);
			mGraphics.DrawString(
				Text,
				((GDIPlusFontWrapper)(theFont.Font)).Font,
				textBrush,
				textLocation);
		}

		public GuiLabs.Canvas.Point StringSize(string Text, IFontInfo theFont)
		{
			System.Drawing.SizeF s = new System.Drawing.SizeF();

			s = mGraphics.MeasureString(Text, ((GDIPlusFontWrapper)theFont).Font);

			GuiLabs.Canvas.Point result = new GuiLabs.Canvas.Point((int)s.Width, (int)s.Height);
			return result;
		}

		// =====================================================================

		private IDrawInfoFactory mFactory = new GDIPlusDrawInfoFactory();
		public IDrawInfoFactory Factory
		{
			get
			{
				return mFactory;
			}
			set
			{
				mFactory = value;
			}
		}

        public void DrawStringWithSelection(Rect Block, int StartSelPos, int CaretPosition, string Text, IFontStyleInfo FontStyle)
        {
        }
	}
}
