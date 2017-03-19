using System.Collections.Generic;
using GuiLabs.Canvas.DrawStyle;
using GuiLabs.Canvas.Renderer;
using GuiLabs.Canvas.Utils;

namespace GuiLabs.Canvas.DrawOperations
{
	public class TransformDrawOperations : AbstractDrawOperations, IDrawOperations
	{
		private IDrawOperations mSourceDrawOperations;
		public IDrawOperations SourceDrawOperations
		{
			get
			{
				return mSourceDrawOperations;
			}
			set
			{
				mSourceDrawOperations = value;
			}
		}

		#region IDrawOperations Members

		public IDrawInfoFactory Factory
		{
			get
			{
				if (SourceDrawOperations == null)
				{
					return null;
				}
				return SourceDrawOperations.Factory;
			}
			set
			{
				if (SourceDrawOperations != null)
				{
					SourceDrawOperations.Factory = value;
				}
			}
		}

		protected Rect R = new Rect();
		private Point P1 = new Point();
		private Point P2 = new Point();

		protected virtual void TransformRect(Rect src)
		{
			Point p = R.Location;
			TransformPoint(src.Location, p);
			R.Size.Set(src.Size);
		}

		protected virtual void TransformPoint(Point src, Point dst)
		{
			dst.Set(src);
		}

		public virtual void DeTransformPoint(Point src, Point dst)
		{
			dst.Set(src);
		}

		private void TransformPointList(IList<Point> src)
		{
			foreach (Point p in src)
			{
				Point t = p;
				TransformPoint(p, t);
			}
		}

		public void DrawRectangle(Rect theRect, ILineStyleInfo theStyle)
		{
			TransformRect(theRect);
			SourceDrawOperations.DrawRectangle(R, theStyle);
		}

		public void DrawEllipse(Rect theRect, ILineStyleInfo theStyle)
		{
			TransformRect(theRect);
			SourceDrawOperations.DrawEllipse(R, theStyle);
		}

		public void DrawLine(Point p1, Point p2, ILineStyleInfo theStyle)
		{
			TransformPoint(p1, P1);
			TransformPoint(p2, P2);
			SourceDrawOperations.DrawLine(P1, P2, theStyle);
		}

		public override void DrawLine(int x1, int y1, int x2, int y2, ILineStyleInfo theStyle)
		{
			P1.Set(x1, y1);
			P2.Set(x2, y2);
			TransformPoint(P1, P1);
			TransformPoint(P2, P2);
			SourceDrawOperations.DrawLine(P1.X, P1.Y, P2.X, P2.Y, theStyle);
		}

		public void FillRectangle(Rect theRect, IFillStyleInfo theStyle)
		{
			TransformRect(theRect);
			SourceDrawOperations.FillRectangle(R, theStyle);
		}

		public void FillRectangle(Rect theRect, System.Drawing.Color fillColor)
		{
			TransformRect(theRect);
			SourceDrawOperations.FillRectangle(R, fillColor);
		}

		public void DrawFilledRectangle(Rect theRect, ILineStyleInfo Line, IFillStyleInfo Fill)
		{
			TransformRect(theRect);
			SourceDrawOperations.DrawFilledRectangle(R, Line, Fill);
		}

		public void DrawFilledEllipse(Rect theRect, ILineStyleInfo Line, IFillStyleInfo Fill)
		{
			TransformRect(theRect);
			SourceDrawOperations.DrawFilledEllipse(R, Line, Fill);
		}

		public void GradientFillRectangle(Rect theRect, System.Drawing.Color Color1, System.Drawing.Color Color2, System.Drawing.Drawing2D.LinearGradientMode GradientType)
		{
			TransformRect(theRect);
			SourceDrawOperations.GradientFillRectangle(R, Color1, Color2, GradientType);
		}

		public void FillPolygon(IList<Point> Points, ILineStyleInfo LineStyle, IFillStyleInfo FillStyle)
		{
			TransformPointList(Points);
			SourceDrawOperations.FillPolygon(Points, LineStyle, FillStyle);
		}

		public void DrawString(string Text, Rect theRect, IFontStyleInfo theFont)
		{
			TransformRect(theRect);
			SourceDrawOperations.DrawString(Text, R, theFont);
		}

		public Point StringSize(string Text, IFontInfo theFont)
		{
			return SourceDrawOperations.StringSize(Text, theFont);
		}


        public void DrawStringWithSelection(Rect Block, int StartSelPos, int CaretPosition, string Text, IFontStyleInfo FontStyle)
        {
			TransformRect(Block);
            SourceDrawOperations.DrawStringWithSelection(R, StartSelPos, CaretPosition, Text, FontStyle);
        }

		public void DrawImage(IPicture picture, Point p)
		{
			TransformPoint(p, P1);
			SourceDrawOperations.DrawImage(picture, P1);
		}

		//Rect oldCaretRect = new Rect();
		//public override void DrawCaret(Caret caret)
		//{
		//    if (!caret.Visible)
		//    {
		//        return;
		//    }

		//    // backup old caret coordinates
		//    oldCaretRect.Set(caret.Bounds);
			
		//    // transform the caret coordinates
		//    TransformRect(caret.Bounds);
		//    caret.SetNewBounds(R.Location.X, R.Location.Y, R.Size.Y);
			
		//    // draw the caret with transformed coordinates
		//    Source.DrawCaret(caret);

		//    // restore caret bounds
		//    caret.SetNewBounds(
		//        oldCaretRect.Location.X, 
		//        oldCaretRect.Location.Y, 
		//        oldCaretRect.Size.Y);
		//}

		#endregion
	}
}