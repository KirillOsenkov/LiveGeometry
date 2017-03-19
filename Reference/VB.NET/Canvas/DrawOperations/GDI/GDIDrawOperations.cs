using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;


using GuiLabs.Canvas.Utils;
using GuiLabs.Canvas.DrawStyle;
using GuiLabs.Canvas.Renderer;

namespace GuiLabs.Canvas.DrawOperations
{
	internal class GDIDrawOperations : AbstractDrawOperations, IDrawOperations
	{
		public GDIDrawOperations(IntPtr InitialDC)
		{
			hDC = InitialDC;
			API.SetBkMode(hDC, 1);
		}

		private IntPtr hDC = IntPtr.Zero;
		private API.POINT NULLPOINT;

		public void DrawFilledRectangle(Rect theRect, ILineStyleInfo Line, IFillStyleInfo Fill)
		{
			GDILineStyle LineStyle = (GDILineStyle)Line;
			GDIFillStyle FillStyle = (GDIFillStyle)Fill;

			IntPtr hPen = API.CreatePen(0, Line.Width, LineStyle.Win32ForeColor);
			IntPtr hOldPen = API.SelectObject(hDC, hPen);
			IntPtr hBrush = API.CreateSolidBrush(FillStyle.Win32FillColor);
			IntPtr hOldBrush = API.SelectObject(hDC, hBrush);

			API.Rectangle(hDC, theRect.Location.X, theRect.Location.Y, theRect.Right, theRect.Bottom);

			API.SelectObject(hDC, hOldBrush);
			API.DeleteObject(hBrush);
			API.SelectObject(hDC, hOldPen);
			API.DeleteObject(hPen);
		}

		public void DrawFilledEllipse(Rect theRect, ILineStyleInfo Line, IFillStyleInfo Fill)
		{
			GDILineStyle LineStyle = (GDILineStyle)Line;
			GDIFillStyle FillStyle = (GDIFillStyle)Fill;

			IntPtr hPen = API.CreatePen(0, Line.Width, LineStyle.Win32ForeColor);
			IntPtr hOldPen = API.SelectObject(hDC, hPen);
			IntPtr hBrush = API.CreateSolidBrush(FillStyle.Win32FillColor);
			IntPtr hOldBrush = API.SelectObject(hDC, hBrush);

			API.Ellipse(hDC, theRect.Location.X, theRect.Location.Y, theRect.Right, theRect.Bottom);

			API.SelectObject(hDC, hOldBrush);
			API.DeleteObject(hBrush);
			API.SelectObject(hDC, hOldPen);
			API.DeleteObject(hPen);
		}

		public void FillRectangle(Rect theRect, IFillStyleInfo theStyle)
		{
			API.FillRectangle(hDC, theStyle.FillColor, theRect.GetRectangle());
		}

		public void FillRectangle(Rect theRect, System.Drawing.Color fillColor)
		{
			API.FillRectangle(hDC, fillColor, theRect.GetRectangle());
		}

		public void DrawRectangle(Rect theRect, ILineStyleInfo theStyle)
		{
			GDILineStyle Style = (GDILineStyle)theStyle;

			IntPtr hPen = API.CreatePen(0, theStyle.Width, Style.Win32ForeColor);
			IntPtr hOldPen = API.SelectObject(hDC, hPen);

			IntPtr hBrush = API.GetStockObject(5); // NULL_BRUSH
			IntPtr hOldBrush = API.SelectObject(hDC, hBrush);

			API.Rectangle(hDC, theRect.Location.X, theRect.Location.Y, theRect.Right, theRect.Bottom);

			API.SelectObject(hDC, hOldBrush);
			API.SelectObject(hDC, hOldPen);
			API.DeleteObject(hPen);

			//			int hPen = API.CreatePen(0, theStyle.Width, theStyle.Win32ForeColor);
			//			int hOldPen = API.SelectObject(hDC, hPen);
			//
			//			int left = theRect.Location.X;
			//			int top = theRect.Location.Y;
			//			int right = theRect.Right;
			//			int bottom = theRect.Bottom;
			//
			//			API.MoveToEx(hDC, left, top, ref NULLPOINTAPI);
			//			API.LineTo(hDC, right, top);
			//			API.LineTo(hDC, right, bottom);
			//			API.LineTo(hDC, left, bottom);
			//			API.LineTo(hDC, left, top);
			//
			//			API.SelectObject(hDC, hOldPen);
			//			API.DeleteObject(hPen);
		}

		public void DrawEllipse(Rect theRect, ILineStyleInfo theStyle)
		{
			GDILineStyle Style = theStyle as GDILineStyle;
			if (Style == null)
			{
				Log.Instance.WriteWarning("DrawEllipse: Style == null");
				return;
			}

			IntPtr hPen = API.CreatePen(0, theStyle.Width, Style.Win32ForeColor);
			IntPtr hOldPen = API.SelectObject(hDC, hPen);

			IntPtr hBrush = API.GetStockObject(5); // NULL_BRUSH
			IntPtr hOldBrush = API.SelectObject(hDC, hBrush);

			API.Ellipse(hDC, theRect.Location.X, theRect.Location.Y, theRect.Right, theRect.Bottom);

			API.SelectObject(hDC, hOldBrush);
			API.SelectObject(hDC, hOldPen);
			API.DeleteObject(hPen);
		}

		#region Gradients

		public void GradientFillRectangle(Rect theRect, System.Drawing.Color Color1, System.Drawing.Color Color2, System.Drawing.Drawing2D.LinearGradientMode GradientType)
		{
			int ColorR1 = Color1.R;
			int ColorG1 = Color1.G;
			int ColorB1 = Color1.B;

			int ColorR2 = Color2.R - ColorR1;
			int ColorG2 = Color2.G - ColorG1;
			int ColorB2 = Color2.B - ColorB1;

			int Width = theRect.Size.X;
			int Height = theRect.Size.Y;
			int x0 = theRect.Location.X;
			int y0 = theRect.Location.Y;
			int x1 = theRect.Right;
			int y1 = theRect.Bottom;

			if (Width <= 0 || Height <= 0)
				return;

			double Coeff;
			int StepSize;

			API.RECT R = new API.RECT();

			const int NumberOfSteps = 128; // number of steps

			if (GradientType == System.Drawing.Drawing2D.LinearGradientMode.Horizontal)
			{
				double InvWidth = 1.0 / Width;

				StepSize = Width / NumberOfSteps;
				if (StepSize < 1)
					StepSize = 1;

				R.Top = y0;
				R.Bottom = y1;

				for (int i = x0; i <= x1; i += StepSize)
				{
					R.Left = i;
					R.Right = i + StepSize;
					if (R.Right > x1)
					{
						R.Right = x1;
					}

					Coeff = (i - x0) * InvWidth;

					IntPtr hBrush = API.CreateSolidBrush((int)
						(int)(ColorR1 + (double)ColorR2 * Coeff) |
						(int)(ColorG1 + (double)ColorG2 * Coeff) << 8 |
						(int)(ColorB1 + (double)ColorB2 * Coeff) << 16
						);

					API.FillRect(hDC, ref R, hBrush);
					API.DeleteObject(hBrush);
				}
			}
			else
			{
				double InvHeight = 1.0 / Height;

				StepSize = Height / NumberOfSteps;

				if (StepSize < 1)
					StepSize = 1;

				R.Left = x0;
				R.Right = x1;

				for (int i = y0; i <= y1; i += StepSize)
				{
					R.Top = i;
					R.Bottom = i + StepSize;
					if (R.Bottom > y1)
					{

						R.Bottom = y1;

					}

					Coeff = (i - y0) * InvHeight;
					IntPtr hBrush = API.CreateSolidBrush(
						(int)(ColorR1 + (double)ColorR2 * Coeff) |
						(int)(ColorG1 + (double)ColorG2 * Coeff) << 8 |
						(int)(ColorB1 + (double)ColorB2 * Coeff) << 16
					);

					API.FillRect(hDC, ref R, hBrush);
					API.DeleteObject(hBrush);
				}
			}
		}

		#endregion

		public void DrawLine(Point p1, Point p2, ILineStyleInfo theStyle)
		{
			GDILineStyle Style = (GDILineStyle)theStyle;

			IntPtr hPen = API.CreatePen(0, theStyle.Width, Style.Win32ForeColor);
			IntPtr hOldPen = API.SelectObject(hDC, hPen);

			API.MoveToEx(hDC, p1.X, p1.Y, ref NULLPOINT);
			API.LineTo(hDC, p2.X, p2.Y);

			API.SelectObject(hDC, hOldPen);
			API.DeleteObject(hPen);
		}

		public override void DrawLine(int x1, int y1, int x2, int y2, ILineStyleInfo theStyle)
		{
			GDILineStyle Style = (GDILineStyle)theStyle;

			IntPtr hPen = API.CreatePen(0, theStyle.Width, Style.Win32ForeColor);
			IntPtr hOldPen = API.SelectObject(hDC, hPen);

			API.MoveToEx(hDC, x1, y1, ref NULLPOINT);
			API.LineTo(hDC, x2, y2);

			API.SelectObject(hDC, hOldPen);
			API.DeleteObject(hPen);
		}

		public void FillPolygon(IList<Point> Points, ILineStyleInfo LineStyle, IFillStyleInfo FillStyle)
		{
			GDILineStyle Line = (GDILineStyle)LineStyle;
			GDIFillStyle Fill = (GDIFillStyle)FillStyle;

			API.POINT[] P = new API.POINT[Points.Count];

			for (int i = 0; i < Points.Count; i++)
			{
				P[i].x = Points[i].X;
				P[i].y = Points[i].Y;
			}

			IntPtr hPen = API.CreatePen(0, Line.Width, Line.Win32ForeColor);
			IntPtr hOldPen = API.SelectObject(hDC, hPen);
			IntPtr hBrush = API.CreateSolidBrush(Fill.Win32FillColor);
			IntPtr hOldBrush = API.SelectObject(hDC, hBrush);

			API.Polygon(hDC, ref P[0], Points.Count);

			API.SelectObject(hDC, hOldBrush);
			API.DeleteObject(hBrush);
			API.SelectObject(hDC, hOldPen);
			API.DeleteObject(hPen);
		}

		#region Factory

		private IDrawInfoFactory mFactory = new GDIDrawInfoFactory();
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

		#endregion

		#region DrawString

		private int CurrentTextColor = 0;

		public void DrawString(string Text, Rect theRect, IFontStyleInfo theFont)
		{
			GDIFontStyle FontStyle = (GDIFontStyle)theFont;

			if (CurrentTextColor != FontStyle.Win32ForeColor)
			{
				CurrentTextColor = FontStyle.Win32ForeColor;
				API.SetTextColor(hDC, FontStyle.Win32ForeColor);
			}

			IntPtr hOldFont = API.SelectObject(hDC, ((GDIFont)FontStyle.Font).hFont);

			API.RECT r = new API.RECT();

			r.Left = theRect.Location.X;
			r.Top = theRect.Location.Y;
			r.Right = theRect.Right;
			r.Bottom = theRect.Bottom;

			// API.DrawText(hDC, Text, Text.Length, ref r, 2368);

			API.ExtTextOut(hDC, r.Left, r.Top, 4, ref r, Text, (uint)Text.Length, null);

			API.SelectObject(hDC, hOldFont);

			// No need to Delete hFont because we're going to reuse it
			// it is being saved in GDIFontStyle FontStyle
			// API.DeleteObject(hFont);
			
			// No need to restore old text color 
			// because we're setting it new each time anyway
			// API.SetTextColor(hDC, hOldColor);
		}

		private Point stringPos = new Point();
		private Point stringSize = new Point();
		private Rect stringRect = new Rect();

		public void DrawStringWithSelection
		(
			Rect Block,
			int StartSelPos,
			int CaretPosition,
			string Text,
			IFontStyleInfo FontStyle
		)
		{
			// API.SetROP2(hDC, 14);
			// API.Rectangle(hDC, theRect.Location.X, theRect.Location.Y, theRect.Right, theRect.Bottom);
			// FillRectangle(theRect, selectionFillStyle);

			DrawString(Text, Block, FontStyle);

			if (StartSelPos == CaretPosition) return;

			int SelStart, SelEnd;
			if (CaretPosition > StartSelPos)
			{
				SelStart = StartSelPos;
				SelEnd = CaretPosition;
			}
			else
			{
				SelStart = CaretPosition;
				SelEnd = StartSelPos;
			}

			if (SelStart < 0)
			{
				SelStart = 0;
			}

			if (SelEnd > Text.Length)
			{
				SelEnd = Text.Length;
			}

			// Added the if-statement to check if the selection borders 
			// are within the textlength
			string select = "";
			if ((SelStart < Text.Length) && (SelEnd <= Text.Length))
				select = Text.Substring(SelStart, SelEnd - SelStart);

			stringSize.Set(StringSize(Text.Substring(0, SelStart), FontStyle.Font));
			stringPos.Set(Block.Location.X + stringSize.X, Block.Location.Y);
			stringRect.Set(stringPos, StringSize(select, FontStyle.Font));

			FillRectangle(stringRect, System.Drawing.SystemColors.Highlight);

			System.Drawing.Color OldColor = FontStyle.ForeColor;
			FontStyle.ForeColor = System.Drawing.SystemColors.HighlightText;
			DrawString(select, stringRect, FontStyle);
			FontStyle.ForeColor = OldColor;
		}

		public Point StringSize(string Text, IFontInfo theFont)
		{
			IntPtr hOldFont = API.SelectObject(hDC, ((GDIFont)theFont).hFont);

			API.SIZE theSize = new API.SIZE();
			API.GetTextExtentPoint32(hDC, Text, Text.Length, out theSize);
			Point result = new Point(theSize.x, theSize.y);

			API.SelectObject(hDC, hOldFont);

			return result;
		}

		#endregion

		#region DrawImage

		public void DrawImage(IPicture picture, Point p)
		{
			Picture pict = picture as Picture;
			if (pict == null)
			{
				return;
			}

			if (pict.Transparent)
			{
				DrawImageTransparent(pict, p);
				return;
			}

			IntPtr hPictureDC = API.CreateCompatibleDC(hDC);
			IntPtr hOldBitmap = API.SelectObject(hPictureDC, pict.hBitmap);
			
			API.BitBlt(
				hDC, 
				p.X, 
				p.Y, 
				picture.Size.X,
				picture.Size.Y, 
				hPictureDC, 
				0, 
				0, 
				API.SRCCOPY);
			
			API.SelectObject(hPictureDC, hOldBitmap);
			API.DeleteDC(hPictureDC);
 		}

		private void DrawImageTransparent(Picture pict, Point p)
		{
			IntPtr hPictureDC = API.CreateCompatibleDC(hDC);
			IntPtr hOldBitmap = API.SelectObject(hPictureDC, pict.hBitmap);

			API.TransparentBlt(
				hDC,
				p.X,
				p.Y,
				pict.Size.X,
				pict.Size.Y,
				hPictureDC,
				0,
				0,
				pict.Size.X,
				pict.Size.Y,
				pict.Win32TransparentColor
			);

			API.SelectObject(hPictureDC, hOldBitmap);
			API.DeleteDC(hPictureDC);
		}

		#endregion
	}
}
