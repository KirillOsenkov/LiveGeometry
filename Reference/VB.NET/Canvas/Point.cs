using System.Diagnostics;

namespace GuiLabs.Canvas
{
	[DebuggerStepThrough]
	public class Point
	{
		#region Constructors

		public Point(int x, int y)
		{
			X = x;
			Y = y;
		}

		public Point()
		{
		}

		public Point(Point CloneFrom)
		{
			X = CloneFrom.X;
			Y = CloneFrom.Y;
		}

		#endregion

		#region X

		private int mx = 0;
		public int X
		{
			get
			{
				return mx;
			}
			set
			{
				mx = value;
			}
		}

		#endregion

		#region Y

		private int my = 0;
		public int Y
		{
			get
			{
				return my;
			}
			set
			{
				my = value;
			}
		}

		#endregion

		#region Add

		public virtual void Add(int x, int y)
		{
			X += x;
			Y += y;
		}

		public void Add(int delta)
		{
			Add(delta, delta);
		}

		public void Add(Point p)
		{
			Add(p.X, p.Y);
		}

		#endregion

		#region Set

		public virtual void Set(int x, int y)
		{
			X = x;
			Y = y;
		}

		public void Set(int sizeOfBoth)
		{
			Set(sizeOfBoth, sizeOfBoth);
		}

		public void Set(Point p)
		{
			Set(p.X, p.Y);
		}

		public void Set(System.Drawing.Point point)
		{
			Set(point.X, point.Y);
		}

		public void Set(System.Drawing.Size point)
		{
			Set(point.Width, point.Height);
		}

		public void Set0()
		{
			Set(0, 0);
		}

		#endregion

		public void FillPoint(ref System.Drawing.PointF p)
		{
			p.X = X;
			p.Y = Y;
		}

		#region Operators

		public static Point operator +(Point p1, Point p2)
		{
			Point newPoint = new Point(p1.X + p2.X, p1.Y + p2.Y);
			return newPoint;
		}

		public static Point operator -(Point p1, Point p2)
		{
			Point newPoint = new Point(p1.X - p2.X, p1.Y - p2.Y);
			return newPoint;
		}

		public static Point operator +(Point p1, int i)
		{
			Point newPoint = new Point(p1.X + i, p1.Y + i);
			return newPoint;
		}

		public static Point operator -(Point p1, int i)
		{
			Point newPoint = new Point(p1.mx - i, p1.my - i);
			return newPoint;
		}

		#endregion

		public override string ToString()
		{
			System.Text.StringBuilder s = new System.Text.StringBuilder("(");
			s.Append(X);
			s.Append(", ");
			s.Append(Y);
			s.Append(")");
			return s.ToString();
		}
	}
}