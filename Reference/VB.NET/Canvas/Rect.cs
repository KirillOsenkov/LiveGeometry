using System.Diagnostics;

namespace GuiLabs.Canvas
{
	[DebuggerStepThrough]
	public class Rect
	{
		public Rect()
		{
			Location = new Point(0, 0);
			Size = new Point(0, 0);
		}

		public Rect(Point NewLocation, Point NewSize)
		{
			Location = NewLocation;
			Size = NewSize;
			if (Location == null)
				Location = new Point(0, 0);
			if (Size == null)
				Size = new Point(0, 0);
		}

		public Rect(System.Drawing.Rectangle r)
		{
			Location = new Point(r.Left, r.Top);
			Size = new Point(r.Width, r.Height);
		}

		public Rect(Rect r)
		{
			Location = new Point(r.Location);
			Size = new Point(r.Size);
		}

		public Rect(int x, int y, int width, int height)
		{
			Location = new Point(x, y);
			Size = new Point(width, height);
		}

		public static Rect NullRect
		{
			get
			{
				return new Rect();
			}
		}

		private Point mLocation;
		public Point Location
		{
			get
			{
				return mLocation;
			}
			set
			{
				mLocation = value;
			}
		}

		public void Set(Rect R)
		{
			this.Location.Set(R.Location);
			this.Size.Set(R.Size);
		}

		public void Set(Point location, Point size)
		{
			this.Location.Set(location);
			this.Size.Set(size);
		}

		public void Set(int x, int y, int width, int height)
		{
			this.Location.Set(x, y);
			this.Size.Set(width, height);
		}

		public void Set(System.Drawing.Rectangle dotnetRect)
		{
			this.Location.X = dotnetRect.Left;
			this.Location.Y = dotnetRect.Top;
			this.Size.X = dotnetRect.Width;
			this.Size.Y = dotnetRect.Height;
		}

		public void Set0()
		{
			this.Location.Set0();
			this.Size.Set0();
		}

		public bool HitTest(Point HitPoint)
		{
			return HitTest(HitPoint.X, HitPoint.Y);
		}

		public bool HitTest(int x, int y)
		{
			return (
				(x >= this.Location.X) &&
				(x <= this.Right) &&
				(y >= this.Location.Y) &&
				(y <= this.Bottom)
				);
		}

		public bool Contains(Rect inner)
		{
			return
				inner.Location.X >= this.Location.X
				&& inner.Location.Y >= this.Location.Y
				&& inner.Right <= this.Right
				&& inner.Bottom <= this.Bottom;
		}

		private Point mSize;
		public Point Size
		{
			get
			{
				return mSize;
			}
			set
			{
				mSize = value;
			}
		}

		public int Right
		{
			get
			{
				return mLocation.X + mSize.X;
			}
		}

		public int Bottom
		{
			get
			{
				return mLocation.Y + mSize.Y;
			}
		}

		public int RelativeToRectY(Rect Reference)
		{
			int y1 = Reference.Location.Y;
			int y2 = Reference.Bottom;
			int top = this.Location.Y;
			int bottom = this.Bottom;

			if (top >= y2) return 5;

			if (bottom <= y1) return 1;

			if (top < y1)
			{
				if (bottom <= y2) return 2;
				return 0;
			}

			if (bottom > y2) return 4;

			return 3;
		}

		public bool IntersectsRectX(Rect Reference)
		{
			int x1 = Reference.Location.X;
			int x2 = Reference.Right;
			int left = this.Location.X;
			int right = this.Right;

			return (right > x1 && left < x2);
		}

		public bool IntersectsRectY(Rect Reference)
		{
			int y1 = Reference.Location.Y;
			int y2 = Reference.Bottom;
			int top = this.Location.Y;
			int bottom = this.Bottom;

			return (bottom > y1 && top < y2);
		}

		public bool IntersectsRect(Rect Reference)
		{
			return IntersectsRectX(Reference) && IntersectsRectY(Reference);
		}

		public void EnsurePositiveSize()
		{
			EnsurePositiveSizeX();
			EnsurePositiveSizeY();
		}

		public void EnsurePositiveSizeX()
		{
			int width = this.Size.X;
			if (width < 0)
			{
				this.Location.X += width;
				this.Size.X = -width;
			}
		}

		public void EnsurePositiveSizeY()
		{
			int height = this.Size.Y;
			if (height < 0)
			{
				this.Location.Y += height;
				this.Size.Y = -height;
			}
		}

		public void Unite(Rect ToContain)
		{
			int x1 = ToContain.mLocation.X;
			int y1 = ToContain.mLocation.Y;
			int x2 = ToContain.Right;
			int y2 = ToContain.Bottom;

			if (x1 < this.Location.X)
			{
				this.Size.X += this.Location.X - x1;
				this.Location.X = x1;
			}

			if (y1 < this.Location.Y)
			{
				this.Size.Y += this.Location.Y - y1;
				this.Location.Y = y1;
			}

			if (x2 > this.Right)
			{
				this.Size.X = x2 - this.Location.X;
			}

			if (y2 > this.Bottom)
			{
				this.Size.Y = y2 - this.Location.Y;
			}
		}

		public System.Drawing.Rectangle GetRectangle()
		{
			return new System.Drawing.Rectangle(mLocation.X, mLocation.Y, mSize.X, mSize.Y);
		}
	}
}