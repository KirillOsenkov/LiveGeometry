using GuiLabs.Canvas.Events;
using GuiLabs.Canvas.Renderer;
using GuiLabs.Canvas.Utils;

namespace GuiLabs.Canvas.Shapes
{
	/// <summary>
	/// Base implementation for IShape
	/// </summary>
	public class Shape : KeyMouseHandler, IShape
	{
		public Shape()
		{

		}

		#region Events

		public event NeedRedrawHandler NeedRedraw;
		public event SizeChangedHandler SizeChanged;

		public void RaiseNeedRedraw(IDrawableRect ShapeToRedraw)
		{
			if (NeedRedraw != null)
			{
				NeedRedraw(ShapeToRedraw);
			}
		}

		public void RaiseNeedRedraw()
		{
			RaiseNeedRedraw(this);
		}

		public void RaiseSizeChanged(IShape ResizedShape, Point OldSize)
		{
			if (SizeChanged != null
				&&
				(OldSize.X != ResizedShape.Bounds.Size.X ||
				 OldSize.Y != ResizedShape.Bounds.Size.Y)
			)
			{
				SizeChanged(ResizedShape, OldSize);
			}
			//else
			//{
			//    if (SizeChanged != null)
			//    {
			//        SizeChanged(ResizedShape, OldSize);
			//        // System.Diagnostics.Debugger.Break();
			//    }
			//}
		}

		#endregion

		#region Draw

		private IDrawableRect mDefaultDrawHandler;
		/// <summary>
		/// If not null, all Draw requests will be
		/// redirected to this DefaultDrawHandler
		/// </summary>
		public IDrawableRect DefaultDrawHandler
		{
			get
			{
				return mDefaultDrawHandler;
			}
			set
			{
				mDefaultDrawHandler = value;
			}
		}

		public virtual void Draw(IRenderer Renderer)
		{
			if (DefaultDrawHandler != null)
			{
				DefaultDrawHandler.Draw(Renderer);
			}
		}

		#endregion

		protected Rect mBounds = new Rect();
		public virtual Rect Bounds
		{
			get
			{
				return mBounds;
			}
			set
			{
				mBounds = value;
			}
		}

		private Point oldSize = new Point();
		public void Layout()
		{
			oldSize.Set(this.Bounds.Size);

			LayoutCore();

			RaiseSizeChanged(this, oldSize);
		}

		public virtual void LayoutCore()
		{
		}

		public virtual IShape HitTest(int x, int y)
		{
			if (this.Bounds.HitTest(x, y) && this.Visible && this.Enabled)
			{
				return this;
			}
			return null;
		}

		public virtual void Move(int deltaX, int deltaY)
		{
			this.Bounds.Location.Add(deltaX, deltaY);
		}

		public void MoveTo(int x, int y)
		{
			Move(x - this.Bounds.Location.X, y - this.Bounds.Location.Y);
		}

		public void MoveTo(Point point)
		{
			Move(point.X - this.Bounds.Location.X, point.Y - this.Bounds.Location.Y);
		}

		private bool mEnabled = true;
		public virtual bool Enabled
		{
			get
			{
				return mEnabled;
			}
			set
			{
				mEnabled = value;
			}
		}

		private bool mVisible = true;
		public virtual bool Visible
		{
			get
			{
				return mVisible;
			}
			set
			{
				mVisible = value;
			}
		}
	}
}
