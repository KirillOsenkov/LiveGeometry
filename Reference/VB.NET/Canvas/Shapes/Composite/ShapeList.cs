using System.Windows.Forms;

using GuiLabs.Canvas.Events;
using GuiLabs.Canvas.Renderer;
using GuiLabs.Canvas.Utils;

namespace GuiLabs.Canvas.Shapes
{
	/// <summary>
	/// Multiple shapes represented as one shape.
	/// </summary>
	/// <remarks>
	/// This class is very important, 
	/// because it contains the logic
	/// for handling several shapes as single shape,
	/// such as mouse redirection, drawing of several shapes, etc.
	/// </remarks>
	/// <typeparam name="T"></typeparam>
	public class ShapeList<T> : ShapeWithEvents, IShapeList<T>
		where T : class, IShape
	{
		public ShapeList()
			: base()
		{
		}

		#region Events

		/// <summary>
		/// ShouldSubscribeItem is raised when a new item
		/// was added to the ShapeList.
		/// Clients should subscribe to this new item's events.
		/// </summary>
		public event ChangeHandler<T> ShouldSubscribeItem;

		/// <summary>
		/// ShouldUnSubscribeItem is raised when an item
		/// was deleted from the ShapeList.
		/// </summary>
		public event ChangeHandler<T> ShouldUnSubscribeItem;
		public event EmptyHandler ShouldCallLayout;

		protected void RaiseShouldCallLayout()
		{
			if (ShouldCallLayout != null)
			{
				ShouldCallLayout();
			}
		}

		protected void RaiseSubscribeItem(T itemToSubscribe)
		{
			if (ShouldSubscribeItem != null)
			{
				ShouldSubscribeItem(itemToSubscribe);
			}
		}

		protected void RaiseUnSubscribeItem(T itemToUnsubscribe)
		{
			if (ShouldUnSubscribeItem != null)
			{
				ShouldUnSubscribeItem(itemToUnsubscribe);
			}
		}

		#endregion

		#region IShape

		private Point oldSize = new Point();

		/// <summary>
		/// Notify the client that this ShapeList's shapes
		/// have changed.
		/// </summary>
		public override void LayoutCore()
		{
			//oldSize.Set(mBounds.Size);

			//bool AlreadyInitialized = false;
			//mBounds.Size.Set0();
			//foreach (T s in Children)
			//{
			//    if (!AlreadyInitialized)
			//    {
			//        mBounds.Size.Set(s.Bounds.Size);
			//        AlreadyInitialized = true;
			//    }
			//    else
			//    {
			//        mBounds.Unite(s.Bounds);
			//    }
			//}

			RaiseShouldCallLayout();
		}

		/// <summary>
		/// Draw all shapes in this ShapeList on the Renderer.
		/// </summary>
		/// <param name="Renderer">Renderer object that provides DrawOperations etc.</param>
		public override void Draw(IRenderer Renderer)
		{
			foreach (T child in Children)
			{
				if (child.Visible)
				{
					child.Draw(Renderer);
				}
			}
		}

		public override void Move(int deltaX, int deltaY)
		{
			if (deltaX == 0 && deltaY == 0)
			{
				return;
			}
			this.Bounds.Location.Add(deltaX, deltaY);
			foreach (T child in Children)
			{
				child.Move(deltaX, deltaY);
			}
		}

		#region HitTest

		/// <summary>
		///		Tests if a point is inside this shape or its children
		/// </summary>
		/// <param name="HitPoint"></param>
		///		The point to check
		/// <returns>
		///		Reference to a child in which the point is (recursive),
		///		this, if the point is inside this shape but inside none of the children
		///		null else
		/// </returns>
		public override IShape HitTest(int x, int y)
		{
			if (!this.Bounds.HitTest(x, y) 
				|| !this.Visible
				|| !this.Enabled)
			{
				return null;
			}

			IShape foundChild = HitTestChildrenOnly(x, y);

			return foundChild != null ? foundChild : this;
		}

		/// <summary>
		/// Finds a subchild that contains a point
		/// </summary>
		/// <param name="x">x-coordinate of a hit test point</param>
		/// <param name="y">y-coordinate of a hit test point</param>
		/// <returns>Found child. null if none found.</returns>
		public IShape HitTestChildrenOnly(int x, int y)
		{
			foreach (T s in Children.Reversed)
			{
				IShape found = s.HitTest(x, y);
				if (found != null && found.Visible && found.Enabled)
				{
					return found;
				}
			}

			return null;
		}


		#endregion

		#endregion

		#region Children

		private ICollectionWithEvents<T> mChildren;
		public ICollectionWithEvents<T> Children
		{
			get
			{
				return mChildren;
			}
			set
			{
				if (mChildren != null)
				{
					mChildren.ElementAdded -= new ElementAddedHandler<T>(mList_ElementAdded);
					mChildren.ElementRemoved -= new ElementRemovedHandler<T>(mList_ElementRemoved);
					mChildren.ElementReplaced -= new ElementReplacedHandler<T>(mList_ElementReplaced);
					mChildren.CollectionChanged -= new EmptyHandler(mChildren_CollectionChanged);
				}
				mChildren = value;
				if (mChildren != null)
				{
					mChildren.ElementAdded += new ElementAddedHandler<T>(mList_ElementAdded);
					mChildren.ElementRemoved += new ElementRemovedHandler<T>(mList_ElementRemoved);
					mChildren.ElementReplaced += new ElementReplacedHandler<T>(mList_ElementReplaced);
					mChildren.CollectionChanged += new EmptyHandler(mChildren_CollectionChanged);
					// Don't need layout here. Will be called later.
					//Layout();
				}
			}
		}

		void mChildren_CollectionChanged()
		{
			//OnSizeChanged();
		}

		protected void item_SizeChanged(IShape ResizedShape, Point OldSize)
		{
			OnSizeChanged();
		}

		private void OnSizeChanged()
		{
			Layout();
		}

		void mList_ElementReplaced(T oldElement, T newElement)
		{
			UnsubscribeItem(oldElement);
			SubscribeItem(newElement);
			OnSizeChanged();
		}

		void mList_ElementRemoved(T element)
		{
			UnsubscribeItem(element);
			OnSizeChanged();
		}

		void mList_ElementAdded(T element)
		{
			SubscribeItem(element);
			OnSizeChanged();
		}

		protected virtual void SubscribeItem(T item)
		{
			item.NeedRedraw += new NeedRedrawHandler(item_NeedRedraw);
			item.SizeChanged += new SizeChangedHandler(item_SizeChanged);
			RaiseSubscribeItem(item);
		}

		protected virtual void UnsubscribeItem(T item)
		{
			item.NeedRedraw -= new NeedRedrawHandler(item_NeedRedraw);
			item.SizeChanged -= new SizeChangedHandler(item_SizeChanged);
			RaiseUnSubscribeItem(item);
		}

		protected void item_NeedRedraw(IDrawableRect ShapeToRedraw)
		{
			RaiseNeedRedraw(ShapeToRedraw);
		}

		#endregion

		#region IMouseHandler Members

		private IShape mCapture = null;
		public IShape Capture
		{
			get
			{
				return mCapture;
			}
			set
			{
				mCapture = value;
			}
		}

		private IShape ShapeToForwardMouseEventTo(MouseEventArgsWithKeys e)
		{
			IShape found = HitTestChildrenOnly(e.X, e.Y);
			return found;
		}

		public override void OnClick(MouseEventArgsWithKeys e)
		{
			DefaultMouseHandler = ShapeToForwardMouseEventTo(e);
			if (DefaultMouseHandler != null)
			{
				DefaultMouseHandler.OnClick(e);
			}
		}

		public override void OnDoubleClick(MouseEventArgsWithKeys e)
		{
			if (Capture != null)
			{
				Capture = null;
			}

			DefaultMouseHandler = ShapeToForwardMouseEventTo(e);
			if (DefaultMouseHandler != null)
			{
				DefaultMouseHandler.OnDoubleClick(e);
			}
		}

		public override void OnMouseDown(MouseEventArgsWithKeys e)
		{
			if (Capture != null)
			{
				Capture.OnMouseDown(e);
				return;
			}

			IShape clicked = ShapeToForwardMouseEventTo(e);
			DefaultMouseHandler = clicked;
			if (clicked != null)
			{
				Capture = clicked;
				DefaultMouseHandler.OnMouseDown(e);
			}
		}

		public override void OnMouseHover(MouseEventArgsWithKeys e)
		{
			DefaultMouseHandler = ShapeToForwardMouseEventTo(e);
			if (DefaultMouseHandler != null)
			{
				DefaultMouseHandler.OnMouseHover(e);
			}
		}

		public override void OnMouseMove(MouseEventArgsWithKeys e)
		{
			if (Capture != null)
			{
				Capture.OnMouseMove(e);
				return;
			}

			DefaultMouseHandler = ShapeToForwardMouseEventTo(e);
			if (DefaultMouseHandler != null)
			{
				DefaultMouseHandler.OnMouseMove(e);
			}
		}

		public override void OnMouseUp(MouseEventArgsWithKeys e)
		{
			if (Capture != null)
			{
				Capture.OnMouseUp(e);
				Capture = null;
				return;
			}

			DefaultMouseHandler = ShapeToForwardMouseEventTo(e);
			if (DefaultMouseHandler != null)
			{
				DefaultMouseHandler.OnMouseUp(e);
			}
		}

		public override void OnMouseWheel(MouseEventArgsWithKeys e)
		{
			if (Capture != null)
			{
				Capture = null;
			}

			DefaultMouseHandler = ShapeToForwardMouseEventTo(e);
			if (DefaultMouseHandler != null)
			{
				DefaultMouseHandler.OnMouseWheel(e);
			}
		}

		#endregion
	}
}
