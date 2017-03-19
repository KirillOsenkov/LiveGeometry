using System.Windows.Forms;
using GuiLabs.Canvas.Events;

namespace GuiLabs.Canvas.Shapes
{
	public class ShapeWithEvents : Shape, IShapeWithEvents
	{
		public ShapeWithEvents()
			: base()
		{

		}

		#region Events

		public event MouseWithKeysEventHandler Click;
		public event MouseWithKeysEventHandler DoubleClick;
		public event MouseWithKeysEventHandler MouseDown;
		public event MouseWithKeysEventHandler MouseMove;
		public event MouseWithKeysEventHandler MouseUp;
		public event MouseWithKeysEventHandler MouseHover;
		public event MouseWithKeysEventHandler MouseWheel;
		
		public event KeyEventHandler KeyDown;
		public event KeyPressEventHandler KeyPress;
		public event KeyEventHandler KeyUp;

		#endregion

		#region RaiseMouseEvent

		protected void RaiseClick(MouseEventArgsWithKeys e)
		{
			if (Click != null)
			{
				Click(e);
			}
		}

		protected void RaiseDoubleClick(MouseEventArgsWithKeys e)
		{
			if (DoubleClick != null)
			{
				DoubleClick(e);
			}
		}

		protected void RaiseMouseDown(MouseEventArgsWithKeys e)
		{
			if (MouseDown != null)
			{
				MouseDown(e);
			}
		}

		protected void RaiseMouseMove(MouseEventArgsWithKeys e)
		{
			if (MouseMove != null)
			{
				MouseMove(e);
			}
		}

		protected void RaiseMouseUp(MouseEventArgsWithKeys e)
		{
			if (MouseUp != null)
			{
				MouseUp(e);
			}
		}

		protected void RaiseMouseHover(MouseEventArgsWithKeys e)
		{
			if (MouseHover != null)
			{
				MouseHover(e);
			}
		}

		protected void RaiseMouseWheel(MouseEventArgsWithKeys e)
		{
			if (MouseWheel != null)
			{
				MouseWheel(e);
			}
		}

		#endregion

		#region RaiseKeyEvent

		public void RaiseKeyDown(KeyEventArgs e)
		{
			if (KeyDown != null)
			{
				KeyDown(this, e);
			}
		}

		protected void RaiseKeyPress(KeyPressEventArgs e)
		{
			if (KeyPress != null)
			{
				KeyPress(this, e);
			}
		}

		protected void RaiseKeyUp(KeyEventArgs e)
		{
			if (KeyUp != null)
			{
				KeyUp(this, e);
			}
		}

		#endregion

		#region OnMouseEvent

		public override void OnClick(MouseEventArgsWithKeys e)
		{
			base.OnClick(e);
			RaiseClick(e);
		}

		public override void OnDoubleClick(MouseEventArgsWithKeys e)
		{
			base.OnDoubleClick(e);
			RaiseDoubleClick(e);
		}

		public override void OnMouseDown(MouseEventArgsWithKeys e)
		{
			base.OnMouseDown(e);
			RaiseMouseDown(e);
		}

		public override void OnMouseMove(MouseEventArgsWithKeys e)
		{
			base.OnMouseMove(e);
			RaiseMouseMove(e);
		}

		public override void OnMouseUp(MouseEventArgsWithKeys e)
		{
			base.OnMouseUp(e);
			RaiseMouseUp(e);
		}

		public override void OnMouseHover(MouseEventArgsWithKeys e)
		{
			base.OnMouseHover(e);
			RaiseMouseHover(e);
		}

		public override void OnMouseWheel(MouseEventArgsWithKeys e)
		{
			base.OnMouseWheel(e);
			RaiseMouseWheel(e);
		}

		#endregion

		#region OnKeyEvent

		public override void OnKeyDown(KeyEventArgs e)
		{
			base.OnKeyDown(e);
			RaiseKeyDown(e);
		}

		public override void OnKeyPress(KeyPressEventArgs e)
		{
			base.OnKeyPress(e);
			RaiseKeyPress(e);
		}

		public override void OnKeyUp(KeyEventArgs e)
		{
			base.OnKeyUp(e);
			RaiseKeyUp(e);
		}

		#endregion
	}
}
