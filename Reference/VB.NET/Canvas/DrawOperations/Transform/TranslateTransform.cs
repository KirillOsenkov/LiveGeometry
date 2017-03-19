using GuiLabs.Canvas.Utils;

namespace GuiLabs.Canvas.DrawOperations
{
	public class TranslateTransform : TransformDrawOperations
	{
		protected override void TransformPoint(Point src, Point dst)
		{
			dst.X = src.X - mDeltaX;
			dst.Y = src.Y - mDeltaY;
		}

		public override void DeTransformPoint(Point src, Point dst)
		{
			dst.X = src.X + mDeltaX;
			dst.Y = src.Y + mDeltaY;
		}

		private int mDeltaX;
		public int DeltaX
		{
			get
			{
				return mDeltaX;
			}
			set
			{
				mDeltaX = value;
			}
		}

		private int mDeltaY;
		public int DeltaY
		{
			get
			{
				return mDeltaY;
			}
			set
			{
				mDeltaY = value;
			}
		}
	}
}