using GuiLabs.Canvas.Utils;

namespace GuiLabs.Canvas.Shapes
{
	using ShapeCollection = CollectionWithEvents<IShape>;

	public class CompositeShape : ShapeList<IShape>
	{
		public CompositeShape()
		{
			WritableList = new ShapeCollection();
			Children = WritableList;
		}

		public void Add(IShape newShape)
		{
			WritableList.Add(newShape);
		}

		public void Remove(IShape shapeToRemove)
		{
			WritableList.Remove(shapeToRemove);
		}

		private ShapeCollection mWritableList;
		protected ShapeCollection WritableList
		{
			get
			{
				return mWritableList;
			}
			set
			{
				mWritableList = value;
			}
		}
	}
}
