using System;
using GuiLabs.Canvas.Events;
using GuiLabs.Canvas.Utils;

namespace GuiLabs.Canvas.Shapes
{
	public interface IShapeList<T> : IShapeWithEvents
		where T : class, IShape
	{
		ICollectionWithEvents<T> Children { get; set;}
		IShape Capture { get; set;}

		event ChangeHandler<T> ShouldSubscribeItem;
		event ChangeHandler<T> ShouldUnSubscribeItem;
		event EmptyHandler ShouldCallLayout;
	}
}
