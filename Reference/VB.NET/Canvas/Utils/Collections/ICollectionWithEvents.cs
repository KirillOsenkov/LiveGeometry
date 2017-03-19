using System.Collections.Generic;
using GuiLabs.Canvas.Events;

namespace GuiLabs.Canvas.Utils
{
	public delegate void ElementAddedHandler<T> (T element);
	public delegate void ElementRemovedHandler<T> (T element);
	public delegate void ElementReplacedHandler<T> (T oldElement, T newElement);

	/// <summary>
	/// Collection that raises events, when elements are added or removed
	/// </summary>
	/// <typeparam name="T">Type of elements in the collection</typeparam>
	public interface ICollectionWithEvents<T> : IEnumerable<T>
	{
		// Events about
		event ElementAddedHandler<T> ElementAdded;
		event ElementRemovedHandler<T> ElementRemoved;
		event ElementReplacedHandler<T> ElementReplaced;
		event EmptyHandler CollectionChanged;
		IEnumerable<T> Reversed { get; }
		bool SuspendEvents { get; set;}
		T GetPrevElement(T element);
		T GetNextElement(T element);
	}
}
