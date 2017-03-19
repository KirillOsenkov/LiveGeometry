using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;
using System.Xml.Linq;
using GuiLabs.Undo;

namespace DynamicGeometry
{
    public class IgnoreAttribute : Attribute { }

    public struct PolarValue
    {
        public double Val;
    }

    public static class Tuple
    {
        public static Tuple<TP1, TP2> Create<TP1, TP2>(TP1 p1, TP2 p2)
            where TP1 : IEquatable<TP1>
            where TP2 : IEquatable<TP2>
        {
            return new Tuple<TP1, TP2>(p1, p2);
        }

        public static Tuple<TP1, TP2, TP3> Create<TP1, TP2, TP3>(TP1 p1, TP2 p2, TP3 p3)
        {
            return new Tuple<TP1, TP2, TP3>(p1, p2, p3);
        }
    }

    public struct Tuple<T1, T2> : IEquatable<Tuple<T1, T2>>
        where T1 : IEquatable<T1>
        where T2 : IEquatable<T2>
    {
        public Tuple(T1 item1, T2 item2)
            : this()
        {
            Item1 = item1;
            Item2 = item2;
        }

        public T1 Item1 { get; private set; }
        public T2 Item2 { get; private set; }

        public bool Equals(Tuple<T1, T2> other)
        {
            return Item1.Equals(other.Item1) && Item2.Equals(other.Item2);
        }

        public override bool Equals(object obj)
        {
            if (!(obj is Tuple<T1, T2>))
            {
                return false;
            }
            return Equals((Tuple<T1, T2>)obj);
        }

        public override int GetHashCode()
        {
            return Item1.GetHashCode() ^ Item2.GetHashCode();
        }

        public override string ToString()
        {
            return "(" + Item1.ToString() + "," + Item2.ToString() + ")";
        }
    }

    public class ListWithEvents<T> : Collection<T>
    {
        public event Action<T> ItemAdded;
        public event Action<T> ItemRemoved;

        protected override void ClearItems()
        {
            foreach (var item in this)
            {
                OnItemRemoved(item);
            }
            base.ClearItems();
        }

        private void OnItemRemoved(T item)
        {
            if (ItemRemoved != null)
            {
                ItemRemoved(item);
            }
        }

        protected override void InsertItem(int index, T item)
        {
            base.InsertItem(index, item);
            OnItemAdded(item);
        }

        private void OnItemAdded(T item)
        {
            if (ItemAdded != null)
            {
                ItemAdded(item);
            }
        }

        protected override void RemoveItem(int index)
        {
            OnItemRemoved(this[index]);
            base.RemoveItem(index);
        }

        protected override void SetItem(int index, T item)
        {
            OnItemRemoved(this[index]);
            base.SetItem(index, item);
            OnItemAdded(item);
        }
    }

    public struct Tuple<T1, T2, T3> : IEquatable<Tuple<T1, T2, T3>>
    {
        public Tuple(T1 item1, T2 item2, T3 item3)
            : this()
        {
            Item1 = item1;
            Item2 = item2;
            Item3 = item3;
        }

        public T1 Item1 { get; private set; }
        public T2 Item2 { get; private set; }
        public T3 Item3 { get; private set; }

        public bool Equals(Tuple<T1, T2, T3> other)
        {
            return Item1.Equals(other.Item1)
                && Item2.Equals(other.Item2)
                && Item3.Equals(other.Item3);
        }

        public override bool Equals(object obj)
        {
            if (!(obj is Tuple<T1, T2, T3>))
            {
                return false;
            }
            return Equals((Tuple<T1, T2, T3>)obj);
        }

        public override int GetHashCode()
        {
            return Item1.GetHashCode()
                ^ Item2.GetHashCode()
                ^ Item3.GetHashCode();
        }

        public override string ToString()
        {
            return "(" + Item1.ToString() + "," + Item2.ToString() + "," + Item3.ToString() + ")";
        }
    }

    public static partial class Extensions
    {
        public static string ToListString<T>(this IEnumerable<T> list, string separator)
        {
            if (list == null)
            {
                return null;
            }
            var strings = list.Select(i => i.ToString()).ToArray();
            var result = string.Join(separator, strings);
            return result;
        }

        public static void InsertDependencyCore(this IFigure figure, int index, IFigure dependency)
        {
            figure.Dependencies.Insert(index, dependency);
            dependency.Dependents.Add(figure);
            figure.RecalculateAndUpdateVisual();
        }

        public static void RemoveDependencyCore(this IFigure figure, int index, IFigure dependency)
        {
            figure.Dependencies.RemoveAt(index);
            dependency.Dependents.Remove(figure);
            figure.RecalculateAndUpdateVisual();
        }

#if !PLAYER

        /// <summary>
        /// Replaces a figure in a drawing with another figure, redirecting the dependents of the source figure to the new figure.
        /// </summary>
        /// <param name="figure">A figure (that already is in a Drawing) with a new figure</param>
        /// <param name="replacement">A new figure to replace the original (must belong to the same drawing!)</param>
        [Obsolete("Use Actions.ReplaceWithExisting() instead")]
        public static void ReplaceWithExisting(this IFigure figure, IFigure replacement)
        {
            Actions.ReplaceWithExisting(figure, replacement);
        }

        /// <summary>
        /// Replaces a figure in a drawing with another figure, redirecting the dependents of the source figure to the new figure.
        /// </summary>
        /// <param name="figure">A figure (that already is in a Drawing) with a new figure</param>
        /// <param name="replacement">A new figure to replace the original (must belong to the same drawing!)</param>
        [Obsolete("Use Actions.ReplaceWithNew() instead")]
        public static void ReplaceWithNew(this IFigure figure, IFigure replacement)
        {
            Actions.ReplaceWithNew(figure, replacement);
        }

        [Obsolete("Use Actions.Remove()")]
        public static void Delete(this IFigure figure)
        {
            Actions.Remove(figure);
        }

#endif

        public static string Format(this string str, params object[] arguments)
        {
            return string.Format(str, arguments);
        }

        public static Exception AsException(this string exceptionMessage)
        {
            return new Exception(exceptionMessage);
        }

        public static CompileResult CompileExpression(this Drawing drawing, string expression)
        {
            return MEFHost.Instance.CompilerService.CompileExpression(drawing, expression, f => true);
        }

#if !PLAYER

        public static void SetProperty(this ActionManager actionManager, object instance, string propertyName, object value)
        {
            var variable = new PropertyValue(propertyName, instance);
            Actions.SetProperty(actionManager, variable, value);
        }

#endif

        public static void Add<T>(this ICollection<T> list, params T[] items)
        {
            foreach (var item in items)
            {
                list.Add(item);
            }
        }

        public static void RemoveAll<T>(this ICollection<T> list, IEnumerable<T> itemsToRemove)
        {
            foreach (var item in itemsToRemove)
            {
                list.Remove(item);
            }
        }

        public static T Last<T>(this IList<T> list)
        {
            if (list == null || list.Count == 0)
            {
                return default(T);
            }
            return list[list.Count - 1];
        }

        public static IEnumerable<T> Without<T>(this IEnumerable<T> list, T itemToExclude)
        {
            var comparer = EqualityComparer<T>.Default;
            foreach (var item in list)
            {
                if (!comparer.Equals(item, itemToExclude))
                {
                    yield return item;
                }
            }
        }

        public static Point Point(this IEnumerable<IFigure> figures, int index)
        {
            return (figures.ElementAt(index) as IPoint).Coordinates;
        }

        public static PointPair Line(this IEnumerable<IFigure> figures, int index)
        {
            return (figures.ElementAt(index) as ILine).Coordinates;
        }

        public static void WriteAttributeDouble(this System.Xml.XmlWriter writer, string attributeName, double value)
        {
            writer.WriteAttributeString(attributeName, value.ToStringInvariant());
        }

        public static void WriteAttributeBool(this System.Xml.XmlWriter writer, string attributeName, bool value)
        {
            writer.WriteAttributeString(attributeName, value ? "true" : "false");
        }

        /// <summary>
        /// Doesn't check for duplicates
        /// </summary>
        public static void AddRange<T>(this ICollection<T> collection, IEnumerable<T> items)
        {
            foreach (var item in items)
            {
                collection.Add(item);
            }
        }

        /// <summary>
        /// Checks for duplicates and doesn't add items that are already there
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="collection"></param>
        /// <param name="items"></param>
        public static void Merge<T>(this ICollection<T> collection, IEnumerable<T> items)
        {
            foreach (var item in items)
            {
                if (!collection.Contains(item))
                {
                    collection.Add(item);
                }
            }
        }

        public static void SetItems<T>(this ICollection<T> collection, IEnumerable<T> items)
        {
            collection.Clear();
            items.ForEach(i => collection.Add(i));
        }

        public static PointPair GetSegment(this IList<Point> points, int startIndex)
        {
            return new PointPair(points[startIndex], points[(startIndex + 1) % points.Count]);
        }

        public static PointPair GetPreviousSegment(this IList<Point> points, int startIndex)
        {
            return new PointPair(points[startIndex > 0 ? startIndex - 1 : points.Count - 1], points[startIndex]);
        }

        public static int RotateNext(this int index, int count)
        {
            return (index + 1) % count;
        }

        public static int RotatePrevious(this int index, int count)
        {
            return index > 0 ? index - 1 : count - 1;
        }

        public static int RotatePrevious(this int index, int count, int steps)
        {
            int result = index - steps;
            if (result < 0)
            {
                result += count;
            }
            return result;
        }

        public static void RemoveLast<T>(this IList<T> list)
        {
            list.RemoveAt(list.Count - 1);
        }

        public static IEnumerable<int> Rotate(this int startIndex, int count)
        {
            while (true)
            {
                startIndex = startIndex.RotateNext(count);
                yield return startIndex;
            }
        }

        public static void SetZIndex<TShape>(this ShapeBase<TShape> shape, ZOrder typeOfLayer)
            where TShape : FrameworkElement
        {
            shape.ZIndex = (int)typeOfLayer;
            Canvas.SetZIndex(shape.Shape, (int)typeOfLayer);
        }

        public static void CenterAt(this FrameworkElement element, Point center)
        {
#if SILVERLIGHT
            var x = center.X - element.ActualWidth / 2;
            var y = center.Y - element.ActualHeight / 2;
#else
            var x = center.X - element.Width / 2;
            var y = center.Y - element.Height / 2;
#endif
            Canvas.SetLeft(element, x);
            Canvas.SetTop(element, y);
        }

        public static void MoveTo(this FrameworkElement element, Point center)
        {
            Canvas.SetLeft(element, center.X);
            Canvas.SetTop(element, center.Y);
        }

        public static void MoveOffset(this UIElement element, double xOffset, double yOffset)
        {
            if (element == null || double.IsNaN(xOffset) || double.IsNaN(yOffset))
            {
                return;
            }
            var coordinates = element.GetCoordinates();
            Canvas.SetLeft(element, coordinates.X + xOffset);
            Canvas.SetTop(element, coordinates.Y + yOffset);
        }

        public static void Move(
            this Line line,
            double x1,
            double y1,
            double x2,
            double y2,
            CoordinateSystem coordinateSystem)
        {
            var p1 = coordinateSystem.ToPhysical(new Point(x1, y1));
            var p2 = coordinateSystem.ToPhysical(new Point(x2, y2));
            line.X1 = p1.X;
            line.Y1 = p1.Y;
            line.X2 = p2.X;
            line.Y2 = p2.Y;
        }

        public static Point GetCoordinates(this UIElement element)
        {
            return new Point(Canvas.GetLeft(element), Canvas.GetTop(element));
        }

        public static IEnumerable<T> AsEnumerable<T>(this T singleElement)
        {
            yield return singleElement;
        }

        public static IEnumerable<T> Replace<T>(this IEnumerable<T> source, T oldItem, T newItem)
            where T : IEquatable<T>
        {
            foreach (var item in source)
            {
                if (item.Equals(oldItem))
                {
                    yield return newItem;
                }
                else
                {
                    yield return item;
                }
            }
        }

        public static void ForEach<T>(this IEnumerable<T> items, Action<T> action)
        {
            foreach (var item in items)
            {
                action(item);
            }
        }

        public static Point[] ToPoints(this IEnumerable<IFigure> figures)
        {
            return figures
                .OfType<IPoint>()
                .Select(p => p.Coordinates)
                .ToArray();
        }

        public static IEnumerable<Segment> ToSegments(this IEnumerable<IFigure> figures, Point coordinates)
        {
            return figures
                .OfType<Segment>()
                .Select(p => (Segment)p.GetFigureIfPointWithinTolerance(coordinates)).Where(p => p != null);
        }

        public static PointCollection ToPointCollection(this IEnumerable<Point> points)
        {
            var result = new PointCollection();
            foreach (var point in points)
            {
                result.Add(point);
            }

            return result;
        }

        public static IEnumerable<Point> ToLogical(this PointCollection points, CoordinateSystem coordinateSystem)
        {
            return coordinateSystem.ToLogical(points);
        }

        public static Point SetX(this Point p, double x)
        {
            return new Point(x, p.Y);
        }

        public static Point SetY(this Point p, double y)
        {
            return new Point(p.X, y);
        }

        public static Color ToColor(this int number)
        {
            return Color.FromArgb(
                255,
                (byte)(number & 0xFF),
                (byte)((number & 0xFF00) >> 8),
                (byte)((number & 0xFF0000) >> 16)
                );
        }

        public static Color ToColor(this string s)
        {
            if (s.IsEmpty())
            {
                return Colors.Black;
            }
            if (s.Length == 7)
            {
                return Color.FromArgb(
                    255,
                    Convert.ToByte(s.Substring(1, 2), 16),
                    Convert.ToByte(s.Substring(3, 2), 16),
                    Convert.ToByte(s.Substring(5, 2), 16));
            }
            if (s.Length == 6)
            {
                return Color.FromArgb(
                    255,
                    Convert.ToByte(s.Substring(0, 2), 16),
                    Convert.ToByte(s.Substring(2, 2), 16),
                    Convert.ToByte(s.Substring(4, 2), 16));
            }
            if (s.Length == 8)
            {
                return Color.FromArgb(
                    Convert.ToByte(s.Substring(0, 2), 16),
                    Convert.ToByte(s.Substring(2, 2), 16),
                    Convert.ToByte(s.Substring(4, 2), 16),
                    Convert.ToByte(s.Substring(6, 2), 16));
            }
            if (s.Length == 9)
            {
                return Color.FromArgb(
                    Convert.ToByte(s.Substring(1, 2), 16),
                    Convert.ToByte(s.Substring(3, 2), 16),
                    Convert.ToByte(s.Substring(5, 2), 16),
                    Convert.ToByte(s.Substring(7, 2), 16));
            }
            return Colors.Black;
        }

        public static bool IsEmpty<T>(this IEnumerable<T> collection)
        {
            return collection == null || !collection.Any();
        }

        public static bool IsEmpty<T>(this IList<T> collection)
        {
            return collection == null || collection.Count == 0;
        }

        public static bool IsEmpty(this string s)
        {
            return string.IsNullOrEmpty(s);
        }

        public static bool HasAttribute<T>(
            this MemberInfo attributeHost)
            where T : Attribute
        {
            return attributeHost.GetAttribute<T>() != null;
        }

        public static T GetAttribute<T>(
            this MemberInfo attributeHost)
            where T : Attribute
        {
            return (T)Attribute.GetCustomAttribute(attributeHost, typeof(T));
        }

        public static IEnumerable<T> GetAttributes<T>(
            this MemberInfo attributeHost)
            where T : Attribute
        {
            return Attribute.GetCustomAttributes(attributeHost, typeof(T)).OfType<T>();
        }

        public static bool HasInterface<T>(this Type type)
        {
            return typeof(T).IsAssignableFrom(type);
        }

        public static Visibility ToVisibility(this bool visible)
        {
            return visible ? Visibility.Visible : Visibility.Collapsed;
        }

        /// <summary>
        /// Returns physical coordinates
        /// </summary>
        /// <param name="canvas"></param>
        /// <returns></returns>
        public static PointPair GetBorderRectangle(this FrameworkElement canvas)
        {
            return new PointPair()
            {
                P1 = { X = 0, Y = 0 },
                P2 = { X = canvas.ActualWidth, Y = canvas.ActualHeight }
            };
        }

        public static void Set(this System.Windows.Shapes.Line line, PointPair coordinates)
        {
            line.X1 = coordinates.P1.X.Round() + 0.5;
            line.Y1 = coordinates.P1.Y.Round() + 0.5;
            line.X2 = coordinates.P2.X.Round() + 0.5;
            line.Y2 = coordinates.P2.Y.Round() + 0.5;
        }

        public static string ToStringInvariant(this double number)
        {
            return number.ToString(CultureInfo.InvariantCulture);
        }

        public static double ReadDouble(this XElement element, string attributeName)
        {
            double result = 0;
            var attribute = element.Attribute(attributeName);
            if (attribute != null)
            {
                double.TryParse(
                    attribute.Value,
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    out result);
            }
            return result;
        }

        public static bool ReadBool(this XElement element, string attributeName, bool defaultValue)
        {
            var attribute = element.Attribute(attributeName);
            if (attribute != null)
            {
                return bool.Parse(attribute.Value);
            }
            return defaultValue;
        }

        public static string ReadString(this XElement element, string attributeName)
        {
            var attribute = element.Attribute(attributeName);
            if (attribute != null)
            {
                return attribute.Value;
            }
            return null;
        }
    }

    public static partial class Utilities
    {
        public static string StripByteOrderMark(string text)
        {
            int index = text.IndexOf('<');
            if (index > 0)
            {
                text = text.Substring(index);
            }
            return text;
        }

        public static string ElapsedTime(Action code)
        {
            var time = DateTime.Now.Ticks;
            code();
            return ((DateTime.Now.Ticks - time) / 10000).ToString();
        }

        public static string Join(string separator, params object[] strings)
        {
            StringBuilder sb = new StringBuilder();
            bool first = true;
            foreach (var str in strings)
            {
                if (str == null)
                {
                    continue;
                }
                if (first)
                {
                    first = false;
                }
                else
                {
                    sb.Append(separator);
                }
                sb.Append(str);
            }
            return sb.ToString();
        }
    }

    public static partial class Check
    {
        [DebuggerHidden]
        public static void NotNull<T>(T obj) where T : class
        {
            if (obj == null)
            {
                throw new ArgumentNullException(typeof(T).FullName);
            }
        }

        [DebuggerHidden]
        public static void NotNull<T>(T obj, string objectName)
        {
            if (obj == null)
            {
                throw new ArgumentNullException(objectName ?? typeof(T).FullName);
            }
        }

        [DebuggerHidden]
        public static void Positive(double value)
        {
            if (value <= 0)
            {
                throw new ArgumentException("Value should be positive, and it is: " + value.ToString());
            }
        }

        [DebuggerHidden]
        public static void NotEmpty(string str)
        {
            if (str == null)
            {
                throw new ArgumentException("String argument is null.");
            }
            if (str == string.Empty)
            {
                throw new ArgumentException("String argument is an empty string.");
            }
        }

        [DebuggerHidden]
        public static void ElementCount<T>(IEnumerable<T> list, int expectedCount)
        {
            if (list == null)
            {
                throw "list is null".AsException();
            }
            if (list.Count() != expectedCount)
            {
                throw "list.Count() == {0} and expected {1}"
                    .Format(list.Count(), expectedCount)
                    .AsException();
            }
        }

        [DebuggerHidden]
        public static void NoNullElements<T>(IEnumerable<T> list)
        {
            int i = 0;
            foreach (var item in list)
            {
                if (item == null)
                {
                    throw "element {0} in the list is null"
                        .Format(i).AsException();
                }
                i++;
            }
        }
    }
}

#if SILVERLIGHT
namespace System.Windows.Media
{
    public static class Brushes
    {
        public static SolidColorBrush Black = new SolidColorBrush(Colors.Black);
        public static SolidColorBrush White = new SolidColorBrush(Colors.White);
    }
}
#endif