// WPF/Silverlight compatibility aliases for the Avalonia port.
// The geometry library was written against the WPF API surface; these aliases
// map the WPF type names that appear throughout the code onto their closest
// Avalonia equivalents so the port stays a minimal diff against the WPF library.

global using UIElement = Avalonia.Controls.Control;
global using FrameworkElement = Avalonia.Controls.Control;

// WPF Mouse* event args all map onto Avalonia pointer event args. Behavior.cs
// funnels all input through virtual methods, so only the arg types matter here.
global using MouseEventArgs = Avalonia.Input.PointerEventArgs;
global using MouseButtonEventArgs = Avalonia.Input.PointerEventArgs;
global using MouseWheelEventArgs = Avalonia.Input.PointerWheelEventArgs;

global using PointCollection = Avalonia.Collections.AvaloniaList<Avalonia.Point>;
global using DoubleCollection = Avalonia.Collections.AvaloniaList<double>;
global using PathSegmentCollection = Avalonia.Media.PathSegments;
global using PathFigureCollection = Avalonia.Media.PathFigures;
global using GradientStopCollection = Avalonia.Media.GradientStops;
global using Setter = Avalonia.Styling.Setter;
global using Selector = Avalonia.Controls.Primitives.SelectingItemsControl;
global using RoutedEventHandler = System.EventHandler<Avalonia.Interactivity.RoutedEventArgs>;
global using SizeChangedEventHandler = System.EventHandler<Avalonia.Controls.SizeChangedEventArgs>;
