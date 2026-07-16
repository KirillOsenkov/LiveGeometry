using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using Avalonia.Controls;
using Avalonia.Controls.Shapes;
using Avalonia.Input.Platform;
using Avalonia.Media;

namespace DynamicGeometry
{
    /// <summary>
    /// Stand-in for the WPF Style class. The library only ever uses Style as a
    /// bag of property setters applied to a single element, so this applies the
    /// setters directly instead of going through Avalonia's selector-based styling.
    /// </summary>
    public class Style
    {
        public Style()
        {
        }

        public Style(Type targetType)
        {
            TargetType = targetType;
        }

        public Type TargetType { get; set; }
        public List<Setter> Setters { get; } = new List<Setter>();

        // Tracks the style last applied to each element so that applying a new
        // style first clears the previous style's properties (WPF replace semantics).
        static readonly ConditionalWeakTable<Control, Style> appliedStyles = new();

        public void ApplyTo(Control element)
        {
            if (appliedStyles.TryGetValue(element, out var previous))
            {
                foreach (var setter in previous.Setters)
                {
                    element.ClearValue(setter.Property);
                }

                appliedStyles.Remove(element);
            }

            foreach (var setter in Setters)
            {
                element.SetValue(setter.Property, setter.Value);
            }

            appliedStyles.Add(element, this);
        }
    }

    /// <summary>Minimal WPF MessageBox stand-in; the host application decides how to present it.</summary>
    public static class MessageBox
    {
        public static Action<string> Handler { get; set; } = text => System.Diagnostics.Debug.WriteLine("MessageBox: " + text);

        public static void Show(string text)
        {
            Handler?.Invoke(text);
        }
    }

    /// <summary>
    /// WPF had a three-state Visibility; Avalonia only has IsVisible.
    /// Hidden and Collapsed both map to IsVisible = false (this library never
    /// relied on Hidden reserving layout space on the drawing canvas).
    /// </summary>
    public enum Visibility
    {
        Visible,
        Hidden,
        Collapsed
    }

    public static class WpfCompatExtensions
    {
        extension(Control control)
        {
            public Visibility Visibility
            {
                get => control.IsVisible ? Visibility.Visible : Visibility.Collapsed;
                set => control.IsVisible = value == Visibility.Visible;
            }

            public double ActualWidth => control.Bounds.Width;

            public double ActualHeight => control.Bounds.Height;

            // WPF-style parameterless mouse capture has no Avalonia equivalent
            // (capture is per-pointer, done from event args). The library only
            // used these in places where IsEnabled/IsHitTestVisible already
            // provide the behavior, so these are safe no-ops.
            public bool CaptureMouse() => false;

            public void ReleaseMouseCapture()
            {
            }
        }

        // WPF distinguishes start/end line caps; Avalonia has a single StrokeLineCap.
        extension(Shape shape)
        {
            public PenLineCap StrokeStartLineCap
            {
                get => shape.StrokeLineCap;
                set => shape.StrokeLineCap = value;
            }

            public PenLineCap StrokeEndLineCap
            {
                get => shape.StrokeLineCap;
                set => shape.StrokeLineCap = value;
            }
        }

        extension(LinearGradientBrush brush)
        {
            /// <summary>
            /// Reproduces the WPF LinearGradientBrush(stops, angle) constructor: gradient
            /// axis from (0,0) at the given angle in degrees, in relative coordinates.
            /// </summary>
            public void SetAngle(double angleDegrees)
            {
                var radians = angleDegrees * System.Math.PI / 180.0;
                brush.StartPoint = new Avalonia.RelativePoint(0, 0, Avalonia.RelativeUnit.Relative);
                brush.EndPoint = new Avalonia.RelativePoint(
                    System.Math.Cos(radians),
                    System.Math.Sin(radians),
                    Avalonia.RelativeUnit.Relative);
            }
        }

        extension(Avalonia.Input.PointerEventArgs e)
        {
            /// <summary>WPF MouseButtonEventArgs.ClickCount; only pointer-pressed events carry it in Avalonia.</summary>
            public int ClickCount => (e as Avalonia.Input.PointerPressedEventArgs)?.ClickCount ?? 1;
        }
    }

    /// <summary>
    /// WPF-style synchronous clipboard over Avalonia's async, TopLevel-scoped clipboard.
    /// An in-process cache keeps copy/paste inside the app fully synchronous (this also
    /// works in the browser sandbox); the system clipboard is updated best-effort.
    /// </summary>
    public static class Clipboard
    {
        static string text;

        /// <summary>Set by the host app to the TopLevel's clipboard once available.</summary>
        public static Avalonia.Input.Platform.IClipboard SystemClipboard { get; set; }

        public static bool ContainsText() => !string.IsNullOrEmpty(text);

        public static string GetText() => text;

        public static void SetText(string value)
        {
            text = value;
            try
            {
                SystemClipboard?.SetTextAsync(value);
            }
            catch
            {
                // Clipboard access can be denied (e.g. browser permissions); the in-process cache still works.
            }
        }

        // WPF Line exposes X1/Y1/X2/Y2; Avalonia Line has StartPoint/EndPoint.
        extension(Line line)
        {
            public double X1
            {
                get => line.StartPoint.X;
                set => line.StartPoint = line.StartPoint.WithX(value);
            }

            public double Y1
            {
                get => line.StartPoint.Y;
                set => line.StartPoint = line.StartPoint.WithY(value);
            }

            public double X2
            {
                get => line.EndPoint.X;
                set => line.EndPoint = line.EndPoint.WithX(value);
            }

            public double Y2
            {
                get => line.EndPoint.Y;
                set => line.EndPoint = line.EndPoint.WithY(value);
            }
        }
    }

    /// <summary>WPF exposed FontWeight values via the FontWeights class; Avalonia uses the FontWeight enum directly.</summary>
    public static class FontWeights
    {
        public static FontWeight Normal => FontWeight.Normal;
        public static FontWeight Bold => FontWeight.Bold;
        public static FontWeight SemiBold => FontWeight.SemiBold;
        public static FontWeight Light => FontWeight.Light;
    }

    /// <summary>WPF exposed FontStyle values via the FontStyles class; Avalonia uses the FontStyle enum directly.</summary>
    public static class FontStyles
    {
        public static FontStyle Normal => FontStyle.Normal;
        public static FontStyle Italic => FontStyle.Italic;
        public static FontStyle Oblique => FontStyle.Oblique;
    }
}
