using System;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Shapes;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;

namespace DynamicGeometry;

/// <summary>
/// Two small stacked buttons, a triangle up and a triangle down, meant to sit to the right
/// of a text box holding a number. Holding a button down repeats. The control knows nothing
/// about the number: it reports <see cref="Up"/> and <see cref="Down"/>, and
/// <see cref="Step"/> offers the usual arithmetic for whoever handles them.
/// </summary>
public class UpDownControl : Grid
{
    public const double ButtonWidth = 18;

    public UpDownControl()
    {
        RowDefinitions.Add(new RowDefinition());
        RowDefinitions.Add(new RowDefinition());
        Width = ButtonWidth;
        VerticalAlignment = VerticalAlignment.Stretch;

        var up = CreateButton(isUp: true);
        var down = CreateButton(isUp: false);
        SetRow(down, 1);
        Children.Add(up);
        Children.Add(down);

        up.Click += (s, e) => Up?.Invoke();
        down.Click += (s, e) => Down?.Invoke();
    }

    public event Action Up;
    public event Action Down;

    /// <summary>
    /// Makes the Up and Down arrow keys and the mouse wheel over <paramref name="textBox"/>
    /// do the same as the buttons.
    /// </summary>
    public void AttachTo(TextBox textBox)
    {
        textBox.AddHandler(
            KeyDownEvent,
            (s, e) =>
            {
                if (e.Key == Key.Up)
                {
                    Up?.Invoke();
                    e.Handled = true;
                }
                else if (e.Key == Key.Down)
                {
                    Down?.Invoke();
                    e.Handled = true;
                }
            },
            Avalonia.Interactivity.RoutingStrategies.Tunnel);

        textBox.PointerWheelChanged += (s, e) =>
        {
            if (e.Delta.Y > 0)
            {
                Up?.Invoke();
            }
            else if (e.Delta.Y < 0)
            {
                Down?.Invoke();
            }

            e.Handled = true;
        };
    }

    /// <summary>
    /// One step from <paramref name="value"/> in the given direction, landing on a multiple
    /// of <paramref name="step"/>: 1.4 goes up to 2 and down to 1, not to 2.4 and 0.4.
    /// Stays within the limits.
    /// </summary>
    public static double Step(double value, bool up, double step, double minimum, double maximum)
    {
        const double tolerance = 1e-9;
        double steps = value / step;
        double next = up
            ? System.Math.Floor(steps + tolerance) + 1
            : System.Math.Ceiling(steps - tolerance) - 1;
        return System.Math.Clamp(System.Math.Round(next * step, 6), minimum, maximum);
    }

    static RepeatButton CreateButton(bool isUp)
    {
        var triangle = new Avalonia.Controls.Shapes.Path()
        {
            Data = Geometry.Parse(isUp ? "M0,4 L3.5,0 L7,4 Z" : "M0,0 L3.5,4 L7,0 Z"),
            Fill = RibbonTheme.Text,
            Width = 7,
            Height = 4,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center
        };

        return new RepeatButton()
        {
            Content = triangle,
            Padding = new Thickness(0),
            MinHeight = 0,
            MinWidth = 0,
            HorizontalAlignment = HorizontalAlignment.Stretch,
            VerticalAlignment = VerticalAlignment.Stretch,
            HorizontalContentAlignment = HorizontalAlignment.Center,
            VerticalContentAlignment = VerticalAlignment.Center,
            Background = RibbonTheme.ButtonHover,
            BorderBrush = RibbonTheme.TabLine,
            BorderThickness = new Thickness(1),
            CornerRadius = isUp ? new CornerRadius(0, 4, 0, 0) : new CornerRadius(0, 0, 4, 0),
            Margin = isUp ? new Thickness(-1, 0, 0, 0) : new Thickness(-1, -1, 0, 0),
            Focusable = false,
            Delay = 400,
            Interval = 60
        };
    }
}
