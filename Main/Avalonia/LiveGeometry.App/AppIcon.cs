using System;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Shapes;
using Avalonia.Media;

namespace LiveGeometry;

/// <summary>
/// The Live Geometry mark, drawn in vectors: a classical column and a yellow set square leaning
/// next to it - what the original DG icon (Assets/DG.ico, 32 px) shows. On a 64 x 64 grid,
/// scaled to any size.
/// </summary>
public static class AppIcon
{
    const double Grid = 64;

    static readonly IBrush stoneOutline = new SolidColorBrush(Color.FromRgb(0x5C, 0x58, 0x52));
    static readonly IBrush flute = new SolidColorBrush(Color.FromArgb(0x60, 0x5C, 0x58, 0x52));
    static readonly IBrush rulerOutline = new SolidColorBrush(Color.FromRgb(0x9E, 0x80, 0x00));
    static readonly IBrush tick = new SolidColorBrush(Color.FromRgb(0x7E, 0x66, 0x00));
    static readonly IBrush shadow = new SolidColorBrush(Color.FromArgb(0x30, 0x00, 0x00, 0x00));

    public static Control Create(double size)
    {
        var canvas = new Canvas() { Width = Grid, Height = Grid };

        // marble: light on the left where the light comes from, darker on the right
        var marble = new LinearGradientBrush()
        {
            StartPoint = new RelativePoint(0, 0, RelativeUnit.Relative),
            EndPoint = new RelativePoint(1, 0, RelativeUnit.Relative),
            GradientStops =
            {
                new GradientStop(Color.FromRgb(0xF4, 0xF2, 0xEE), 0),
                new GradientStop(Color.FromRgb(0xD9, 0xD6, 0xD0), 0.45),
                new GradientStop(Color.FromRgb(0x9A, 0x96, 0x90), 1)
            }
        };

        // the set square's shadow on the floor, then the column, then the square in front
        canvas.Children.Add(Shape("M30,60 L62,60 L60,63 L28,63 Z", shadow, null));

        // column: plinth, base moulding, shaft, echinus, abacus
        canvas.Children.Add(Shape("M3,56 H27 V61 H3 Z", marble, stoneOutline));
        canvas.Children.Add(Shape("M5,52 H25 V56 H5 Z", marble, stoneOutline));
        canvas.Children.Add(Shape("M8,12 H22 V52 H8 Z", marble, stoneOutline));
        canvas.Children.Add(Shape("M6,9 Q15,5 24,9 V12 H6 Z", marble, stoneOutline));
        canvas.Children.Add(Shape("M3,4 H27 V9 H3 Z", marble, stoneOutline));

        // fluting. Path data is parsed with a point as the decimal separator, whatever the
        // user's culture: formatted in it, "17,5" read as two numbers and crashed the app
        // at startup in every browser set to German, Russian and the like.
        for (double x = 11.5; x < 22; x += 3.5)
        {
            canvas.Children.Add(Shape(FormattableString.Invariant($"M{x},13.5 V50.5"), null, flute, thickness: 1.2));
        }

        // set square: a right triangle with a triangular window, leaning on the column
        var yellow = new LinearGradientBrush()
        {
            StartPoint = new RelativePoint(0, 0, RelativeUnit.Relative),
            EndPoint = new RelativePoint(1, 1, RelativeUnit.Relative),
            GradientStops =
            {
                new GradientStop(Color.FromRgb(0xFF, 0xF4, 0xA0), 0),
                new GradientStop(Color.FromRgb(0xFF, 0xE5, 0x38), 0.5),
                new GradientStop(Color.FromRgb(0xF6, 0xCC, 0x1A), 1)
            }
        };
        canvas.Children.Add(Shape("M33,3 L63,60 L33,60 Z M38,24 L38,54 L54,54 Z", yellow, rulerOutline, thickness: 1.6, fillRule: FillRule.EvenOdd));

        // ticks down the long leg: every unit, a longer one every five - never past the
        // hypotenuse, which is close to the leg near the tip
        for (int i = 0; i < 11; i++)
        {
            double y = 8 + i * 4.8;
            double length = i % 5 == 0 ? 4.5 : 2.5;
            double room = (y - 3) * 30 / 57 - 1.2;
            length = System.Math.Min(length, room);
            if (length >= 1)
            {
                canvas.Children.Add(Shape(FormattableString.Invariant($"M33,{y} H{33 + length}"), null, tick, thickness: 1.1));
            }
        }

        return new Viewbox()
        {
            Width = size,
            Height = size,
            Stretch = Stretch.Uniform,
            Child = canvas
        };
    }

    static Path Shape(string data, IBrush fill, IBrush stroke, double thickness = 1.3, FillRule fillRule = FillRule.NonZero)
    {
        // the even-odd rule is spelled inside the path data (F0/F1), Geometry.Parse reads it
        var geometry = Geometry.Parse((fillRule == FillRule.EvenOdd ? "F0 " : "F1 ") + data);
        return new Path()
        {
            Data = geometry,
            Fill = fill,
            Stroke = stroke,
            StrokeThickness = stroke != null ? thickness : 0,
            StrokeJoin = PenLineJoin.Round,
            StrokeLineCap = PenLineCap.Round
        };
    }
}
