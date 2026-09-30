using System;
using System.Collections.Generic;
using System.Linq;
using Avalonia.Controls;
using Avalonia.Media;

namespace DynamicGeometry;

/// <summary>
/// The icon of a figure: the ribbon icon of the tool that makes it (a triangle gets the
/// Triangle tool's, a point at an intersection the Intersection tool's), drawn smaller.
/// </summary>
public static class FigureIcons
{
    /// <summary>Figure type -> the tool whose icon it takes; a type not here takes its base type's</summary>
    static readonly Dictionary<Type, Func<Behavior>> tools = new()
    {
        [typeof(PointBase)] = () => new FreePointCreator(),
        [typeof(PointByCoordinates)] = () => new PointByCoordinatesCreator(),
        [typeof(IntersectionPoint)] = () => new IntersectionCreator(),
        [typeof(MidPoint)] = () => new MidpointCreator(),
        [typeof(ReflectedPoint)] = () => new ReflectionCreator(),
        [typeof(RotatedPoint)] = () => new RotationCreator(),
        [typeof(DilatedPoint)] = () => new DilationCreator(),
        [typeof(TranslatedPoint)] = () => new TranslationCreator(),
        [typeof(Segment)] = () => new SegmentCreator(),
        [typeof(Ray)] = () => new RayCreator(),
        [typeof(LineBase)] = () => new LineTwoPointsCreator(),
        [typeof(ParallelLine)] = () => new ParallelLineCreator(),
        [typeof(PerpendicularLine)] = () => new PerpendicularLineCreator(),
        [typeof(SegmentBisector)] = () => new SegmentBisectorCreator(),
        [typeof(AngleBisector)] = () => new AngleBisectorCreator(),
        [typeof(LineAtAngle)] = () => new LineAtAngleCreator(),
        [typeof(LineByEquation)] = () => new LineByEquationCreator(),
        [typeof(Vector)] = () => new VectorCreator(),
        [typeof(EllipseBase)] = () => new CircleCreator(),
        [typeof(CircleByRadius)] = () => new CircleByRadiusCreator(),
        [typeof(CircleByEquation)] = () => new CircleByEquationCreator(),
        [typeof(Ellipse)] = () => new EllipseCreator(),
        [typeof(EllipseArcBase)] = () => new EllipseArcCreator(),
        [typeof(CircleArcBase)] = () => new CircleArcCreator(),
        [typeof(AngleArc)] = () => new AngleMeasurementCreator(),
        [typeof(PolygonBase)] = () => new PolygonCreator(),
        [typeof(RegularPolygon)] = () => new RegularPolygonCreator(),
        [typeof(PolygonIntersection)] = () => new PolygonIntersectionCreator(),
        [typeof(Polyline)] = () => new PolylineCreator(),
        [typeof(Bezier)] = () => new BezierCreator(),
        [typeof(Curve)] = () => new FunctionGraphCreator(),
        [typeof(Locus)] = () => new LocusCreator(),
        [typeof(DistanceMeasurement)] = () => new DistanceMeasurementCreator(),
        [typeof(AngleMeasurementBase)] = () => new AngleMeasurementCreator(),
        [typeof(AreaMeasurement)] = () => new AreaMeasurementCreator(),
        [typeof(ControlBase)] = () => new LabelCreator(),
        [typeof(Slider)] = () => new SliderCreator(),
        [typeof(Number)] = () => new SliderCreator(),
    };

    static readonly Dictionary<string, Behavior> toolsMade = new();

    /// <summary>Tells which icon the figure gets, so that a list can see when it changes (a triangle given a fourth vertex)</summary>
    public static string GetKey(IFigure figure)
    {
        if (figure is Polygon polygon)
        {
            var points = polygon.Dependencies.OfType<IPoint>().ToArray();
            if (points.Length == 3)
            {
                return nameof(TriangleCreator);
            }

            if (points.Length == 4
                && Quadrilaterals.Classify(points[0].Coordinates, points[1].Coordinates, points[2].Coordinates, points[3].Coordinates) == "Square")
            {
                return nameof(SquareCreator);
            }
        }

        for (var type = figure.GetType(); type != null; type = type.BaseType)
        {
            if (tools.ContainsKey(type))
            {
                return type.FullName;
            }
        }

        return null;
    }

    /// <summary>The icon for <see cref="GetKey"/>, <paramref name="size"/> pixels square; null for a figure no tool makes</summary>
    public static Control Create(string key, double size)
    {
        var icon = CreateFullSize(key);
        if (icon == null)
        {
            return null;
        }

        return new Viewbox()
        {
            Width = size,
            Height = size,
            Stretch = Stretch.Uniform,
            Child = icon
        };
    }

    static Control CreateFullSize(string key)
    {
        switch (key)
        {
            case null:
                return null;
            case nameof(TriangleCreator):
                return GetTool(key, () => new TriangleCreator()).CreateIcon();
            case nameof(SquareCreator):
                return GetTool(key, () => new SquareCreator()).CreateIcon();
        }

        var entry = tools.First(pair => pair.Key.FullName == key);
        return GetTool(key, entry.Value).CreateIcon();
    }

    /// <summary>One instance of each tool, only ever asked for its icon</summary>
    static Behavior GetTool(string key, Func<Behavior> create)
    {
        if (!toolsMade.TryGetValue(key, out var tool))
        {
            tool = create();
            toolsMade.Add(key, tool);
        }

        return tool;
    }
}
