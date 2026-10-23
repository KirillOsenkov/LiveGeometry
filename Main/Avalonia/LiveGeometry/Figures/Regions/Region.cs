using System.Collections.Generic;
using System.Linq;
using Avalonia;
using Avalonia.Controls.Shapes;
using Avalonia.Media;

namespace LiveGeometry;

/// <summary>
/// A set made of other figures with an inside (<see cref="IRegion"/>: polygons, circles,
/// ellipses, sectors, circular segments, Bezier paths, other regions) by a
/// <see cref="RegionOperation"/>: the first without the others, what they all share, or all
/// of them together. The inside alone, with no edge of its own: it is filled, and its
/// boundary is traced from the operands' (<see cref="RegionBoundary"/>) for the fill, the
/// area, the perimeter and the point kept inside it. Disjoint parts are one region, and a
/// set that comes to nothing (two circles pulled apart under an intersection) is a region
/// of area 0 that draws nothing.
/// </summary>
public class Region : ShapeBase<Path>, IShapeWithInterior, IPerimeter, IRegion
{
    RegionOperation operation = RegionOperation.Difference;
    RegionBoundary boundary = RegionBoundary.Empty;

    protected override string Kind => "Region";

    /// <summary>Region r, s... (p is a polygon's)</summary>
    protected override string FirstLetter => "r";

    /// <summary>"c without ABCD", "intersection of c and d", "union of c, d and e"</summary>
    public override string Construction => Describe(operation, Dependencies);

    /// <summary>What a region of the operands would say of itself (the tool's panel says it before the region is made); null without any</summary>
    public static string Describe(RegionOperation operation, IList<IFigure> operands)
    {
        if (operands.Count == 0)
        {
            return null;
        }

        var names = operands.Select(ConstructionText.Of).ToList();
        switch (operation)
        {
            case RegionOperation.Difference:
                return names.Count == 1 ? names[0] : names[0] + " without " + Join(names.Skip(1).ToList());
            case RegionOperation.Intersection:
                return "intersection of " + Join(names);
            default:
                return "union of " + Join(names);
        }
    }

    /// <summary>"c", "c and d", "c, d and e"</summary>
    static string Join(IReadOnlyList<string> names)
    {
        if (names.Count <= 1)
        {
            return names.Count == 1 ? names[0] : "";
        }

        return string.Join(", ", names.Take(names.Count - 1)) + " and " + names[names.Count - 1];
    }

    /// <summary>What the region makes of its operands; the Figure List and the grid say it in words (<see cref="Construction"/>)</summary>
    [PropertyGridVisible]
    [PropertyGridName("Operation")]
    public RegionOperation Operation
    {
        get => operation;
        set
        {
            if (operation == value)
            {
                return;
            }

            operation = value;
            if (Drawing != null)
            {
                this.RecalculateAllDependents();
                RaiseConstructionChanged();
            }

            RaisePropertyChanged(nameof(Operation));
        }
    }

    /// <summary>The figures the region is made of, in order (the first is what a difference starts from); null where a dependency has no inside</summary>
    public IReadOnlyList<IRegion> Operands => Dependencies.Select(dependency => dependency as IRegion).ToList();

    /// <summary>The boundary as it was last worked out</summary>
    public RegionBoundary Boundary => boundary;

    protected override ZOrder DefaultLayer()
    {
        return ZOrder.Polygons;
    }

    protected override Path CreateShape()
    {
        return new Path()
        {
            Data = new PathGeometry() { FillRule = FillRule.EvenOdd },
            StrokeJoin = PenLineJoin.Round
        };
    }

    /// <summary>Not while a dependency is no figure with an inside (a sector turned into an arc)</summary>
    public override void UpdateExistence()
    {
        base.UpdateExistence();
        if (Exists && (Dependencies.Count == 0 || Dependencies.Any(dependency => !(dependency is IRegion))))
        {
            Exists = false;
        }
    }

    public override void Recalculate()
    {
        boundary = Exists ? RegionBoundary.Trace(operation, Operands) : RegionBoundary.Empty;
    }

    public override void UpdateVisual()
    {
        if (!IsShown || Drawing == null)
        {
            return;
        }

        Shape.Data = boundary.CreateGeometry(Drawing.CoordinateSystem);
    }

    /// <summary>Inside where it is filled, and on its edge where that is drawn</summary>
    public override IFigure HitTest(Point point)
    {
        if (!Visible || !Exists || !boundary.HasInside)
        {
            return null;
        }

        if (HasFill && ContainsPoint(point))
        {
            return this;
        }

        if (DrawsOutline && boundary.DistanceTo(point) <= CursorTolerance + ToLogical(Shape.StrokeThickness) / 2)
        {
            return this;
        }

        return null;
    }

    bool HasFill => Shape.Fill != null && !(Shape.Fill is ISolidColorBrush brush && brush.Color.A == 0);

    bool DrawsOutline => Shape.StrokeThickness > 0
        && Shape.Stroke != null
        && !(Shape.Stroke is ISolidColorBrush brush && brush.Color.A == 0);

    public double Area => boundary.Area;

    public double Perimeter => boundary.Perimeter;

    #region The inside (IRegion)

    public bool HasInside => Exists && boundary.HasInside;

    public bool ContainsPoint(Point point)
    {
        return Exists && RegionBoundary.Contains(operation, Operands, point);
    }

    public Point NearestPointInside(Point point)
    {
        if (!HasInside)
        {
            return Math.InfinitePoint;
        }

        return ContainsPoint(point) ? point : boundary.NearestPoint(point);
    }

    /// <summary>The centroid of the set; the first operand's center while there is none</summary>
    public override Point Center
    {
        get
        {
            var centroid = boundary.Centroid;
            if (centroid.Exists())
            {
                return centroid;
            }

            return Dependencies.Count > 0 ? Dependencies[0].Center : new Point();
        }
    }

    #endregion

    public override void ReadXml(System.Xml.Linq.XElement element)
    {
        base.ReadXml(element);
        operation = System.Enum.TryParse(element.ReadString("Operation"), out RegionOperation read) ? read : RegionOperation.Difference;
    }

    public override void WriteXml(System.Xml.XmlWriter writer)
    {
        base.WriteXml(writer);
        writer.WriteAttributeString("Operation", operation.ToString());
    }
}
