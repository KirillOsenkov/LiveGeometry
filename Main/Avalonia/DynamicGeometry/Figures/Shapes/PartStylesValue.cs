using System.Collections.Generic;
using System.Linq;

namespace DynamicGeometry;

/// <summary>
/// A row of a figure's grid that gives several of its parts one style (all the sides of a
/// regular polygon): it shows the style when they all have it and nothing when they differ,
/// and undo gives each part back the style it had, since <see cref="SetPropertyAction"/>
/// keeps one per part of a composite value.
/// </summary>
public class PartStylesValue : CompositeValueProvider
{
    readonly string name;
    readonly string caption;

    public PartStylesValue(string name, string caption, IEnumerable<IFigure> parts)
        : base(parts.Select(part => PropertyDiscoveryStrategy.CreateValueProvider(part, nameof(FigureBase.StyleDisplay))))
    {
        this.name = name;
        this.caption = caption;
    }

    public override string Name => name;

    public override string DisplayName => caption;

    // the parts are of one kind: the picker offers the styles of that kind
    public override object Parent => InnerList[0].Parent;
}
