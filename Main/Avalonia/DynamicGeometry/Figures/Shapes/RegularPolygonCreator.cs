using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Avalonia;

namespace DynamicGeometry
{
    /// <summary>
    /// A regular polygon by its center and a vertex, with as many sides as the tool's panel
    /// says (<see cref="SidesPanel"/>). Plain figures (<see cref="RegularPolygonBuilder"/>):
    /// points by coordinates for the vertices it adds, segments, a polygon. Right after it is
    /// made, its own panel (<see cref="RegularPolygonPanel"/>) changes its number of sides and
    /// its side; OK there brings this tool's panel back for the next polygon.
    /// </summary>
    [Category(BehaviorCategories.Shapes)]
    [Order(4)]
    public class RegularPolygonCreator : FigureCreator
    {
        /// <summary>How many sides the next regular polygon gets, by either regular polygon tool: the last count typed in any of their panels</summary>
        public static int NextNumberOfSides { get; set; } = 5;

        /// <summary>Whether the two clicks are two adjacent vertices (<see cref="RegularPolygonOnSideCreator"/>) and not the center and a vertex</summary>
        protected virtual bool OnSide => false;

        SidesPanel panel;

        // the first click of the polygon just made, for its panel: the clicks are gone by the
        // time it shows
        IPoint first;

        /// <summary>The side panel: how many sides the next polygon gets</summary>
        [PropertyGridNoUndo]
        public class SidesPanel
        {
            public SidesPanel(RegularPolygonCreator parent)
            {
                this.parent = parent;
            }

            readonly RegularPolygonCreator parent;

            // not [PropertyGridFocus]: the panel is back after every polygon, and the count
            // rarely changes; taking the keyboard each time would kill the tool letters
            [PropertyGridVisible]
            [PropertyGridName("Number of sides")]
            [Domain(RegularPolygonBuilder.MinimumSides, RegularPolygonBuilder.MaximumSides)]
            public int NumberOfSides
            {
                get
                {
                    return NextNumberOfSides;
                }
                set
                {
                    NextNumberOfSides = value;
                }
            }

            // the title of the panel
            public override string ToString()
            {
                return parent.Name;
            }
        }

        /// <summary>The panel for the number of sides, or the one for typed points while those are on (Coordinates tab)</summary>
        public override object PropertyBag
        {
            get
            {
                if (Settings.Instance.EnablePointByCoordinates && ExpectingAPoint())
                {
                    return base.PropertyBag;
                }

                if (panel == null)
                {
                    panel = new SidesPanel(this);
                }

                return panel;
            }
        }

        protected override IEnumerable<IFigure> CreateFigures()
        {
            first = (IPoint)FoundDependencies[0];
            return RegularPolygonBuilder.Create(Drawing, first, (IPoint)FoundDependencies[1], NextNumberOfSides, OnSide);
        }

        /// <summary>The polygon's own panel: its number of sides and its side (<see cref="RegularPolygonPanel"/>)</summary>
        protected override void ShowCreatedFigure(IList<IFigure> figures)
        {
            if (!(figures.LastOrDefault() is Polygon))
            {
                return;
            }

            var shown = new RegularPolygonPanel(first, NextNumberOfSides, OnSide, figures);
            Drawing.RaiseDisplayProperties(shown);
            Drawing.RaiseStatusNotification(shown.Hint);
        }

        protected override DependencyList InitExpectedDependencies()
        {
            return DependencyList.PointPoint;
        }

        public override string Name
        {
            get
            {
                return "Regular polygon";
            }
        }

        public override string HintText
        {
            get
            {
                return "Click the polygon center and then click a vertex.";
            }
        }

        public override FrameworkElement CreateIcon()
        {
            return IconBuilder.BuildIcon()
                .Polygon(
                    nameof(AppTheme.ShapeIconFill),
                    nameof(AppTheme.ShapeOutline),
                    new Point(0.68, 0.98),
                    new Point(1.01, 0.63),
                    new Point(0.875, 0.166),
                    new Point(0.401, 0.055),
                    new Point(0.068, 0.409),
                    new Point(0.208, 0.874))
                // yellow is what the tool's two clicks make and what can be dragged: the
                // center and one vertex; the other vertices follow
                .DependentPoint(1.01, 0.63)
                .DependentPoint(0.875, 0.166)
                .DependentPoint(0.401, 0.055)
                .DependentPoint(0.068, 0.409)
                .DependentPoint(0.208, 0.874)
                .Point(0.68, 0.98)
                .Point(0.54, 0.52)
                .Canvas;
        }
    }
}
