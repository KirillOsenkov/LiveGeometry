using System.Collections.Generic;
using System.Linq;
using Avalonia.Controls;
using Avalonia.Controls.Shapes;

namespace DynamicGeometry
{
    public class RectangularGridLinesCollection : GridLinesCollection
    {
        List<Line> Lines = new List<Line>();
        List<Line> MinorLines = new List<Line>();

        /// <summary>The finer lines between the labeled ones, fainter than <see cref="FigureBase.Style"/></summary>
        public LineStyle MinorStyle { get; set; }

        IEnumerable<Line> AllLines
        {
            get
            {
                return Lines.Concat(MinorLines);
            }
        }

        public override bool Visible
        {
            get
            {
                return base.Visible;
            }
            set
            {
                base.Visible = value;
                foreach (var line in AllLines)
                {
                    line.Visibility = value ? Visibility.Visible : Visibility.Collapsed;
                }
            }
        }

        public override void UpdateVisual()
        {
            CoordinateSystem coordinateSystem = Drawing.CoordinateSystem;
            Place(
                Lines,
                Style,
                coordinateSystem.GetVisibleXPoints(),
                coordinateSystem.GetVisibleYPoints(),
                coordinateSystem);
            Place(
                MinorLines,
                MinorStyle,
                coordinateSystem.GetMinorXPoints(),
                coordinateSystem.GetMinorYPoints(),
                coordinateSystem);
        }

        void Place(
            List<Line> lines,
            IFigureStyle style,
            IEnumerable<double> xPoints,
            IEnumerable<double> yPoints,
            CoordinateSystem coordinateSystem)
        {
            int count = xPoints.Count() + yPoints.Count();
            if (lines.Count < count)
            {
                AddNewLines(lines, style, count - lines.Count);
            }
            else if (lines.Count > count)
            {
                RemoveExcessLines(lines, lines.Count - count);
            }

            int i = 0;
            foreach (var x in xPoints)
            {
                MoveLineX(lines[i++], x, coordinateSystem);
            }
            foreach (var y in yPoints)
            {
                MoveLineY(lines[i++], y, coordinateSystem);
            }
        }

        public override void OnAddingToCanvas(Canvas newContainer)
        {
            foreach (var line in AllLines)
            {
                if (line.Parent == null)
                {
                    newContainer.Children.Add(line);
                }
            }
        }

        public override void OnRemovingFromCanvas(Canvas leavingContainer)
        {
            foreach (var line in AllLines)
            {
                leavingContainer.Children.Remove(line);
            }
        }

        void MoveLineX(Line line, double x, CoordinateSystem coordinateSystem)
        {
            line.Move(
                x,
                coordinateSystem.MinimalVisibleY,
                x,
                coordinateSystem.MaximalVisibleY,
                coordinateSystem);
        }

        void MoveLineY(Line line, double y, CoordinateSystem coordinateSystem)
        {
            line.Move(
                coordinateSystem.MinimalVisibleX,
                y,
                coordinateSystem.MaximalVisibleX,
                y,
                coordinateSystem);
        }

        void RemoveExcessLines(List<Line> lines, int count)
        {
            for (int i = lines.Count - count; i < lines.Count; i++)
            {
                Drawing.Canvas.Children.Remove(lines[i]);
            }
            lines.RemoveRange(lines.Count - count, count);
        }

        void AddNewLines(List<Line> lines, IFigureStyle style, int count)
        {
            for (int i = 0; i < count; i++)
            {
                var newLine = new Line()
                {
                    Visibility = this.Visible.ToVisibility()
                };
                newLine.ZIndex = (int)ZOrder.Grid;
                lines.Add(newLine);
                Drawing.Canvas.Children.Add(newLine);
                newLine.Apply(style.GetWpfStyle());
            }
        }

        public override void ApplyStyle()
        {
        }
    }
}
