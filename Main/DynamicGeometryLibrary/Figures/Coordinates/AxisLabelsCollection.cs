using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace DynamicGeometry
{
    public class AxisLabel
    {
        public AxisLabel()
        {
            TextBlock = new TextBlock();
            TextBlock.Foreground = new SolidColorBrush(Axis.Color);
        }

        public TextBlock TextBlock { get; set; }

        public void PositionOnXAxis(double x, CoordinateSystem coordinateSystem)
        {
            SetLabelText(TextBlock, x);
            double width = TextBlock.ActualWidth;
            var coordinates = coordinateSystem
                .ToPhysical(new Point(x, 0))
                .OffsetX(-width / 2);
            if (x.EqualsWithPrecision(0))
            {
                coordinates = coordinates.OffsetX(-width / 2 - 2);
            }
            MoveLabel(TextBlock, coordinates);
        }

        public void PositionOnYAxis(double y, CoordinateSystem coordinateSystem)
        {
            SetLabelText(TextBlock, y);
            double width = TextBlock.ActualWidth;
            double height = TextBlock.ActualHeight;
            var coordinates = coordinateSystem
                .ToPhysical(new Point(0, y))
                .Plus(new Point(-width - 2, -height / 2));
            MoveLabel(TextBlock, coordinates);
        }

        private void MoveLabel(TextBlock label, Point coordinates)
        {
            Point old = TextBlock.GetCoordinates();
            if (old != coordinates)
            {
                TextBlock.MoveTo(coordinates);
            }
        }

        void SetLabelText(TextBlock label, double x)
        {
            string text = x.ToString(CultureInfo.InvariantCulture);
            if (label.Text != text)
            {
                label.Text = text;
            }
        }
    }

    public class AxisLabelRange
    {
        public AxisLabelRange(AxisLabelsCollection collection)
        {
            Collection = collection;
        }

        public AxisLabelsCollection Collection { get; set; }

        public Drawing Drawing
        {
            get
            {
                return Collection.Drawing;
            }
        }
        public CoordinateSystem CoordinateSystem
        {
            get
            {
                return Drawing.CoordinateSystem;
            }
        }

        public void OnAddingToCanvas(Canvas canvas)
        {
            var style = Collection.Style.GetWpfStyle();
            foreach (var item in List)
            {
                var textBlock = item.TextBlock;
                if (textBlock.Parent == null)
                {
                    canvas.Children.Add(textBlock);
                    textBlock.Apply(style); 
                }
            }
        }

        public void OnRemovingFromCanvas(Canvas canvas)
        {
            List.ForEach(t => canvas.Children.Remove(t.TextBlock));
        }

        private bool mVisible = true;
        public bool Visible
        {
            get
            {
                return mVisible;
            }
            set
            {
                mVisible = value;
                List.ForEach(t => t.TextBlock.Visibility = value ? Visibility.Visible : Visibility.Collapsed);
            }
        }

        List<AxisLabel> List = new List<AxisLabel>();

        public void AdjustToNewRange(IEnumerable<double> newPoints, bool isX)
        {
            int count = newPoints.Count();

            if (List.Count < count)
            {
                AddMissingElements(List, count - List.Count);
            }
            else if (List.Count > count)
            {
                RemoveExcessElements(List, List.Count - count);
            }

            int current = 0;
            foreach (var coordinate in newPoints)
            {
                if (isX)
                {
                    List[current++].PositionOnXAxis(coordinate, CoordinateSystem);
                }
                else
                {
                    List[current++].PositionOnYAxis(coordinate, CoordinateSystem);
                }
            }
        }

        void RemoveExcessElements(List<AxisLabel> List, int count)
        {
            var children = Drawing.Canvas.Children;
            for (int i = List.Count - count; i < List.Count; i++)
            {
                children.Remove(List[i].TextBlock);
            }
            List.RemoveRange(List.Count - count, count);
        }

        void AddMissingElements(List<AxisLabel> List, int count)
        {
            var style = Collection.Style.GetWpfStyle();
            var visible = this.Visible.ToVisibility();
            var children = Drawing.Canvas.Children;
            for (int i = 0; i < count; i++)
            {
                var newLabel = new AxisLabel();
                List.Add(newLabel);
                TextBlock textBlock = newLabel.TextBlock;
                textBlock.Visibility = visible;
                textBlock.Apply(style);
                children.Add(textBlock);
            }
        }
    }

    public class AxisLabelsCollection : FigureBase
    {
        public AxisLabelsCollection()
        {
            XAxisLabels = new AxisLabelRange(this);
            YAxisLabels = new AxisLabelRange(this);
        }

        AxisLabelRange XAxisLabels { get; set; }
        AxisLabelRange YAxisLabels { get; set; }

        public override void UpdateVisual()
        {
            var coordinateSystem = Drawing.CoordinateSystem;
            XAxisLabels.AdjustToNewRange(coordinateSystem.GetVisibleXPoints(), true);
            YAxisLabels.AdjustToNewRange(coordinateSystem.GetVisibleYPoints()
                .Where(y => y.Abs() > 0.001), false);
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
                if (XAxisLabels != null)
                {
                    XAxisLabels.Visible = value;
                }
                if (YAxisLabels != null)
                {
                    YAxisLabels.Visible = value;
                }
            }
        }

        public override IFigure HitTest(Point point)
        {
            return null;
        }

        public override void OnAddingToCanvas(Canvas newContainer)
        {
            XAxisLabels.OnAddingToCanvas(newContainer);
            YAxisLabels.OnAddingToCanvas(newContainer);
        }

        public override void OnRemovingFromCanvas(Canvas leavingContainer)
        {
            XAxisLabels.OnRemovingFromCanvas(leavingContainer);
            YAxisLabels.OnRemovingFromCanvas(leavingContainer);
        }

        public override void ApplyStyle()
        {

        }
    }
}
