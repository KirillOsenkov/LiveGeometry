using System.ComponentModel;
using System.Linq;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Shapes;
using System.Windows;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Points)]
    [Order(1)]
    public class FreePointCreator : Behavior
    {
        [PropertyGridName("Point by coordinates")]
        public class Dialog
        {
            public Dialog(FreePointCreator parent)
            {
                this.parent = parent;
                this.style = parent.Drawing.StyleManager.GetStyles<PointStyle>().FirstOrDefault();
            }

            FreePointCreator parent;

            [PropertyGridVisible]
            [PropertyGridFocus]
            [PropertyGridEvent("KeyDown", "X_KeyDown")]
            [PropertyGridName("X = ")]
            public string X { get; set; }

            [PropertyGridVisible]
            [PropertyGridEvent("KeyDown", "Y_KeyDown")]
            [PropertyGridName("Y = ")]
            public string Y { get; set; }

            private IFigureStyle style;
            [PropertyGridVisible]
            public IFigureStyle Style
            {
                get
                {
                    return this.style;
                }
                set
                {
                    this.style = value;
                    Canvas canvas = this.parent.Icon as Canvas;
                    var pointShape = canvas.Children[0] as Shape;
                    pointShape.Apply(style.GetWpfStyle());
                }
            }

            internal void X_KeyDown(object sender, KeyEventArgs e)
            {
                Common_KeyDown(sender, e);
                if (e.Handled)
                {
                    return;
                }
            }

            internal void Y_KeyDown(object sender, KeyEventArgs e)
            {
                Common_KeyDown(sender, e);
                if (e.Handled)
                {
                    return;
                }
            }

            internal void Common_KeyDown(object sender, KeyEventArgs e)
            {
                if (e.Key == System.Windows.Input.Key.Escape)
                {
                    e.Handled = true;
                }
                else if (e.Key == System.Windows.Input.Key.Enter)
                {
                    AddPoint();
                    e.Handled = true;
                }
            }

            [PropertyGridVisible]
            [PropertyGridName("Add point")]
            public void AddPoint()
            {
                var xresult = parent.Drawing.CompileExpression(X);
                var yresult = parent.Drawing.CompileExpression(Y);

                if (xresult.IsSuccess && yresult.IsSuccess)
                {
                    var point = Factory.CreatePointByCoordinates(parent.Drawing, X, Y);
                    Actions.Add(parent.Drawing, point);
                }
            }
        }

        Dialog dialog;

        public override object PropertyBag
        {
            get
            {
                if (dialog == null)
                {
                    dialog = new Dialog(this);
                }
                return dialog;
            }
        }

        public override void MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            var coordinates = Coordinates(e);
            var list = Drawing.Figures.HitTestMany(coordinates);
            var figureList = list.Where(f => f is ILinearFigure).ToArray();

            IFigure created = null;

            if (!figureList.IsEmpty())
            {
                if (figureList.Length == 2
                    && IntersectionAlgorithms.CanIntersect(figureList[0], figureList[1]))
                {
                    created = Factory.CreateIntersectionPoint(
                        Drawing,
                        figureList[0],
                        figureList[1],
                        coordinates);
                    Actions.Add(Drawing, created);
                }
                else if (figureList.Length == 1 && PointOnFigure.CanBeOnFigure(figureList[0]))
                {
                    created = Factory.CreatePointOnFigure(
                        Drawing,
                        figureList[0],
                        coordinates);
                    Actions.Add(Drawing, created);
                }
            }
            else
            {
                created = CreatePointAtCurrentPosition(coordinates);
            }

            if (created != null && dialog != null && dialog.Style != null)
            {
                created.Style = dialog.Style;
            }
        }

        public override FrameworkElement CreateIcon()
        {
            return IconBuilder.BuildIcon()
                .Point(0.5, 0.5)
                .Canvas;
        }

        public override string Name
        {
            get { return "Point"; }
        }

        public override string HintText
        {
            get { return "Click to create a point. You can also click on a figure or an intersection."; }
        }
    }
}