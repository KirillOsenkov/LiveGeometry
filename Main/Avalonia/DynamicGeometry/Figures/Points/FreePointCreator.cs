using System.ComponentModel;
using System.Linq;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Controls.Shapes;

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
                if (e.Key == Avalonia.Input.Key.Escape)
                {
                    e.Handled = true;
                }
                else if (e.Key == Avalonia.Input.Key.Enter)
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

        public override void MouseDown(object sender, MouseButtonEventArgs e)
        {
            var placement = FindPointPlacement(e);
            IFigure created = null;

            if (placement.Kind == PointPlacementKind.Free)
            {
                created = CreatePointAtCurrentPosition(placement.Coordinates);
            }
            else if (placement.IsDependent)
            {
                created = placement.Create(Drawing);
                Actions.Add(Drawing, created);
            }

            if (created != null && dialog != null && dialog.Style != null)
            {
                created.Style = dialog.Style;
            }
        }

        PointPlacement FindPointPlacement(MouseEventArgs e)
        {
            return PointPlacement.Find(Drawing, Coordinates(e), Settings.Instance.EnableSnapToCenter);
        }

        protected override PointPlacement GetClickPreview(MouseEventArgs e)
        {
            return FindPointPlacement(e);
        }

        protected override IFigureStyle ClickPreviewPointStyle
        {
            get
            {
                if (dialog != null && dialog.Style != null)
                {
                    return dialog.Style;
                }

                return base.ClickPreviewPointStyle;
            }
        }

        // there already is a point here: a click doesn't put another one on top of it
        protected override Cursor GetCursor(Avalonia.Point coordinates)
        {
            return Drawing.Figures.HitTest<IPoint>(coordinates) != null ? ArrowCursor : CrossCursor;
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