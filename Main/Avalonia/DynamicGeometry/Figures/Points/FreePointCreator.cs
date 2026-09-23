using System.ComponentModel;
using Avalonia.Input;

namespace DynamicGeometry
{
    [Category(BehaviorCategories.Points)]
    [Order(1)]
    public class FreePointCreator : Behavior
    {
        /// <summary>
        /// The tool's panel, only while "Point by coordinates" is switched on: the coordinates
        /// to type. Otherwise the tool has no panel (each kind of point gets its default style).
        /// </summary>
        [PropertyGridName("Point by coordinates")]
        public class CoordinatesDialog
        {
            public CoordinatesDialog(FreePointCreator parent)
            {
                this.parent = parent;
            }

            readonly FreePointCreator parent;

            [PropertyGridVisible]
            [PropertyGridFocus]
            [PropertyGridEvent("KeyDown", "X_KeyDown")]
            [PropertyGridName("X = ")]
            public string X { get; set; }

            [PropertyGridVisible]
            [PropertyGridEvent("KeyDown", "Y_KeyDown")]
            [PropertyGridName("Y = ")]
            public string Y { get; set; }

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

        CoordinatesDialog dialog;

        public override object PropertyBag
        {
            get
            {
                if (!Settings.Instance.EnablePointByCoordinates)
                {
                    return null;
                }

                dialog ??= new CoordinatesDialog(this);
                return dialog;
            }
        }

        public override void MouseDown(object sender, MouseButtonEventArgs e)
        {
            var placement = FindPointPlacement(e);
            if (placement.Kind == PointPlacementKind.Free)
            {
                CreatePointAtCurrentPosition(placement.Coordinates);
            }
            else if (placement.IsDependent)
            {
                Actions.Add(Drawing, placement.Create(Drawing));
            }
        }

        PointPlacement FindPointPlacement(MouseEventArgs e)
        {
            return PointPlacement.Find(Drawing, Coordinates(e), Settings.Instance.EnableSnapToCenter);
        }

        PointPlacement hoverPlacement;

        public override void MouseMove(object sender, MouseEventArgs e)
        {
            hoverPlacement = FindPointPlacement(e);
        }

        protected override PointPlacement GetClickPreview(MouseEventArgs e)
        {
            return hoverPlacement;
        }

        protected override Cursor GetCursor(Avalonia.Point coordinates)
        {
            // there already is a point here: a click doesn't put another one on top of it
            if (hoverPlacement != null && hoverPlacement.Kind == PointPlacementKind.Existing)
            {
                return ArrowCursor;
            }

            return GetCursor(hoverPlacement);
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