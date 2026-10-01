using System.Linq;
using Avalonia;
using Avalonia.Input;

namespace DynamicGeometry
{
    public abstract class ShapeCreator : FigureCreator
    {
        [PropertyGridName("Point by coordinates")]
        public class ShapeDialog : ToolPanel
        {
            public ShapeDialog(ShapeCreator parent)
            {
                this.parent = parent;
            }

            ShapeCreator parent;

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
                if (e.Key == Avalonia.Input.Key.Enter)
                {
                    AddPoint();
                    e.Handled = true;
                }
            }

            [PropertyGridVisible]
            [PropertyGridName("Add point")]
            [PropertyGridIcon(PropertyGridIcon.Plus)]
            public void AddPoint()
            {
                var xresult = Compile(parent.Drawing, nameof(X), X);
                var yresult = Compile(parent.Drawing, nameof(Y), Y);

                if (xresult.IsSuccess && yresult.IsSuccess)
                {
                    // an expression (A.X + 3) gives its value now: the vertex is free
                    double x = Evaluate(nameof(X), xresult);
                    double y = Evaluate(nameof(Y), yresult);
                    if (!new Point(x, y).Exists())
                    {
                        return;
                    }

                    // the coordinates of the first vertex again close the figure, when it has
                    // vertices enough; a click on the first vertex does the same
                    FreePoint first = this.parent.FoundDependencies.FirstOrDefault() as FreePoint;
                    bool isFirstAgain = first != null && first != this.parent.TempPoint && first.X == x && first.Y == y;
                    if (!isFirstAgain)
                    {
                        this.parent.AddTypedPoint(new Point(x, y));
                    }
                    else
                    {
                        CloseFigure();
                    }
                }
            }

            [PropertyGridVisible]
            [PropertyGridName("Close figure")]
            [PropertyGridIcon(PropertyGridIcon.Check)]
            public void CloseFigure()
            {
                if (this.parent.FoundDependencies.Count > 3)
                {
                    if (this.parent.TempPoint != null)
                    {
                        this.parent.FoundDependencies.Remove(this.parent.TempPoint);
                    }
                    this.parent.AddFiguresAndRestart();
                }
            }
        }

        public override object PropertyBag
        {
            get
            {
                if (Settings.Instance.EnablePointByCoordinates && ExpectingAPoint())
                {
                    return new ShapeDialog(this);
                }
                return null;
            }
        }
    }
}
