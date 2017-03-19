using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace DynamicGeometry
{
    public abstract class ShapeCreator : FigureCreator
    {
        [PropertyGridName("Point by coordinates")]
        public class ShapeDialog
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
                if (e.Key == System.Windows.Input.Key.Enter)
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
                    double x = double.Parse(X, CultureInfo.CurrentUICulture);
                    double y = double.Parse(Y, CultureInfo.CurrentUICulture);

                    FreePoint first = (FreePoint)this.parent.FoundDependencies.FirstOrDefault(f => f is FreePoint);
                    if (first == null || !(first.X == x && first.Y == y))
                    {
                        this.parent.AddDependency(new Point(x, y));
                    }
                    else
                    {
                        if (this.parent.TempPoint != null)
                        {
                            this.parent.FoundDependencies.Remove(this.parent.TempPoint);
                        }
                        this.parent.AddFiguresAndRestart();
                    }
                }
            }

            [PropertyGridVisible]
            [PropertyGridName("Close figure")]
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
                if (ExpectingAPoint())
                {
                    return new ShapeDialog(this);
                }
                return null;
            }
        }
    }
}
