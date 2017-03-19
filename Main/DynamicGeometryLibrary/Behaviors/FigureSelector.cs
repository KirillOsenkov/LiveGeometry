using System;
using System.Collections.Generic;
using System.Windows;

namespace DynamicGeometry
{
    [Ignore]
    public partial class FigureSelector : Behavior
    {
        public override Drawing Drawing
        {
            get
            {
                return mDrawing;
            }
            set
            {
                mDrawing = value;
                UpdateEnabledFigures();
            }
        }

        public override System.Windows.Controls.Canvas ParentCanvas
        {
            get
            {
                return Drawing.Canvas;
            }
            set
            {
                throw new NotImplementedException();
            }
        }

        public override void MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            var coordinates = Coordinates(e);
            var underMouse = Drawing.Figures.HitTest(coordinates);
            if (underMouse != null)
            {
                if (IsFigureSelected(underMouse))
                {
                    DeselectFigure(underMouse);
                }
                else
                {
                    TrySelectFigure(underMouse);
                }
            }
        }

        public void UpdateEnabledFigures()
        {
            foreach (var figure in Drawing.Figures)
            {
                var canSelect = CanSelectFigure(figure);
                if (canSelect != figure.Enabled)
                {
                    figure.Enabled = canSelect;
                }
            }
        }

        protected virtual void TrySelectFigure(IFigure figure)
        {
            if (!CanSelectFigure(figure))
            {
                return;
            }
            SelectFigure(figure);
        }

        protected virtual bool CanSelectFigure(IFigure figure)
        {
            return true;
        }

        public void SelectFigure(IFigure figure)
        {
            figure.Selected = true;
            UpdateEnabledFigures();
        }

        public void DeselectFigure(IFigure figure)
        {
            figure.Selected = false;
            UpdateEnabledFigures();
        }

        public bool IsFigureSelected(IFigure figure)
        {
            return figure.Selected;
        }

        public override FrameworkElement CreateIcon()
        {
            return IconBuilder.BuildIcon()
                .Point(0.5, 0.5)
                .Canvas;
        }

        public override string Name
        {
            get { return "Figure selector"; }
        }

        public IList<IFigure> GetSelection()
        {
            return new List<IFigure>(Drawing.GetSelectedFigures());
        }
    }
}
