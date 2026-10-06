using System;
using System.Collections.Generic;
using System.Linq;

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

        public override Avalonia.Controls.Canvas ParentCanvas
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

        public override void MouseDown(object sender, MouseButtonEventArgs e)
        {
            Toggle(Choice.Pick(FindFiguresToToggle(Coordinates(e))));
        }

        /// <summary>Selects the figure, or lets go of it when this selector selected it</summary>
        public void Toggle(IFigure figure)
        {
            if (figure == null)
            {
                return;
            }

            if (clicked.Contains(figure))
            {
                DeselectFigure(figure);
            }
            else
            {
                TrySelectFigure(figure);
            }
        }

        /// <summary>
        /// The figures a click here could select or let go of, the one it takes first: those
        /// this selector selected, and those it can select. Not one an earlier step selected
        /// (an input, while the results are picked), which a click used to unselect to no
        /// purpose; nor an axis.
        /// </summary>
        public IReadOnlyList<IFigure> FindFiguresToToggle(Avalonia.Point coordinates)
        {
            return Drawing.Figures.HitTestAll(
                coordinates,
                figure => figure.Visible
                    && figure.IsHitTestVisible
                    && !(figure is AxisLine)
                    && (clicked.Contains(figure) || CanSelectFigure(figure)));
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

        // the figures selected here, in the order they were clicked
        readonly List<IFigure> clicked = new List<IFigure>();

        public void SelectFigure(IFigure figure)
        {
            figure.Selected = true;
            clicked.Remove(figure);
            clicked.Add(figure);
            UpdateEnabledFigures();
        }

        public void DeselectFigure(IFigure figure)
        {
            figure.Selected = false;
            clicked.Remove(figure);
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

        /// <summary>
        /// What this selector selected, in the order of the clicks: the order a tool defined
        /// from them asks for its inputs. Not the drawing's selection, which lists them in the
        /// drawing's order and still holds what an earlier step selected (the inputs, while
        /// the results are picked).
        /// </summary>
        public IList<IFigure> GetSelection()
        {
            return clicked.Where(f => f.Selected).ToList();
        }
    }
}
