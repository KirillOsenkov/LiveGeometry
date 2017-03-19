using System;
using System.Windows.Media;
namespace DynamicGeometry
{
    public abstract class FigureCreator : Behavior
    {
        public FigureCreator()
        {
            FoundDependencies = new FigureList();
            ExpectedDependencies = InitExpectedDependencies();
        }

        protected override void Reset()
        {
            base.Reset();
            FoundDependencies = new FigureList();
            RemoveTempPoint();
        }

        protected ExpectedDependencyList ExpectedDependencies { get; private set; }
        protected FigureList FoundDependencies { get; set; }

        protected Type ExpectedDependency
        {
            get
            {
                if (FoundDependencies.Count < ExpectedDependencies.Count)
                {
                    return ExpectedDependencies[FoundDependencies.Count];
                }
                return null;
            }
        }

        protected abstract ExpectedDependencyList InitExpectedDependencies();

        protected virtual void AddFoundDependency(IFigure figure)
        {
            if (ExpectedDependency.IsAssignableFrom(figure.GetType()))
            {
                FoundDependencies.Add(figure);
            }
        }

        public FreePoint TempPoint { get; set; }

        protected override void MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.RightButton == System.Windows.Input.MouseButtonState.Pressed)
            {
                AbortAndSetDefaultTool();
                return;
            }

            RemoveTempPoint();

            var underMouse = Drawing.Figures.HitTest(Coordinates(e));
            if (underMouse == null && ExpectingAPoint())
            {
                underMouse = CreatePointAtCurrentPosition(e);
            }

            AddFoundDependency(underMouse);

            if (FoundDependencies.Count < ExpectedDependencies.Count)
            {
                if (ExpectingAPoint())
                {
                    TempPoint = CreateTempPoint(e);
                    AddFoundDependency(TempPoint);
                    if (FoundDependencies.Count == ExpectedDependencies.Count)
                    {
                        CreateAndAddFigure();
                    }
                }
            }
            else
            {
                Finish(e);
            }
        }

        protected override void MouseMove(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (TempPoint != null)
            {
                TempPoint.MoveTo(Coordinates(e));
                Drawing.Figures.Recalculate();
            }
        }

        private FreePoint CreateTempPoint(System.Windows.Input.MouseButtonEventArgs e)
        {
            var result = CreatePointAtCurrentPosition(e);
            result.Shape.Fill = new SolidColorBrush(Colors.LightGreen);
            return result;
        }

        private LinearGradientBrush gradient = new LinearGradientBrush(Colors.Yellow, Colors.White, 10);
        private void Finish(System.Windows.Input.MouseButtonEventArgs e)
        {
            if (TempPoint != null)
            {
                TempPoint.Shape.Fill = new SolidColorBrush(Colors.LightYellow);
                TempPoint = null;
            }
            Reset();
        }

        protected abstract IFigure CreateFigure();

        protected void CreateAndAddFigure()
        {
            IFigure result = CreateFigure();
            Drawing.Figures.Add(result);
        }

        private bool ExpectingAPoint()
        {
            return typeof(IPoint).IsAssignableFrom(ExpectedDependency);
        }

        private void RemoveTempPoint()
        {
            if (TempPoint != null)
            {
                Drawing.Figures.Remove(TempPoint);
                FoundDependencies.Remove(TempPoint);
            }
        }
    }
}