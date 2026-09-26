using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Avalonia;
using Avalonia.Input;
using Avalonia.Media;

namespace DynamicGeometry
{
    /// <summary>
    /// Translation, one thing per step: the figure to translate, then the magnitude, then
    /// the direction. Each of the two is a figure clicked on the canvas (a segment or
    /// anything with a length, an angle; a vector gives both at once), a number typed in
    /// the side panel and confirmed with OK, or - when the source is a single point - left
    /// Free, so that the point can be dragged; then one more click places it. The panel
    /// shows only the step at hand and goes away when the point is made.
    /// </summary>
    [Category(BehaviorCategories.Transform)]
    [Order(3)]
    public class TranslationCreator : FigureCreator
    {
        enum Step
        {
            Source,
            Magnitude,
            Direction,
            Placement
        }

        Step step;
        IFigure magnitudeSource;
        IFigure directionSource;
        bool magnitudeFree;
        bool directionFree;
        double typedMagnitude;
        double typedDirection;
        Point? placement;
        ValueStep panel;

        // the next translation is likely the same as the last
        static double lastMagnitude = 1;
        static double lastDirection = 0;

        #region The panel of a step

        /// <summary>
        /// The side panel while a value is wanted: the number box, OK, and Free when the
        /// point may be left draggable. Enter is OK.
        /// </summary>
        public class ValueStep : ICustomMethodProvider, IConditionalProperties
        {
            public ValueStep(TranslationCreator parent, bool isDirection, bool canFree, double value)
            {
                this.parent = parent;
                this.isDirection = isDirection;
                this.canFree = canFree;
                Value = value;
            }

            readonly TranslationCreator parent;
            readonly bool isDirection;
            readonly bool canFree;

            [PropertyGridVisible]
            [PropertyGridFocus]
            [PropertyGridPreferredEditor("UpDown")]
            [PropertyGridCustomValueProvider(typeof(ConditionalPropertyValue))]
            [PropertyGridEvent("KeyDown", "Value_KeyDown")]
            public double Value { get; set; }

            public void Value_KeyDown(object sender, KeyEventArgs e)
            {
                if (e.Key == Key.Enter)
                {
                    OK();
                    e.Handled = true;
                }
            }

            [PropertyGridVisible]
            public void OK()
            {
                parent.Accept(Value);
            }

            [PropertyGridVisible]
            public void Free()
            {
                parent.LeaveFree();
            }

            public IEnumerable<IOperationDescription> GetMethods()
            {
                yield return MethodDescription.Get<ValueStep>("OK");
                if (canFree)
                {
                    yield return MethodDescription.Get<ValueStep>("Free");
                }
            }

            public bool CanEdit(string propertyName)
            {
                return true;
            }

            public string Caption(string propertyName, string defaultCaption)
            {
                return isDirection ? "Direction (degrees)" : "Magnitude";
            }

            // the title of the panel
            public override string ToString()
            {
                return isDirection ? "Translation: direction" : "Translation: magnitude";
            }
        }

        public override object PropertyBag
        {
            get { return panel; }
        }

        #endregion

        #region Steps

        public override void Started()
        {
            base.Started();
            step = Step.Source;
            magnitudeSource = null;
            directionSource = null;
            magnitudeFree = false;
            directionFree = false;
            placement = null;
            panel = null;
        }

        // before the base raises "construction complete", which is what shows the panel again
        public override void Stopping()
        {
            panel = null;
            base.Stopping();
        }

        IFigure Source
        {
            get { return FoundDependencies.Count > 0 ? FoundDependencies[0] : null; }
        }

        /// <summary>Only a point can be left partly free: a figure's points would each get a freedom of their own</summary>
        bool CanFree
        {
            get { return Source is IPoint; }
        }

        protected override void Click(Point coordinates)
        {
            var underMouse = FindFigureToPick(ClickedUnconstrainedCoordinates);
            switch (step)
            {
                case Step.Source:
                    if (underMouse != null)
                    {
                        Drawing.RaiseConstructionStepStarted();
                        FoundDependencies.Add(underMouse);
                        Advance(Step.Magnitude);
                    }

                    break;

                case Step.Magnitude:
                    if (underMouse is Vector)
                    {
                        magnitudeSource = underMouse;
                        directionSource = underMouse;
                        Finish();
                    }
                    else if (underMouse != null)
                    {
                        magnitudeSource = underMouse;
                        Advance(Step.Direction);
                    }

                    break;

                case Step.Direction:
                    if (underMouse != null)
                    {
                        directionSource = underMouse;
                        Finish();
                    }

                    break;

                case Step.Placement:
                    placement = coordinates;
                    AddFiguresAndRestart();
                    break;
            }
        }

        /// <summary>OK in the panel: the typed value is the magnitude or the direction of this step</summary>
        void Accept(double value)
        {
            if (step == Step.Magnitude)
            {
                typedMagnitude = value;
                lastMagnitude = value;
                Advance(Step.Direction);
            }
            else if (step == Step.Direction)
            {
                typedDirection = value;
                lastDirection = value;
                Finish();
            }
        }

        /// <summary>Free in the panel: this step's quantity is the one dragging will change</summary>
        void LeaveFree()
        {
            if (step == Step.Magnitude)
            {
                magnitudeFree = true;
                Advance(Step.Direction);
            }
            else if (step == Step.Direction)
            {
                directionFree = true;
                Finish();
            }
        }

        void Advance(Step next)
        {
            step = next;
            switch (next)
            {
                case Step.Magnitude:
                    panel = new ValueStep(this, isDirection: false, canFree: CanFree, value: lastMagnitude);
                    break;
                case Step.Direction:
                    // one freedom at most: a point free in both would just follow the mouse
                    panel = new ValueStep(this, isDirection: true, canFree: CanFree && !magnitudeFree, value: lastDirection);
                    break;
                default:
                    panel = null;
                    break;
            }

            AdvertiseNextDependency();
        }

        /// <summary>Both values known: the point is made, or placed by one more click when something is free</summary>
        void Finish()
        {
            if (magnitudeFree || directionFree)
            {
                Advance(Step.Placement);
                CreateTempResults();
            }
            else
            {
                AddFiguresAndRestart();
            }
        }

        public override void MouseMove(object sender, MouseEventArgs e)
        {
            base.MouseMove(sender, e);
            if (step == Step.Placement)
            {
                // the point rides the cursor along its circle or line
                var riding = TempResults.OfType<TranslatedPoint>().LastOrDefault();
                if (riding != null)
                {
                    riding.MoveTo(Coordinates(e));
                }
            }
        }

        #endregion

        #region What a click may pick

        bool Accepts(IFigure figure)
        {
            switch (step)
            {
                case Step.Source:
                    return Transformer.CanBeTransformSource(figure);
                case Step.Magnitude:
                    return figure is Vector || figure is ILengthProvider;
                case Step.Direction:
                    return figure is IAngleProvider;
                default:
                    return false;
            }
        }

        protected override IFigure FindFigureToPick(Point unconstrainedCoordinates)
        {
            var figure = Drawing.Figures.HitTest(unconstrainedCoordinates);
            return figure != null && Accepts(figure) ? figure : null;
        }

        protected override Type GetExpectedDependencyType()
        {
            switch (step)
            {
                case Step.Source:
                    return typeof(IFigure);
                case Step.Magnitude:
                    return typeof(ILengthProvider);
                case Step.Direction:
                    return typeof(IAngleProvider);
                default:
                    return typeof(IPoint);
            }
        }

        protected override bool ExpectingAPoint()
        {
            return false;
        }

        protected override DependencyList InitExpectedDependencies()
        {
            return DependencyList.Create<IFigure>();
        }

        #endregion

        /// <summary>
        /// A typed value becomes a Number of its own, so that it can be edited later, and
        /// goes into the drawing before the points. The sources are shared by every point
        /// of a translated figure.
        /// </summary>
        protected override IEnumerable<IFigure> CreateFigures()
        {
            var source = Source;
            Check.NotNull(source);
            var magnitude = magnitudeSource;
            var direction = directionSource;
            if (magnitude == null && !magnitudeFree)
            {
                magnitude = Number.CreateAuxiliary(Drawing, typedMagnitude);
                yield return magnitude;
            }

            if (direction == null && !directionFree)
            {
                direction = Number.CreateAuxiliary(Drawing, typedDirection);
                yield return direction;
            }

            var results = Transformer.CreateTranslatedFigure(Drawing, source, magnitude, direction);
            Check.NoNullElements(results);
            if (placement != null && results.LastOrDefault() is TranslatedPoint placed)
            {
                placed.MoveTo(placement.Value);
            }

            foreach (IFigure f in results)
            {
                yield return f;
            }
        }

        public override string Name
        {
            get { return "Translation"; }
        }

        public override string HintText
        {
            get { return "Click the figure to translate."; }
        }

        public override string ConstructionHintText(Drawing.ConstructionStepCompleteEventArgs args)
        {
            switch (step)
            {
                case Step.Magnitude:
                    return "Click a segment or a vector for the magnitude, or type it and press OK."
                        + (CanFree ? " Free leaves it to be dragged." : "");
                case Step.Direction:
                    return "Click a figure with an angle for the direction, or type it and press OK."
                        + (CanFree && !magnitudeFree ? " Free leaves it to be dragged." : "");
                case Step.Placement:
                    return "Click where the point goes.";
                default:
                    return HintText;
            }
        }

        public override FrameworkElement CreateIcon()
        {
            return IconBuilder.BuildIcon()
                .Polygon(
                    new SolidColorBrush(Colors.Yellow),
                    new SolidColorBrush(Colors.Black),
                    new Point(0.1, 0.9),
                    new Point(0.4, 0.9),
                    new Point(0.4, 0.6),
                    new Point(0.1, 0.6))
                .Polygon(
                    new SolidColorBrush(Color.FromArgb(255, 128, 255, 128)),
                    new SolidColorBrush(Colors.Black),
                    new Point(0.6, 0.4),
                    new Point(0.9, 0.4),
                    new Point(0.9, 0.1),
                    new Point(0.6, 0.1))
                .Canvas;
        }
    }
}
