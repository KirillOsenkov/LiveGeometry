using System;
using System.Collections.Generic;
using System.Linq;
using Avalonia;
using Avalonia.Input;
using GuiLabs.Undo;

namespace DynamicGeometry
{
    public abstract partial class FigureCreator : Behavior
    {
        #region Dialog

        [PropertyGridName("Point by coordinates")]
        [PropertyGridNoUndo]
        public class Dialog : ToolPanel
        {
            public Dialog(FigureCreator parent)
            {
                this.parent = parent;
            }

            FigureCreator parent;

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
                if (e.Key == Key.Enter)
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
                    // an expression (A.X + 3) gives its value now: the point is free, not tied to A
                    var point = new Point(Evaluate(nameof(X), xresult), Evaluate(nameof(Y), yresult));
                    if (!point.Exists())
                    {
                        return;
                    }

                    this.parent.AddTypedPoint(point);
                }
            }
        }

        /// <summary>
        /// A point at typed coordinates is the next step of the construction: a free point
        /// exactly there, whatever happens to pass through - or the point that is there
        /// already. Not what was under the last mouse click, which is where the tool would
        /// look otherwise (the polygon tools did: after a click on the first vertex, a typed
        /// second one found the first again and was dropped).
        /// </summary>
        protected void AddTypedPoint(Point point)
        {
            ClickedUnconstrainedCoordinates = point;
            canPlacePointsOnFigures = false;
            try
            {
                AddDependency(point);
            }
            finally
            {
                canPlacePointsOnFigures = true;
            }
        }

        public override object PropertyBag
        {
            get
            {
                if (Settings.Instance.EnablePointByCoordinates && ExpectingAPoint())
                {
                    return new Dialog(this);
                }
                return null;
            }
        }

        #endregion

        #region Behavior initialize and cleanup

        /// <summary>
        /// Transaction is necessary if, for example, you're adding a segment and its two endpoints in one click-drag-release motion.
        /// Both points and the segment will be created, and we want Undo to remove both the points and the segment in one swoop.
        /// It spans one construction: opened by <see cref="StartConstruction"/> at the first
        /// step, committed when the figures are added. While the tool waits for the next
        /// construction there is none, so an edit made then (in the panel that follows a new
        /// segment, say) is an undo step of its own and not part of the next figure. (It used
        /// to be opened when the tool started, which lumped those edits into the next figure.)
        /// </summary>
        protected Transaction Transaction { get; set; }

        protected bool ConstructionComplete { get; set; }

        public override void Started()
        {
            this.ConstructionComplete = true;
            ExpectedDependencies = InitExpectedDependencies();
            FoundDependencies.Clear();
            hoverPlacement = null;
            hoverFigure = null;
        }

        /// <summary>
        /// The first step of a construction: from here on what the tool records is one undo
        /// step, until the figures are added or the construction is abandoned.
        /// </summary>
        protected void StartConstruction()
        {
            ForgetCreatedFigure();
            EnsureTransaction();
            Drawing.RaiseConstructionStepStarted();
        }

        /// <summary>Before anything of a construction is recorded - the point a click makes comes first</summary>
        protected void EnsureTransaction()
        {
            if (Transaction == null)
            {
                Transaction = Transaction.Create(Drawing.ActionManager, false);
            }
        }

        /// <summary>The transaction opened for a click that turned out to add nothing, at the first step</summary>
        protected void DropEmptyTransaction()
        {
            if (Transaction != null && FoundDependencies.IsEmpty() && !Transaction.HasActions())
            {
                Transaction.Rollback();
                Transaction = null;
            }
        }

        public override void Stopping()
        {
            ForgetCreatedFigure();
            if (Transaction != null)
            {
                if (this.ConstructionComplete)
                {
                    Transaction.Dispose();  // Changes in Property Grid
                }
                else
                {
                    Transaction.Rollback(); // Incomplete constructions
                }
                Transaction = null;
            }

            // Raise this is necessary to enable/disable undo/redo properly. - D.H.
            Drawing.RaiseConstructionStepComplete(new Drawing.ConstructionStepCompleteEventArgs()
            {
                ConstructionComplete = true
            });

            Drawing.Figures.EnableAll();

            RemoveTempPointIfNecessary();
            RemoveTempResultsIfNecessary();
            RemoveIntermediateFigureIfNecessary();
        }

        protected virtual void AddFiguresAndRestart()
        {
            // a tool that goes straight from a click to its figures (Distance on a segment)
            // has had no first step
            EnsureTransaction();
            RemoveTempResultsIfNecessary();
            var figures = CreateFigures().ToList();
            foreach (var figure in figures)
            {
                if (figure != null)
                {
                    Actions.Add(Drawing, figure);
                }
            }

            FiguresAdded(figures);
            Drawing.RaiseUserIsAddingFigures(new Drawing.UIAFEventArgs() { Figures = figures });
            Transaction.Commit();
            Transaction = null;
            this.ConstructionComplete = true;
            Drawing.RaiseConstructionStepComplete(new Drawing.ConstructionStepCompleteEventArgs()
            {
                ConstructionComplete = true
            });
            Restart();
            ShowCreatedFigure(figures);
        }

        /// <summary>The figures made are in the drawing, inside the construction's transaction</summary>
        protected virtual void FiguresAdded(IList<IFigure> figures)
        {
        }

        /// <summary>
        /// Right after a figure is made, the side panel offers what is worth adjusting at
        /// once, and the status bar says so - instead of a box for it on every tool: its tied
        /// values (<see cref="TiedValuesPanel"/> - the angle of a line at an angle or of a
        /// rotation, the factor of a dilation, the distance and direction of a translation,
        /// see <see cref="ITiedValues"/>), or the length and Fix length of a figure with one
        /// (<see cref="LengthPanel"/>: a segment, a vector, a square's base side, a regular
        /// polygon).
        /// </summary>
        protected virtual void ShowCreatedFigure(IList<IFigure> figures)
        {
            var values = figures.OfType<ITiedValues>().FirstOrDefault(v => v.TiedValueNames.Any());
            if (values != null)
            {
                ShowTiedValues(figures.Last(f => f != null), values);
                return;
            }

            var withLength = figures.OfType<IFixableLength>().FirstOrDefault();
            // nothing to offer when no point can move (built on existing dependent points)
            if (withLength == null || !withLength.CanEdit("Length"))
            {
                return;
            }

            var what = withLength.Caption("Length", "Length").ToLowerInvariant();
            Drawing.RaiseDisplayProperties(new LengthPanel(withLength));
            Drawing.RaiseStatusNotification(withLength.Title + ": set its " + what + " in the panel, or fix it.");
        }

        #endregion

        #region The figure just made

        /// <summary>
        /// The figure whose panel (<see cref="TiedValuesPanel"/>) is up right after it was made:
        /// a click on a suitable figure ties one of its values to that
        /// (<see cref="TryTieCreatedFigure"/>) instead of starting the next construction. Until
        /// anything else takes the side panel, the next construction starts, or the tool is
        /// put down.
        /// </summary>
        protected IFigure CreatedFigure { get; private set; }

        ITiedValues createdValues;
        Drawing createdIn;

        // under the cursor: the figure a click would tie a value to
        IFigure tieTarget;

        void ShowTiedValues(IFigure shown, ITiedValues values)
        {
            ForgetCreatedFigure();
            CreatedFigure = shown;
            createdValues = values;
            createdIn = Drawing;
            createdIn.DisplayProperties += Drawing_DisplayProperties;
            Drawing.RaiseDisplayProperties(new TiedValuesPanel(shown, values));
            Drawing.RaiseStatusNotification(shown.Title + ": " + CreatedFigureHint(values));
        }

        /// <summary>What the status bar says under the panel of a figure just made; a tool knows what its values can be taken from</summary>
        protected virtual string CreatedFigureHint(ITiedValues values)
        {
            var names = values.TiedValueNames.Select(name => name.ToLowerInvariant()).ToList();
            string what = names.Count == 1
                ? names[0]
                : string.Join(", ", names.Take(names.Count - 1)) + " and " + names[names.Count - 1];
            return "set its " + what + " in the panel, or click a figure that has one to take it from there.";
        }

        /// <summary>
        /// The offer is over. The values as they are become the defaults of the next
        /// construction (<see cref="TakeDefaultsFrom"/>), so that the next figure is made
        /// like the last one ended up.
        /// </summary>
        protected void ForgetCreatedFigure()
        {
            if (createdValues == null)
            {
                return;
            }

            TakeDefaultsFrom(createdValues);
            createdIn.DisplayProperties -= Drawing_DisplayProperties;
            createdIn = null;
            CreatedFigure = null;
            createdValues = null;
            tieTarget = null;
        }

        /// <summary>The typed values of the figure just made, edited in its panel or not, as the tool's defaults from now on</summary>
        protected virtual void TakeDefaultsFrom(ITiedValues created)
        {
        }

        // anything else in the side panel (OK, a figure's properties, undo's null) ends the
        // offer; the figure's own panel shown again keeps it
        void Drawing_DisplayProperties(object sender, Drawing.DisplayPropertiesEventArgs e)
        {
            if (!(e.Object is TiedValuesPanel panel && panel.Values == createdValues))
            {
                ForgetCreatedFigure();
            }
        }

        /// <summary>The figure under the cursor that a click would tie a value of the figure just made to, if any</summary>
        IFigure FindTieTarget(Point unconstrainedCoordinates)
        {
            if (createdValues == null)
            {
                return null;
            }

            var figure = Drawing.Figures.HitTest(unconstrainedCoordinates);
            if (figure == null || figure.DependsOn(createdValues))
            {
                return null;
            }

            return createdValues.TiedValueNames.Any(name => createdValues.Accepts(name, figure)) ? figure : null;
        }

        /// <summary>
        /// A click while the panel of the figure just made is up: on a figure one of its values
        /// can be taken from, the value is tied to it - a vector gives a translation both its
        /// distance and its direction, or only what is typed at that moment ("Type the
        /// direction", then a click on another vector: the distance stays with the first),
        /// anything else goes to the first value that takes it - and the panel shows the
        /// source. True when the click was that.
        /// </summary>
        protected bool TryTieCreatedFigure(Point unconstrainedCoordinates)
        {
            var target = FindTieTarget(unconstrainedCoordinates);
            if (target == null)
            {
                return false;
            }

            var values = createdValues;
            var shown = CreatedFigure;
            var names = values.TiedValueNames.Where(name => values.Accepts(name, target)).ToList();
            if (!(target is Vector))
            {
                names = names.Take(1).ToList();
            }
            else if (names.Any(name => !values.IsTied(name)))
            {
                names = names.Where(name => !values.IsTied(name)).ToList();
            }

            bool tied = false;
            using (Transaction.Create(Drawing.ActionManager, false))
            {
                foreach (var name in names)
                {
                    tied |= values.TieTo(name, target);
                }
            }

            if (tied)
            {
                Drawing.RaiseDisplayProperties(new TiedValuesPanel(shown, values));
                Drawing.RaiseStatusNotification(shown.Title + " now follows " + TiedValues.SourceName(target) + ".");
            }
            else
            {
                Drawing.RaiseStatusNotification(TiedValues.SourceName(target) + " is built on " + shown.Name + ": it can't be taken from.");
            }

            return true;
        }

        #endregion

        #region Intermediate results

        #region TempPoint

        protected IPoint TempPoint { get; set; }

        protected virtual void CreateTempPoint(Point coordinates)
        {
            TempPoint = Factory.CreateFreePoint(Drawing, coordinates);
            (TempPoint as FreePoint).IsHitTestVisible = false;
            (TempPoint as FreePoint).Shape.Opacity = 0.5;
            TempPoint.Name = "TempPoint";
            Drawing.ActionManager.ExecuteImmediatelyWithoutRecording = true;
            Actions.Add(Drawing, TempPoint);
            Drawing.ActionManager.ExecuteImmediatelyWithoutRecording = false;
            AddFoundDependency(TempPoint);
        }

        protected void RemoveTempPointIfNecessary()
        {
            if (TempPoint != null)
            {
                FoundDependencies.Remove(TempPoint);
                Drawing.ActionManager.ExecuteImmediatelyWithoutRecording = true;
                Actions.Remove(TempPoint);
                Drawing.ActionManager.ExecuteImmediatelyWithoutRecording = false;
                TempPoint = null;
            }
        }

        #endregion

        #region IntermediateFigure

        public IFigure IntermediateFigure { get; set; }

        void AddIntermediateFigureIfNecessary()
        {
            IntermediateFigure = CreateIntermediateFigure();
            if (IntermediateFigure != null)
            {
                Drawing.ActionManager.ExecuteImmediatelyWithoutRecording = true;
                Actions.Add(Drawing, IntermediateFigure);
                Drawing.ActionManager.ExecuteImmediatelyWithoutRecording = false;
                //Drawing.RaiseAddingOrRemovingFigures(new Drawing.AddingOrRemovingFiguresEventArgs()
                //{
                //    Figures = new List<IFigure>() {IntermediateFigure}
                //});
            }
        }

        protected virtual IFigure CreateIntermediateFigure()
        {
            return null;
        }

        protected virtual void RemoveIntermediateFigureIfNecessary()
        {
            if (IntermediateFigure != null)
            {
                Drawing.ActionManager.ExecuteImmediatelyWithoutRecording = true;
                Actions.Remove(IntermediateFigure);
                Drawing.ActionManager.ExecuteImmediatelyWithoutRecording = false;
                IntermediateFigure = null;
            }
        }

        #endregion

        #region TempResults

        public readonly List<IFigure> TempResults = new List<IFigure>();

        protected virtual bool CanCreateTempResults()
        {
            return ExpectedDependencies.Count == FoundDependencies.Count;
        }

        protected virtual void CreateTempResults()
        {
            var figures = CreateFigures().ToArray();
            TempResults.AddRange(figures);
            Drawing.ActionManager.ExecuteImmediatelyWithoutRecording = true;
            Actions.AddMany(Drawing, figures);
            Drawing.ActionManager.ExecuteImmediatelyWithoutRecording = false;
            //Drawing.RaiseAddingOrRemovingFigures(new Drawing.AddingOrRemovingFiguresEventArgs()
            //{
            //    Figures = figures.ToList()
            //});
        }

        protected virtual void RemoveTempResultsIfNecessary()
        {
            if (TempResults.Count > 0)
            {
                foreach (var item in TempResults)
                {
                    Drawing.ActionManager.ExecuteImmediatelyWithoutRecording = true;
                    Actions.Remove(item);
                    Drawing.ActionManager.ExecuteImmediatelyWithoutRecording = false;
                }
                TempResults.Clear();
            }
        }

        #endregion

        #endregion

        protected abstract IEnumerable<IFigure> CreateFigures();

        #region Found & next dependencies

        protected bool usePointsUnderMouse = true;
        protected DependencyList ExpectedDependencies { get; set; }
        protected readonly List<IFigure> FoundDependencies = new List<IFigure>();

        /// <summary>
        /// Gets the currently expected type of dependency.
        /// </summary>
        /// <returns>IPoint if TempPoint != null</returns>
        protected virtual Type GetExpectedDependencyType()
        {
            if (TempPoint != null)
            {
                return typeof(IPoint);
            }
            if (FoundDependencies.Count < ExpectedDependencies.Count)
            {
                return ExpectedDependencies[FoundDependencies.Count];
            }
            return null;
        }

        /// <summary>
        /// Is the next expected dependency an IPoint?
        /// </summary>
        /// <returns>IPoint if TempPoint != null</returns>
        protected virtual bool ExpectingAPoint()
        {
            var expected = GetExpectedDependencyType();
            return expected != null && typeof(IPoint).IsAssignableFrom(expected);
        }

        protected void AdvertiseNextDependency()
        {
            var nextDependency = GetExpectedDependencyType();
            this.ConstructionComplete = false;
            Drawing.RaiseConstructionStepComplete(new Drawing.ConstructionStepCompleteEventArgs()
            {
                ConstructionComplete = false,
                FigureTypeNeeded = nextDependency
            });
        }

        protected bool CanReuseDependency { get; set; }

        /// <summary>
        /// A figure the construction has and does not take a second time: any it has found,
        /// unless the tool may come back to one (<see cref="CanReuseDependency"/>: a circle
        /// around an end of its own radius) - and then not where that would make nothing
        /// (<see cref="IsDegenerateRepeat"/>: the same point twice in a row is the second
        /// click of a double click, and By Radius made a circle of radius 0 of it, nowhere
        /// to be seen but in the list and in the file).
        /// </summary>
        protected bool IsTaken(IFigure figure)
        {
            if (figure == null || !FoundDependencies.Contains(figure))
            {
                return false;
            }

            return !CanReuseDependency || IsDegenerateRepeat(figure, FoundDependencies.Where(found => found != TempPoint).ToList());
        }

        /// <summary>For a tool that may take a figure again: whether taking this one, which it has, as the next would make nothing</summary>
        /// <param name="found">What the construction has so far, without the point following the cursor</param>
        protected virtual bool IsDegenerateRepeat(IFigure figure, IList<IFigure> found)
        {
            return false;
        }

        protected abstract DependencyList InitExpectedDependencies();

        protected virtual void AddFoundDependency(IFigure figure)
        {
            if (figure != null && GetExpectedDependencyType().IsAssignableFrom(figure.GetType()))
            {
                FoundDependencies.Add(figure);
            }
        }

        #endregion

        #region State machine transition on clicking

        /// <summary>
        /// Assumes coordinates are logical already
        /// </summary>
        /// <param name="coordinates">Logical coordinates of the click point</param>
        protected virtual void Click(Point coordinates)
        {
            AddDependency(coordinates);
        }

        protected virtual void AddDependency(Point coordinates)
        {
            IFigure underMouse = null;
            EnsureTransaction();

            if (GetExpectedDependencyType() != null)
            {
                // MouseDownUnconstrainedCoordinates used here to properly find the figure under the mouse.
                underMouse = LookForExpectedDependencyUnderCursor(ClickedUnconstrainedCoordinates);

                // Typed coordinates take a point only where it is exactly. Found as a click
                // finds one - anything within reach of a cursor - a vertex typed close to
                // the one before was taken for it and dropped, and one typed near any other
                // point became that point.
                if (!canPlacePointsOnFigures
                    && underMouse is IPoint near
                    && !near.Coordinates.EqualsWithPrecision(ClickedUnconstrainedCoordinates))
                {
                    underMouse = null;
                }

                if (IsTaken(underMouse))
                {
                    return;
                }

                // the preview goes before a new point takes a name: its points hold letters (a
                // square's other two vertices), and the click's point would be G of square DGEF
                RemoveIntermediateFigureIfNecessary();
                RemoveTempResultsIfNecessary();

                if (underMouse == null && ExpectingAPoint())
                {
                    underMouse = CreatePointForClick(coordinates);
                }
                else if (ExpectingAPoint() && !usePointsUnderMouse)
                {
                    var pointUnderMouse = underMouse as IPoint;
                    var freePoint = CreatePointAtCurrentPosition(coordinates);
                    if (pointUnderMouse != null)
                    {
                        freePoint.Coordinates = pointUnderMouse.Coordinates;
                    }
                    underMouse = freePoint;
                }

                // A click that found nothing the tool takes (a figure is wanted, and the
                // click was on empty paper) is no step of a construction. It used to start
                // one, with nothing in it: the Undo button lit up on an empty drawing, the
                // first Ctrl+Z only put the tool back, and Redo, Delete and Paste did
                // nothing until then.
                if (underMouse == null)
                {
                    DropEmptyTransaction();
                    return;
                }
            }

            StartConstruction();
            RemoveIntermediateFigureIfNecessary();
            RemoveTempResultsIfNecessary();

            if (TempPoint != null)
            {
                if (underMouse == null) throw new NullReferenceException("How come underMouse is null at this point?");
                TempPoint.SubstituteWith(underMouse);
                RemoveTempPointIfNecessary();
            }

            if (GetExpectedDependencyType() != null)
            {
                AddFoundDependency(underMouse);
            }

            if (GetExpectedDependencyType() != null)
            {
                if (ExpectingAPoint())
                {
                    CreateTempPoint(coordinates);
                    if (CanCreateTempResults())
                    {
                        CreateTempResults();
                    }
                    else
                    {
                        AddIntermediateFigureIfNecessary();
                    }
                }

                AdvertiseNextDependency();
            }
            else
            {
                AddFiguresAndRestart();
            }

            Drawing.Figures.CheckConsistency();
        }

        /// <summary>
        /// It is important to exclude TempResults from the search since
        /// we don't want the figure to depend on its own parts.
        /// </summary>
        protected virtual IFigure LookForExpectedDependencyUnderCursor(Point coordinates)
        {
            return Drawing.Figures.HitTest(coordinates, f =>
            {
                if (f == null || !f.Visible || !f.IsHitTestVisible)
                {
                    return false;
                }

                var expected = GetExpectedDependencyType();
                if (!expected.IsAssignableFrom(f.GetType()))
                {
                    return false;
                }

                // (a label that says no number is nothing to take a length or an angle
                // from, nor is an angle's mark a length)
                if (expected == typeof(ILengthProvider) && !f.GivesLength()
                    || expected == typeof(IAngleProvider) && !f.GivesAngle())
                {
                    return false;
                }

                if (!TempResults.IsEmpty() && TempResults.Contains(f))
                {
                    return false;
                }

                // Nor a part of what is being drawn (the vertices and sides of a regular
                // polygon, which are not in TempResults themselves): it follows the point
                // under the cursor and is gone with the preview. A double click with the
                // Regular polygon tool took a vertex of the preview, all of them on the
                // center just then, for the polygon's own vertex.
                if (TempPoint != null && f.DependsOn(TempPoint))
                {
                    return false;
                }

                return true;
            });
        }

        #endregion

        #region Point placement

        bool canPlacePointsOnFigures = true;
        PointPlacement hoverPlacement;

        /// <summary>
        /// What a click gives when the tool needs a point: an existing point, or a new one that
        /// is free, on a figure, at an intersection or (with snap to midpoint) in the middle of
        /// a segment. Null if the tool doesn't need a point now.
        /// </summary>
        /// <param name="unconstrainedCoordinates">Where the cursor is</param>
        /// <param name="coordinates">The same after snapping</param>
        protected virtual PointPlacement FindPointPlacement(Point unconstrainedCoordinates, Point coordinates)
        {
            if (!ExpectingAPoint() || FindFigureInsteadOfPoint(unconstrainedCoordinates) != null)
            {
                return null;
            }

            if (!usePointsUnderMouse || !canPlacePointsOnFigures)
            {
                return PointPlacement.Free(coordinates);
            }

            var existing = LookForExpectedDependencyUnderCursor(unconstrainedCoordinates) as IPoint;
            if (existing != null)
            {
                return PointPlacement.Existing(existing);
            }

            return PointPlacement.Find(
                Drawing,
                coordinates,
                Settings.Instance.EnableSnapToCenter,
                CanPlacePointOn);
        }

        /// <summary>
        /// The figure being constructed follows the cursor, so it is always under it:
        /// the new point must not end up depending on it.
        /// </summary>
        protected virtual bool CanPlacePointOn(IFigure figure)
        {
            if (figure == IntermediateFigure || TempResults.Contains(figure))
            {
                return false;
            }

            return TempPoint == null || !figure.DependsOn(TempPoint);
        }

        IFigure CreatePointForClick(Point coordinates)
        {
            var placement = FindPointPlacement(ClickedUnconstrainedCoordinates, coordinates);
            if (placement != null && placement.ExistingPoint != null)
            {
                return placement.ExistingPoint;
            }

            if (placement == null || !placement.IsDependent)
            {
                return CreatePointAtCurrentPosition(coordinates);
            }

            var result = placement.Create(Drawing);
            Actions.Add(Drawing, result);
            return result;
        }

        protected override PointPlacement GetClickPreview(MouseEventArgs e)
        {
            return hoverPlacement;
        }

        IFigure hoverFigure;

        /// <summary>
        /// When the tool needs a figure and not a point: the one a click here would take
        /// </summary>
        protected virtual IFigure FindFigureToPick(Point unconstrainedCoordinates)
        {
            var insteadOfPoint = FindFigureInsteadOfPoint(unconstrainedCoordinates);
            if (insteadOfPoint != null)
            {
                return insteadOfPoint;
            }

            if (GetExpectedDependencyType() == null || ExpectingAPoint())
            {
                return null;
            }

            var figure = LookForExpectedDependencyUnderCursor(unconstrainedCoordinates);
            if (IsTaken(figure))
            {
                return null;
            }

            return figure;
        }

        /// <summary>
        /// A figure that a click takes although the tool is expecting a point: the Distance
        /// tool measures a segment, Circle by Radius takes one as the radius. Null by default.
        /// The hover preview (no ghost point, a halo on the figure, a hand) and the tool's
        /// own click handling must agree, so both go through this.
        /// </summary>
        protected virtual IFigure FindFigureInsteadOfPoint(Point unconstrainedCoordinates)
        {
            return null;
        }

        protected override IFigure GetFigureToPick(MouseEventArgs e)
        {
            if (hoverFigure != null)
            {
                return hoverFigure;
            }

            // an existing point the click would take
            var point = hoverPlacement != null ? hoverPlacement.ExistingPoint : null;
            if (IsTaken(point))
            {
                return null;
            }

            return point;
        }

        #endregion

        #region MouseDown, MouseMove, MouseUp

        protected Point MouseDownCoordinates;
        protected Point ClickedUnconstrainedCoordinates;
        public bool IsMouseButtonDown { get; set; }

        public override void MouseDown(object sender, MouseButtonEventArgs e)
        {
            IsMouseButtonDown = true;
            Point newPosition = Coordinates(e);
            newPosition = AdjustCurrentCoordinates(newPosition);
            MouseDownCoordinates = newPosition;
            ClickedUnconstrainedCoordinates = Coordinates(e, false, false, false);
            if (TryTieCreatedFigure(ClickedUnconstrainedCoordinates))
            {
                return;
            }

            Click(MouseDownCoordinates);
        }

        public override void MouseMove(object sender, MouseEventArgs e)
        {
            Point newPosition = Coordinates(e);
            newPosition = AdjustCurrentCoordinates(newPosition);
            var unconstrainedCoordinates = Coordinates(e, false, false, false);
            tieTarget = FindTieTarget(unconstrainedCoordinates);
            hoverPlacement = tieTarget != null ? null : FindPointPlacement(unconstrainedCoordinates, newPosition);
            hoverFigure = tieTarget ?? FindFigureToPick(unconstrainedCoordinates);

            if (TempPoint != null)
            {
                // the figure being drawn ends where the click would put the point, not at the cursor
                if (hoverPlacement != null)
                {
                    newPosition = hoverPlacement.Coordinates;
                }

                (TempPoint as IMovable).MoveTo(newPosition);
                Drawing.Recalculate();
            }

            Drawing.RaiseConstructionFeedback(new Drawing.ConstructionFeedbackEventArgs()
            {
                FigureTypeNeeded = GetExpectedDependencyType(),
                IsMouseButtonDown = IsMouseButtonDown
            });
        }

        public override void MouseUp(object sender, MouseButtonEventArgs e)
        {
            var coordinates = Coordinates(e);
            coordinates = AdjustCurrentCoordinates(coordinates);
            ClickedUnconstrainedCoordinates = Coordinates(e, false, false, false);
            IsMouseButtonDown = false;

            // In drag-n-drop operations, down and up are considered two different "clicks"
            // This enables creating segments by a simple drag-and-drop operation (down-drag-release)
            if (TempPoint != null && coordinates.Distance(MouseDownCoordinates) > 3 * CursorTolerance)
            {
                Click(coordinates);
            }
        }

        #endregion

        public override bool IsInInitialState
        {
            get
            {
                if (!FoundDependencies.IsEmpty())
                {
                    return false;
                }

                if (Transaction != null && Transaction.HasActions())
                {
                    return false;
                }

                return true;
            }
        }

        #region KeyDown

        public override void KeyDown(object sender, Avalonia.Input.KeyEventArgs e)
        {
            if (e.Key == Avalonia.Input.Key.Escape)
            {
                if (FoundDependencies.IsEmpty())
                {
                    AbortAndSetDefaultTool();
                }
                else
                {
                    Restart();
                }
                e.Handled = true;
            }
        }

        #endregion

        #region Cursor

        /// <summary>
        /// See <see cref="Behavior.GetCursor(Point)"/>. When the tool needs a point the cursor
        /// follows what the click would make of it; when it needs a figure (a line to be
        /// perpendicular to) it is a hand over a suitable one.
        /// </summary>
        protected override Avalonia.Input.Cursor GetCursor(Point coordinates)
        {
            // a click ties a value of the figure just made to what is here
            if (tieTarget != null)
            {
                return HandCursor;
            }

            if (GetExpectedDependencyType() == null)
            {
                return ArrowCursor;
            }

            if (FindFigureInsteadOfPoint(coordinates) != null)
            {
                return HandCursor;
            }

            if (ExpectingAPoint())
            {
                // a point the tool already has: the click is ignored
                if (hoverPlacement != null && IsTaken(hoverPlacement.ExistingPoint))
                {
                    return ArrowCursor;
                }

                return GetCursor(hoverPlacement);
            }

            return hoverFigure != null ? HandCursor : ArrowCursor;
        }

        #endregion

        #region Adjust point coordinates

        protected virtual Point AdjustCurrentCoordinates(Point newPosition)
        {
            List<Point> points = new List<Point>(this.FoundDependencies.ToPoints());
            if (points.Count > 1 && TempPoint is FreePoint)
            {
                newPosition = GetOrthoOrPolar(points[points.Count - 2], newPosition);
            }
            return newPosition;
        }

        /// <summary>
        /// Helper method to avoid code duplication
        /// </summary>
        private Point GetOrthoOrPolar(Point center, Point newPosition)
        {
            if (Settings.Instance.EnableOrtho)
            {
                newPosition = Math.GetOrthoPosition(center, newPosition);
                return newPosition;
            }

            double PolarIncrement = DynamicGeometry.Settings.Instance.PolarIncrement.Val;

            if (Settings.Instance.UserModeAngle && !Settings.Instance.UserModeLength)
            {
                double angle = Settings.Instance.UserAngle;
                newPosition = Math.GetPositionByExactAngle(center, newPosition, angle);
            }
            else if (Settings.Instance.UserModeAngle && Settings.Instance.UserModeLength)
            {
                double angle = Settings.Instance.UserAngle;
                double length = Settings.Instance.UserLength;
                newPosition = Math.GetPositionByExactAngleAndLength(center, newPosition, angle, length);
            }

            return newPosition;
        }

        #endregion
    }
}
