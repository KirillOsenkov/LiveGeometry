using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Shapes;
using GuiLabs.Undo;

namespace DynamicGeometry
{
    public abstract partial class FigureCreator : Behavior
    {
        #region Dialog

        [PropertyGridName("Point by coordinates")]
        public class Dialog
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
            public void AddPoint()
            {
                var xresult = parent.Drawing.CompileExpression(X);
                var yresult = parent.Drawing.CompileExpression(Y);

                if (xresult.IsSuccess && yresult.IsSuccess)
                {
                    this.parent.AddDependency(new Point(double.Parse(X, CultureInfo.InvariantCulture), double.Parse(Y, CultureInfo.InvariantCulture)));
                }
            }
        }

        public override object PropertyBag
        {
            get
            {
                if (ExpectingAPoint())
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
        /// </summary>
        protected Transaction Transaction { get; set; }

        protected bool ConstructionComplete { get; set; }

        public override void Started()
        {
            this.ConstructionComplete = true;
            Transaction = Transaction.Create(Drawing.ActionManager, false);
            ExpectedDependencies = InitExpectedDependencies();
            FoundDependencies.Clear();
        }

        public override void Stopping()
        {
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
            RemoveTempResultsIfNecessary();
            var figures = CreateFigures().ToList();
            foreach (var figure in figures)
            {
                if (figure != null)
                {
                    Actions.Add(Drawing, figure);
                }
            }
            Drawing.RaiseUserIsAddingFigures(new Drawing.UIAFEventArgs() { Figures = figures });
            Transaction.Commit();
            Transaction = null;
            this.ConstructionComplete = true;
            Drawing.RaiseConstructionStepComplete(new Drawing.ConstructionStepCompleteEventArgs()
            {
                ConstructionComplete = true
            });
            Restart();
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

            if (GetExpectedDependencyType() != null)
            {
                // MouseDownUnconstrainedCoordinates used here to properly find the figure under the mouse.
                underMouse = LookForExpectedDependencyUnderCursor(ClickedUnconstrainedCoordinates);

                if (underMouse != null && FoundDependencies.Contains(underMouse) && !CanReuseDependency)
                {
                    return;
                }

                if (underMouse == null && ExpectingAPoint())
                {
                    underMouse = CreatePointAtCurrentPosition(coordinates);
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
            }

            Drawing.RaiseConstructionStepStarted();
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

                if (!GetExpectedDependencyType().IsAssignableFrom(f.GetType()))
                {
                    return false;
                }

                if (!TempResults.IsEmpty() && TempResults.Contains(f))
                {
                    return false;
                }

                return true;
            });
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
            Click(MouseDownCoordinates);
        }

        public override void MouseMove(object sender, MouseEventArgs e)
        {
            if (TempPoint != null)
            {
                Point newPosition = Coordinates(e);
                newPosition = AdjustCurrentCoordinates(newPosition);
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

        public override void KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Escape)
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
