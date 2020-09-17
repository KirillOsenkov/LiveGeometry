using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Linq;

namespace DynamicGeometry
{
    public abstract partial class Behavior : INotifyPropertyChanged
    {
        public static Behavior Default { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void RaisePropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        public virtual string Category
        {
            get
            {
                return "Common";
            }
        }

        public abstract string Name
        {
            get;
        }

        public virtual string HintText
        {
            get
            {
                return "";
            }
        }

        public virtual string ConstructionHintText(Drawing.ConstructionStepCompleteEventArgs args)
        {
            string expectedFigure = "";
            if (args.FigureTypeNeeded.HasInterface<IPoint>())
            {
                expectedFigure = "point";
            }
            else if (args.FigureTypeNeeded == typeof(Vector))
            {
                expectedFigure = "vector or enter values";
            }
            else if (args.FigureTypeNeeded == typeof(IAngleProvider))
            {
                expectedFigure = "figure with an angle such as an arc or enter value";
            }
            else if (args.FigureTypeNeeded == typeof(ILengthProvider))
            {
                expectedFigure = "figure with length such as a segment or enter value";
            }
            else if (args.FigureTypeNeeded.HasInterface<ILine>())
            {
                expectedFigure = "line, ray or a segment";
            }
            else if (args.FigureTypeNeeded.HasInterface<ICircle>())
            {
                expectedFigure = "circle";
            }
            else if (args.FigureTypeNeeded.HasInterface<IEllipse>())
            {
                expectedFigure = "circle or ellipse";
            }
            else if (args.FigureTypeNeeded.HasInterface<ILinearFigure>())
            {
                expectedFigure = "line or a circle";
            }
            else
            {
                expectedFigure = args.FigureTypeNeeded.Name;
            }
            string hint = string.Format("Select a {0}.", expectedFigure);
            return hint;
        }

        private UIElement icon;
        public UIElement Icon
        {
            get
            {
                if (icon == null)
                {
                    icon = CreateIcon();
                }
                return icon;
            }
        }

        public abstract FrameworkElement CreateIcon();

        protected Drawing mDrawing;
        public virtual Drawing Drawing
        {
            get
            {
                return mDrawing;
            }
            set
            {
                if (mDrawing != null)
                {
                    mDrawing.OnAttachToCanvas -= mDrawing_OnAttachToCanvas;
                    mDrawing.OnDetachFromCanvas -= mDrawing_OnDetachFromCanvas;
                    ParentCanvas = null;
                }
                mDrawing = value;
                if (mDrawing != null)
                {
                    mDrawing.OnAttachToCanvas += mDrawing_OnAttachToCanvas;
                    mDrawing.OnDetachFromCanvas += mDrawing_OnDetachFromCanvas;
                    ParentCanvas = mDrawing.Canvas;
                }
            }
        }

        void mDrawing_OnAttachToCanvas(Canvas e)
        {
            ParentCanvas = e;
        }

        void mDrawing_OnDetachFromCanvas(Canvas e)
        {
            ParentCanvas = null;
        }

        private Canvas mParentCanvas;
        public virtual Canvas ParentCanvas
        {
            get
            {
                return mParentCanvas;
            }
            set
            {
                if (mParentCanvas != null)
                {
                    mParentCanvas.MouseWheel -= MouseWheel;
                    mParentCanvas.MouseLeftButtonDown -= SafeMouseDown;
                    mParentCanvas.MouseMove -= SafeMouseMove;
                    mParentCanvas.MouseLeftButtonUp -= SafeMouseUp;
                    mParentCanvas.KeyDown -= SafeKeyDown;
                    mParentCanvas.KeyUp -= SafeKeyUp;
                }
                mParentCanvas = value;
                if (mParentCanvas != null)
                {
                    mParentCanvas.MouseWheel += MouseWheel;
                    mParentCanvas.MouseLeftButtonDown += SafeMouseDown;
                    mParentCanvas.MouseMove += SafeMouseMove;
                    mParentCanvas.MouseLeftButtonUp += SafeMouseUp;
                    mParentCanvas.KeyDown += SafeKeyDown;
                    mParentCanvas.KeyUp += SafeKeyUp;
                }
            }
        }

        public virtual object PropertyBag
        {
            get
            {
                return null;
            }
        }

        public virtual void Started()
        {

        }

        public virtual void Stopping()
        {

        }

        public void Restart()
        {
            Stopping();
            Started();
        }

        public virtual bool IsInInitialState 
        {
            get
            {
                return true;
            }
        }

#if !PLAYER

        protected void AbortAndSetDefaultTool()
        {
            Drawing.SetDefaultBehavior();
        }

#endif

        public static bool IsCtrlPressed()
        {
            return (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control;
        }

        public virtual void KeyDown(object sender, KeyEventArgs e)
        {
#if !PLAYER
            if (e.Key == System.Windows.Input.Key.Escape)
            {
                AbortAndSetDefaultTool();
                e.Handled = true;
            }

            else if (e.Key == System.Windows.Input.Key.Delete)
            {
                try
                {
                    Drawing.DeleteSelection();
                }
                catch (Exception)
                {
                }
            }
            else if (e.Key == System.Windows.Input.Key.A)
            {
                Settings.Instance.AutoLabelPoints = !Settings.Instance.AutoLabelPoints;
                string state = (Settings.Instance.AutoLabelPoints) ? "on." : "off.";
                Drawing.RaiseStatusNotification("Toggling automatic labeling of points " + state);
            }
            else if (e.Key == System.Windows.Input.Key.G)
            {
                Settings.Instance.EnableSnapToGrid = !Settings.Instance.EnableSnapToGrid;
                string state = (Settings.Instance.EnableSnapToGrid) ? "on." : "off.";
                Drawing.RaiseStatusNotification("Toggling constraining to grid " + state);
            }

#endif
        }

        public virtual void KeyUp(object sender, KeyEventArgs e)
        {
        }

        public virtual void MouseDown(object sender, MouseButtonEventArgs e)
        {
        }

        public virtual void MouseMove(object sender, MouseEventArgs e)
        {
        }

        public virtual void MouseUp(object sender, MouseButtonEventArgs e)
        {
        }

        void HandleException(Exception ex)
        {
            Drawing.RaiseError(this, ex);
        }

        protected bool errorHappened = false;
        void SafeKeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                KeyDown(sender, e);
            }
            catch (Exception ex)
            {
                HandleException(ex);
            }
        }

        void SafeKeyUp(object sender, KeyEventArgs e)
        {
            try
            {
                KeyUp(sender, e);
            }
            catch (Exception ex)
            {
                HandleException(ex);
            }
        }

        void SafeMouseDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                MouseDown(sender, e);
            }
            catch (Exception ex)
            {
                errorHappened = true;
                HandleException(ex);
            }
        }

        void SafeMouseMove(object sender, MouseEventArgs e)
        {
            if (errorHappened)
            {
                return;
            }
            try
            {
                MouseMove(sender, e);
            }
            catch (Exception ex)
            {
                HandleException(ex);
                errorHappened = true;
            }
        }

        void SafeMouseUp(object sender, MouseButtonEventArgs e)
        {
            errorHappened = false;
            try
            {
                MouseUp(sender, e);
            }
            catch (Exception ex)
            {
                HandleException(ex);
            }
        }

        public virtual void MouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (Drawing != null)
            {
                if (e.Delta > 0)
                {
                    Drawing.CoordinateSystem.ZoomIn();
                }
                else if (e.Delta < 0)
                {
                    Drawing.CoordinateSystem.ZoomOut();
                }
            }
        }

        #region Coordinates

        protected virtual Point Coordinates(MouseEventArgs e)
        {
            return Coordinates(e, Settings.Instance.EnableSnapToPoint, Settings.Instance.EnableSnapToGrid, Settings.Instance.EnableSnapToCenter);
        }

        protected virtual Point Coordinates(MouseEventArgs e, bool snapToPoint, bool snapToGrid, bool snapToCenter)
        {
            var result = e.GetPosition(ParentCanvas);
            result = ToLogical(result);

            if (snapToCenter)
            {
                result = Math.GetSnapToSegmentCenterPosition(result, new List<Segment>(Drawing.Figures.Where(f => f.Visible).ToSegments(result)));
            }
            else
            {
                if (snapToPoint)
                {
                    result = Math.GetSnapToPointPosition(Settings.Instance.SnapGridSpacing, result, new List<Point>(Drawing.Figures.Where(f => f.Visible).ToPoints()), Settings.Instance.EnableSnapToGrid);
                }
                else if (snapToGrid)
                {
                    result = Math.GetSnapToGridPosition(Settings.Instance.SnapGridSpacing, result);
                }
            }
            return result;
        }

        protected double CursorTolerance
        {
            get
            {
                return Drawing.CoordinateSystem.CursorTolerance;
            }
        }

        protected double ToPhysical(double logicalLength)
        {
            return Drawing.CoordinateSystem.ToPhysical(logicalLength);
        }

        protected Point ToPhysical(Point point)
        {
            return Drawing.CoordinateSystem.ToPhysical(point);
        }

        protected double ToLogical(double pixelLength)
        {
            return Drawing.CoordinateSystem.ToLogical(pixelLength);
        }

        protected Point ToLogical(Point pixel)
        {
            return Drawing.CoordinateSystem.ToLogical(pixel);
        }

        protected System.Windows.Rect ToLogical(System.Windows.Rect rect)
        {
            return Drawing.CoordinateSystem.ToLogical(rect);
        }

        protected System.Windows.Rect ToPhysical(System.Windows.Rect rect)
        {
            return Drawing.CoordinateSystem.ToPhysical(rect);
        }
        #endregion
    }
}
