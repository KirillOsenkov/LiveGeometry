using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using M = System.Math;

namespace DynamicGeometry
{
    public partial class CoordinateSystem : IMovable
    {
        public CoordinateSystem(Drawing drawing)
        {
            Check.NotNull(drawing);

            Drawing = drawing;
            Drawing.SizeChanged += Drawing_SizeChanged;
            Origin = PhysicalSize.Scale(0.5).SnapToIntegers().Minus(0.5);
        }

        public Point PhysicalSize
        {
            get
            {
                return new Point(Canvas.ActualWidth, Canvas.ActualHeight);
            }
        }

        public Drawing Drawing { get; set; }

        /// <summary>
        /// Sets the unitLength and origin of the coordinate system.
        /// </summary>
        public void SetViewport(double minX, double maxX, double minY, double maxY)
        {
            var logicalWidth = maxX - minX;
            var logicalHeight = maxY - minY;
            if (!logicalWidth.IsValidPositiveValue())
            {
                throw new ArgumentException("maxX must be greater than minX and both numbers need to exist");
            }
            if (!logicalHeight.IsValidPositiveValue())
            {
                throw new ArgumentException("maxY must be greater than minY and both numbers need to exist");
            }
            var logicalRatio = logicalWidth / logicalHeight;
            var physicalSize = PhysicalSize;
            if (!physicalSize.Exists())
            {
                throw new ArgumentException("Canvas.ActualWidth and Canvas.ActualHeight must be valid values");
            }
            var physicalRatio = physicalSize.X / physicalSize.Y;
            var originScale = new Point(-minX / logicalWidth, 1 + (minY / logicalHeight));
            Point origin;
            double unitLength = 1;

            if (physicalRatio > logicalRatio)
            {
                origin = new Point(physicalSize.Y * logicalRatio * originScale.X, physicalSize.Y * originScale.Y);
                origin.X = origin.X + (physicalSize.X - physicalSize.Y * logicalRatio) / 2;
                unitLength = physicalSize.Y / logicalHeight;
            }
            else
            {
                origin = new Point(physicalSize.X * originScale.X, physicalSize.X * originScale.Y / logicalRatio);
                origin.Y = origin.Y + (physicalSize.Y - physicalSize.X / logicalRatio) / 2;
                unitLength = physicalSize.X / logicalWidth;
            }

            if (unitLength < 1)
            {
                unitLength = 1;
            }
            if (unitLength > 1000)
            {
                unitLength = 1000;
            }
            this.unitLength = unitLength;
            this.scale = unitLength / Settings.DefaultUnitLength;
            this.origin = origin;

            Recalculate();
        }

        const double zoomStep = 4;
        const double zoomFactor = 1.2;
        const double zoomExtendOffset = .5;
        public void ZoomIn()
        {
            UnitLength *= zoomFactor;
        }

        public void ZoomOut()
        {
            UnitLength *= 1 / zoomFactor;
        }

        public void ZoomExtend()
        {
            PointPair bounds = Drawing.Figures.GetLogicalBounds();
            SetViewport(bounds.P1.X - zoomExtendOffset, bounds.P2.X + zoomExtendOffset, bounds.P1.Y - zoomExtendOffset, bounds.P2.Y + zoomExtendOffset);
        }

        //private double mScale = 1.0;
        //public double Scale
        //{
        //    get
        //    {
        //        return mScale;
        //    }
        //    set
        //    {
        //        mScale = value;
        //        UnitLength = value * Settings.DefaultUnitLength;
        //        Drawing.RaiseZoomChanged();
        //    }
        //}

        #region Bounds

        public Canvas Canvas
        {
            get
            {
                return Drawing.Canvas;
            }
        }

        public Point[] LogicalViewportVertices { get; set; }
        public double MinimalVisibleX { get; set; }
        public double MinimalVisibleY { get; set; }
        public double MaximalVisibleX { get; set; }
        public double MaximalVisibleY { get; set; }

        void Drawing_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            Recalculate();
        }

        private void Recalculate()
        {
            LogicalViewportVertices = GetViewportVerticesInLogical();
            if (LogicalViewportVertices != null)
            {
                MinimalVisibleX = LogicalViewportVertices.Min(p => p.X);
                MinimalVisibleY = LogicalViewportVertices.Min(p => p.Y);
                MaximalVisibleX = LogicalViewportVertices.Max(p => p.X);
                MaximalVisibleY = LogicalViewportVertices.Max(p => p.Y);
                Drawing.Recalculate();
            }
        }

        public Point[] GetViewportVerticesInLogical()
        {
            Point[] result = new Point[4];
            var physicalSize = PhysicalSize;
            if (!physicalSize.X.IsValidPositiveValue() || !physicalSize.Y.IsValidPositiveValue())
            {
                return LogicalViewportVertices;
            }
            result[0] = ToLogical(new Point());
            result[1] = ToLogical(new Point(physicalSize.X, 0));
            result[2] = ToLogical(new Point(physicalSize.X, physicalSize.Y));
            result[3] = ToLogical(new Point(0, physicalSize.Y));
            return result;
        }

        public IEnumerable<double> GetVisibleXPoints()
        {
            for (var x = M.Ceiling(MinimalVisibleX); x <= M.Floor(MaximalVisibleX); x++)
            {
                yield return x;
            }
        }

        public IEnumerable<double> GetVisibleYPoints()
        {
            for (var y = M.Ceiling(MinimalVisibleY); y <= M.Floor(MaximalVisibleY); y++)
            {
                yield return y;
            }
        }

        public IEnumerable<Point> GetVisiblePoints()
        {
            IEnumerable<double> XPoints = GetVisibleXPoints();
            IEnumerable<double> YPoints = GetVisibleYPoints();
            for (int i = 0; i < XPoints.Count(); i++)
            {
                yield return new Point(XPoints.ElementAt(i), YPoints.ElementAt(i));
            } 
        }

        #endregion

        #region Coordinate transforms

        private double scale = 1;
        public double Scale
        {
            get
            {
                return scale;
            }
        }

        private double unitLength = Settings.DefaultUnitLength;
        /// <summary>
        /// How many pixels are in a logical unit?
        /// </summary>
        public double UnitLength
        {
            get
            {
                return unitLength;
            }
            set
            {
                if (value < 1 || value > 1000)
                {
                    return;
                }
                unitLength = value;
                scale = unitLength / Settings.DefaultUnitLength;
                Recalculate();
            }
        }

        private Point origin;
        /// <summary>
        /// Origin is in physical coordinates (for 800x600 it will usually be (400;300))
        /// </summary>
        public Point Origin
        {
            get
            {
                return origin;
            }
            private set
            {
                origin = value;
                Recalculate();
            }
        }

        public double CursorTolerance
        {
            get
            {
                return ToLogical(Math.CursorTolerance);
            }
        }

        public virtual Point ToLogical(Point physicalPoint)
        {
            return new Point(
                 (physicalPoint.X - origin.X) / unitLength,
                -(physicalPoint.Y - origin.Y) / unitLength).RoundToEpsilon();
        }

        public IEnumerable<Point> ToLogical(IEnumerable<Point> physicalPoints)
        {
            return physicalPoints.Select(p => ToLogical(p));
        }

        public PointPair ToLogical(PointPair pointPair)
        {
            var result = new PointPair(ToLogical(pointPair.P1), ToLogical(pointPair.P2));
            if (result.P1.X > result.P2.X)
            {
                result = result.Reverse;
            }
            if (result.P1.Y > result.P2.Y)
            {
                var temp = result.P2.Y;
                result.P2.Y = result.P1.Y;
                result.P1.Y = temp;
            }
            return result;
        }

        public virtual double ToLogical(double length)
        {
            return length / UnitLength;
        }

        public virtual System.Windows.Rect ToLogical(System.Windows.Rect rect)
        {
            var result = new System.Windows.Rect();
            var origin = new Point { X = rect.X, Y = rect.Y };
            var logicalOrigin = ToLogical(origin);
            result.X = logicalOrigin.X;
            result.Y = logicalOrigin.Y;
            result.Width = ToLogical(rect.Width);
            result.Height = ToLogical(rect.Height);
            return result;
        }

        public virtual System.Windows.Rect ToPhysical(System.Windows.Rect rect)
        {
            var result = new System.Windows.Rect();
            var origin = new Point{X = rect.X,Y = rect.Y};
            var physicalOrigin = ToPhysical(origin);
            result.X = physicalOrigin.X;
            result.Y = physicalOrigin.Y;
            result.Width = ToPhysical(rect.Width);
            result.Height = ToPhysical(rect.Height);
            return result;
        }

        public virtual Point ToPhysical(Point logicalPoint)
        {
            return new Point(
                origin.X + logicalPoint.X * unitLength,
                origin.Y - logicalPoint.Y * unitLength);
        }

        public PointPair ToPhysical(PointPair logicalPointPair)
        {
            return new PointPair(ToPhysical(logicalPointPair.P1), ToPhysical(logicalPointPair.P2));
        }

        public IEnumerable<Point> ToPhysical(IEnumerable<Point> logicalPoints)
        {
            return logicalPoints.Select(p => ToPhysical(p));
        }

        public void ToPhysicalInPlace(List<Point> logicalPoints)
        {
            for (int i = 0; i < logicalPoints.Count; i++)
            {
                logicalPoints[i] = ToPhysical(logicalPoints[i]);
            }
        }

        public virtual double ToPhysical(double length)
        {
            return length * UnitLength;
        }

        #endregion

        #region IMovable Members

        public void MoveTo(Point position)
        {
            var old = origin.SnapToIntegers();
            position = ToPhysical(position).SnapToIntegers();
            if (old == position)
            {
                return;
            }
            Origin = position;
        }

        public bool AllowMove()
        {
            return true;
        }

        public Point Coordinates
        {
            get { return new Point(); }
        }

        #endregion
    }
}
