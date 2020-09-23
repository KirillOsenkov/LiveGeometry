using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Media;

namespace DynamicGeometry
{
    public class FunctionGraph : Curve, ILinearFigure
    {
        public FunctionGraph()
        {
            for (int i = 0; i < StepCount; i++)
            {
                pathSegments.Add(new LineSegment());
            }
        }

        int StepCount
        {
            get
            {
                if (Drawing == null || Drawing.CoordinateSystem == null)
                {
                    return 0;
                }
                return (int)Drawing.CoordinateSystem.PhysicalSize.X / 10;
            }
        }

        private Func<double, double> mFunction;
        public Func<double, double> Function
        {
            get
            {
                return mFunction;
            }
            private set
            {
                if (value != null)
                {
                    mFunction = value;
                    UpdateVisual();
                }
            }
        }

        public override void ReadXml(System.Xml.Linq.XElement element)
        {
            base.ReadXml(element);
            mFunctionText = element.ReadString("Function");
        }

        public override void WriteXml(System.Xml.XmlWriter writer)
        {
            base.WriteXml(writer);
            writer.WriteAttributeString("Function", FunctionText);
        }

        string mFunctionText;
        [PropertyGridVisible]
        [PropertyGridName("f(x) = ")]
        public string FunctionText
        {
            get
            {
                return mFunctionText;
            }
            set
            {
                mFunctionText = value;
                var result = Compile();
                if (result.IsSuccess)
                {
                    Drawing.ClearStatus();
                }
                else
                {
                    Drawing.RaiseStatusNotification(result.ToString());
                };
            }
        }

        public CompileResult Compile()
        {
            var result = Compiler.Instance.CompileFunction(Drawing, FunctionText);
            if (result.IsSuccess)
            {
                SetFunction(result);
            }
            return result;
        }

        public override void Recalculate()
        {
            if (Function == null)
            {
                Compile();
            }
            base.Recalculate();
        }

        void SetFunction(CompileResult result)
        {
            Function = result.Function;

            this.UnregisterFromDependencies();
            Dependencies.SetItems(result.Dependencies);

            // See explanation in DrawingExpression.Recalculate().
            if (Drawing.Figures.Contains(this))
            {
                this.RegisterWithDependencies();
                this.RecalculateAllDependents();
            }
        }

        double CallFunction(double parameter)
        {
            if (Function == null)
            {
                return 0;
            }
            try
            {
                return Function(parameter);
            }
            catch (Exception)
            {
                return 0;
            }
        }

        public override void GetPoints(List<Point> result)
        {
            var stepCount = StepCount;
            if (stepCount == 0)
            {
                return;
            }

            CoordinateSystem coordinates = Drawing.CoordinateSystem;
            double minX = coordinates.MinimalVisibleX;
            double maxX = coordinates.MaximalVisibleX;

            var step = (maxX - minX) / StepCount;
            for (double x = minX; x < maxX; x += step)
            {
                double y = CallFunction(x);
                if (y.IsValidValue())
                {
                    result.Add(new Point(x, y));
                }
            }
            double finalY = CallFunction(maxX);
            if (finalY.IsValidValue())
            {
            result.Add(new Point(maxX, finalY));
            }
        }

        public override double GetNearestParameterFromPoint(Point point)
        {
            return point.X;
        }

        public override Point GetPointFromParameter(double parameter)
        {
            return new Point(parameter, CallFunction(parameter));
        }

        public override Tuple<double, double> GetParameterDomain()
        {
            CoordinateSystem coordinates = Drawing.CoordinateSystem;
            return Tuple.Create(coordinates.MinimalVisibleX, coordinates.MaximalVisibleX);
        }

        public override IFigure HitTest(Point point)
        {
            return base.HitTest(point);

            // The solution below fails to detect hits on high slope sections of functions. For example f(x) = x^3 or f(x) = 100 * x. - D.H.
            //return this.IsPointWithinTolerance(point) ? this : null;
        }
    }
}
