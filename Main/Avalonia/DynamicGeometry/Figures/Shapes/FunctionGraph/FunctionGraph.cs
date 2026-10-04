using System;
using System.Collections.Generic;
using Avalonia;

namespace DynamicGeometry
{
    public class FunctionGraph : Curve, ILinearFigure, IRenamableExpressions
    {
        /// <summary>
        /// A sample every couple of pixels: at one per 10 px a sine of frequency 3 was a zigzag
        /// </summary>
        int StepCount
        {
            get
            {
                if (Drawing == null || Drawing.CoordinateSystem == null)
                {
                    return 0;
                }
                return (int)Drawing.CoordinateSystem.PhysicalSize.X / 2;
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

        protected override string Kind
        {
            get
            {
                return "Function";
            }
        }

        /// <summary>Function f, g...</summary>
        protected override string FirstLetter
        {
            get
            {
                return "f";
            }
        }

        /// <summary>"y = sin(x)"</summary>
        public override string Construction
        {
            get
            {
                return string.IsNullOrEmpty(FunctionText) ? null : "y = " + FunctionText;
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
        [PropertyGridPreferredEditor("Function")]
        public string FunctionText
        {
            get
            {
                return mFunctionText;
            }
            set
            {
                // the grid's editor only sets a function that compiles, and shows what is
                // wrong with one that doesn't (FunctionEditor)
                mFunctionText = value;
                Compile();
                RaiseConstructionChanged();
            }
        }

        /// <summary>The text follows renamed figures; the compiled function holds the figures already</summary>
        public void RenameInExpressions(ExpressionRenamer renamer)
        {
            var renamed = renamer.Rewrite(mFunctionText, isFunction: true);
            if (renamed != mFunctionText)
            {
                mFunctionText = renamed;
                RaisePropertyChanged(nameof(FunctionText));
            }
        }

        public void RebindExpressions()
        {
            if (!mFunctionText.IsEmpty())
            {
                Compile();
            }
        }

        public IReadOnlyList<string> ExpressionTexts
        {
            get
            {
                return new[] { mFunctionText };
            }
            set
            {
                if (value[0] != mFunctionText)
                {
                    mFunctionText = value[0];
                    RaisePropertyChanged(nameof(FunctionText));
                }
            }
        }

        public CompileResult Compile()
        {
            var result = Compiler.Instance.CompileFunction(Drawing, FunctionText, figure => !figure.DependsOn(this));
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

        // no value where there is no function, and where it throws (it was 0: a wrong value
        // drawn as if it were one)
        double CallFunction(double parameter)
        {
            if (Function == null)
            {
                return double.NaN;
            }

            try
            {
                return Function(parameter);
            }
            catch (Exception)
            {
                return double.NaN;
            }
        }

        public override void GetPoints(List<Point> result)
        {
            var stepCount = StepCount;
            if (stepCount == 0 || Function == null)
            {
                return;
            }

            CoordinateSystem coordinates = Drawing.CoordinateSystem;
            double minX = coordinates.MinimalVisibleX;
            double maxX = coordinates.MaximalVisibleX;

            // Far beyond the window one value is as good as another, and a coordinate of
            // 1e300 pixels is not drawn at all: exp(x^2) vanished whole. The piece that
            // leaves the window still leaves it steeply.
            double height = coordinates.MaximalVisibleY - coordinates.MinimalVisibleY;
            double lowest = coordinates.MinimalVisibleY - 10 * height;
            double highest = coordinates.MaximalVisibleY + 10 * height;

            Point Clamped(double x, double y)
            {
                return new Point(x, System.Math.Max(lowest, System.Math.Min(highest, y)));
            }

            double previousX = 0;
            double previousY = double.NaN;
            for (int i = 0; i <= stepCount; i++)
            {
                double x = i == stepCount ? maxX : minX + (maxX - minX) * i / stepCount;
                double y = CallFunction(x);
                if (!y.IsValidValue())
                {
                    // the graph goes on to where the function stops having a value, not
                    // only to the last sample before (sqrt(x) starts at 0, not in the air)
                    if (i > 0 && previousY.IsValidValue())
                    {
                        double edge = FindEdge(inside: previousX, outside: x);
                        result.Add(Clamped(edge, CallFunction(edge)));
                    }

                    AddGap(result);
                }
                else
                {
                    if (i > 0 && !previousY.IsValidValue())
                    {
                        double edge = FindEdge(inside: x, outside: previousX);
                        result.Add(Clamped(edge, CallFunction(edge)));
                    }
                    else if (previousY.IsValidValue()
                        && MayBeJump(previousY, y, coordinates)
                        && IsJump(previousX, previousY, x, y))
                    {
                        AddGap(result);
                    }

                    result.Add(Clamped(x, y));
                }

                previousX = x;
                previousY = y;
            }
        }

        /// <summary>Between a place where the function has a value and one where it has none: the last place where it has</summary>
        double FindEdge(double inside, double outside)
        {
            for (int i = 0; i < 20; i++)
            {
                double middle = (inside + outside) / 2;
                if (CallFunction(middle).IsValidValue())
                {
                    inside = middle;
                }
                else
                {
                    outside = middle;
                }
            }

            return inside;
        }

        static void AddGap(List<Point> points)
        {
            if (points.Count > 0 && points[points.Count - 1].Exists())
            {
                points.Add(Gap);
            }
        }

        /// <summary>Units up for one across: a step of the graph steeper than this is looked into</summary>
        const double SteepSlope = 8;

        /// <summary>In pixels: a step of the graph that climbs less than this is no jump to look for</summary>
        const double JumpPixels = 24;

        /// <summary>
        /// Whether the line between two samples next to each other is worth looking into
        /// (<see cref="IsJump"/> costs a dozen more values of the function): it is long on
        /// the screen, and not wholly above or below the window, where it isn't seen
        /// </summary>
        static bool MayBeJump(double y1, double y2, CoordinateSystem coordinates)
        {
            if (System.Math.Abs(y2 - y1) * coordinates.UnitLength < JumpPixels)
            {
                return false;
            }

            bool bothAbove = y1 > coordinates.MaximalVisibleY && y2 > coordinates.MaximalVisibleY;
            bool bothBelow = y1 < coordinates.MinimalVisibleY && y2 < coordinates.MinimalVisibleY;
            return !bothAbove && !bothBelow;
        }

        /// <summary>
        /// Whether the function jumps between two samples next to each other, rather than
        /// climbs: 1/x across 0, tan x across its poles, floor(x) at every whole number.
        /// The two are then not joined - the line between them was drawn as if it were part
        /// of the graph, a vertical one at every pole. Told apart by halving the step
        /// towards the larger change: a climb gets smaller with the step, a jump stays.
        /// </summary>
        bool IsJump(double x1, double y1, double x2, double y2)
        {
            double change = System.Math.Abs(y2 - y1);
            if (change <= SteepSlope * (x2 - x1))
            {
                return false;
            }

            for (int i = 0; i < 12; i++)
            {
                double middle = (x1 + x2) / 2;
                double y = CallFunction(middle);
                if (!y.IsValidValue())
                {
                    return true;
                }

                if (System.Math.Abs(y - y1) > System.Math.Abs(y2 - y))
                {
                    x2 = middle;
                    y2 = y;
                }
                else
                {
                    x1 = middle;
                    y1 = y;
                }
            }

            return System.Math.Abs(y2 - y1) > change / 8;
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
