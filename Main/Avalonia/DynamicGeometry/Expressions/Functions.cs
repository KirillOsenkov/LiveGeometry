using Avalonia;

namespace DynamicGeometry
{
    public static class Functions
    {
        public static double Distance(Point a, Point b)
        {
            return a.Distance(b);
        }

        public static double Dist(Point a, Point b)
        {
            return a.Distance(b);
        }

        public static double Ang(Point a, Point b, Point c)
        {
            return Angle(a, b, c);
        }

        // (not number.SquareRoot(), which takes the root of the absolute value: sqr(-4) was 2)
        public static double Sqr(double number)
        {
            return System.Math.Sqrt(number);
        }

        // Functions of numbers that System.Math has under the same name, as they should
        // be in a drawing (Binder.ResolveMethod asks here first).

        /// <summary>A half goes up, as at school: System.Math.Round takes 2.5 to 2</summary>
        public static double Round(double number)
        {
            return System.Math.Round(number, System.MidpointRounding.AwayFromZero);
        }

        /// <summary>System.Math.Sign throws for what is not a number (the sign of log(x) left of 0)</summary>
        public static double Sign(double number)
        {
            return double.IsNaN(number) ? double.NaN : System.Math.Sign(number);
        }

        /// <summary>System.Math.Clamp throws when the bounds are the wrong way round</summary>
        public static double Clamp(double number, double low, double high)
        {
            return low > high ? double.NaN : System.Math.Max(low, System.Math.Min(high, number));
        }

        public static double Ln(double number)
        {
            return System.Math.Log(number);
        }

        public static double Angle(Point a, Point b, Point c)
        {
            var a1 = Math.GetAngle(b, a);
            var a2 = Math.GetAngle(b, c);
            double result;

            if (a2 < a1)
            {
                result = a1 - a2;
            }
            else
            {
                result = a2 - a1;
            }
            
            if (result >= Math.PI)
            {
                result = 2 * Math.PI - result;
            }
            
            return result;
        }

        public static double OAngle(Point a, Point b, Point c)
        {
            return Math.OAngle(a, b, c);
        }

        public static double XAngle(Point a, Point b)
        {
            return a.AngleTo(b);
        }

        public static double XAng(Point a, Point b)
        {
            return a.AngleTo(b);
        }

        public static double Norm(Point a)
        {
            return a.Length();
        }

        public static double Arg(Point a)
        {
            return a.Arg();
        }

        public static double Area(params IPoint[] points)
        {
            return points.ToPoints().Area();
        }

        public static double Deg(double radians)
        {
            return radians.ToDegrees();
        }

        public static double Rad(double degrees)
        {
            return degrees.ToRadians();
        }

        // the rest of what the VB6 evaluator (modEvaluator.bas) had, for its .dgf files;
        // Sin, Cos, Abs, Round, Sqrt... come from System.Math by name

        public static double OAng(Point a, Point b, Point c)
        {
            return Math.OAngle(a, b, c);
        }

        public static double Int(double number)
        {
            return System.Math.Floor(number);
        }

        public static double Sgn(double number)
        {
            return Sign(number);
        }

        public static double Lg(double number)
        {
            return System.Math.Log10(number);
        }

        public static double ToDeg(double radians)
        {
            return radians.ToDegrees();
        }

        public static double ToRad(double degrees)
        {
            return degrees.ToRadians();
        }
    }
}
