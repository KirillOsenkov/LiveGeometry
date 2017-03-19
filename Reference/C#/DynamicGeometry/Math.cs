using System.Windows;
using System;
namespace DynamicGeometry
{
    public static class PointExtensions
    {
        public static double Distance(this Point p1, Point p2)
        {
            return Math.Sqrt(
                  Math.Pow(p1.X - p2.X, 2) 
                + Math.Pow(p1.Y - p2.Y, 2));
        }
    }
}