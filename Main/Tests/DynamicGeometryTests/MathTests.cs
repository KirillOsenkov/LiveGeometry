using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace DynamicGeometry
{
    [TestClass]
    public class MathTests
    {
        [TestMethod]
        public void IsPointInPolygon_HitTestCollinearWithSide()
        {
            TestIsPointInPolygon(0, 0, false, 1, 0, 2, 0, 2, 1, 1, 1);
        }

        private void TestIsPointInPolygon(
            double x, 
            double y, 
            bool expectedInside, 
            params double[] coordinates)
        {
            var point = new Point(x, y);
            var vertices = Points(coordinates);
            bool actualInside = vertices.IsPointInPolygon(point);
            Assert.AreEqual(expectedInside, actualInside);
        }

        private List<Point> Points(params double[] coordinates)
        {
            List<Point> result = new List<Point>();
            for (int i = 0; i < coordinates.Length / 2; i++)
            {
                result.Add(new Point(coordinates[i * 2], coordinates[i * 2 + 1]));
            }

            return result;
        }
    }
}
