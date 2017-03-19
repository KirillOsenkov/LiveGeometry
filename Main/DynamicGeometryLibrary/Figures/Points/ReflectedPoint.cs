using System.Linq;
using System.Windows.Shapes;

namespace DynamicGeometry
{
    public class ReflectedPoint : PointBase, IPoint
    {
        [PropertyGridVisible]
        [PropertyGridName("Reflection Of ")]
        public IFigure Source
        {
            get
            {
                return Dependencies.ElementAt(0);
            }
        }

        [PropertyGridVisible]
        [PropertyGridName("Reflected About ")]
        public IFigure Mirror
        {
            get
            {
                return Dependencies.ElementAt(1);
            }
        }

        protected override Shape CreateShape()
        {
            return Factory.CreateDependentPointShape();
        }

        protected override void OnDependenciesChanged()
        {
            mirrorPoint = Mirror as IPoint;
            mirrorLine = Mirror as ILine;
            mirrorCircle = Mirror as ICircle;
        }

        private IPoint mirrorPoint;
        private ILine mirrorLine;
        private ICircle mirrorCircle;

        public override void Recalculate()
        {
            var source = Point(0);
            if (mirrorPoint != null)
            {
                Coordinates = Math.GetSymmetricPoint(source, mirrorPoint.Coordinates);
            }
            else if (mirrorLine != null)
            {
                Coordinates = Math.GetSymmetricPoint(source, mirrorLine.Coordinates);
            }
            else if (mirrorCircle != null)
            {
                Coordinates = Math.GetSymmetricPoint(source, mirrorCircle.Center, mirrorCircle.Radius);
            }

            Exists = Coordinates.Exists();
        }
    }
}