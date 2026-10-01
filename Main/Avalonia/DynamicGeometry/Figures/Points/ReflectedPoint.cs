using System.Linq;
using Avalonia.Controls.Shapes;

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

            // The image of a point that is not there is not there either. (Asked only
            // whether its own coordinates were numbers - and the source keeps its last ones -
            // the image of an intersection that had gone stayed on screen, frozen, with all
            // that was built on it. Likewise a rotated, a dilated point, a point by
            // coordinates and an angle bisector.)
            Exists = Dependencies.Exists() && Coordinates.Exists();
        }
    }
}