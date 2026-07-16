using System.Linq;

namespace DynamicGeometry
{
    public class AngleBisector : Ray
    {
        PointPair coordinates;

        public override PointPair Coordinates
        {
            get
            {
                return coordinates;
            }
        }

        [PropertyGridVisible]
        public override double Angle
        {
            get
            {
                var dependencies = GetDependencies();
                double result = 0;
                if (dependencies != null)
                {
                   
                    result = (Flipped) ? 
                        Math.OAngle(
                        dependencies.Point(2),
                        dependencies.Point(0),
                        dependencies.Point(1)).ToDegrees() :
                        Math.OAngle(
                        dependencies.Point(1),
                        dependencies.Point(0),
                        dependencies.Point(2)).ToDegrees();
                }
                return result;
            }
        }

        [PropertyGridVisible]
        [PropertyGridName("Convert to opposite angle")]
        public void ConvertToOpposite()
        {
            IFigure[] dependencies = Dependencies as IFigure[];
            if (dependencies != null)
            {
                var t = dependencies[1];
                dependencies[1] = dependencies[2];
                dependencies[2] = t;
            }
            this.RecalculateAllDependents();
            Drawing.RaiseSelectionChanged(this);
        }

        IFigure[] GetDependencies()
        {
            var dependencies = Dependencies.ToArray();
            if (dependencies.Length == 1)
            {
                AngleMeasurement angle = dependencies[0] as AngleMeasurement;
                if (angle == null)
                {
                    return null;
                }
                dependencies = angle.Dependencies.ToArray();
            }
            if (dependencies.Length != 3)
            {
                dependencies = null;
            }
            return dependencies;
        }

        public override void Recalculate()
        {
            var dependencies = GetDependencies();
            if (dependencies != null)
            {
                coordinates.P1 = dependencies.Point(0);
                coordinates.P2 = (Flipped) ?
                    Math.GetAngleBisectorPoint(dependencies.Point(0), dependencies.Point(2), dependencies.Point(1)) :
                    Math.GetAngleBisectorPoint(dependencies.Point(0), dependencies.Point(1), dependencies.Point(2));
                Exists = coordinates.P2.Exists();
            }
            else
            {
                Exists = false;
            }
        }
    }
}