using System.Collections.Generic;

namespace DynamicGeometry
{
    public class Transformer
    {
        public static bool CanBeTransformSource(IFigure figure)
        {
            // Not yet supported
            if (figure is CircleByEquation || figure is LineByEquation || figure is FunctionGraph || figure is Locus)
            {
                return false;
            }

            // Supported
            if (figure is IPoint || figure is ILine || figure is IEllipse || figure is IPolygonalChain)
            {
                return true;
            }
            return false;
        }

        public static bool CanFigureBeMirrorForSource(IFigure figure, IFigure source)
        {
            if (figure is IPoint || figure is ILine)
            {
                return true;
            }
            else if (figure is ICircle)
            {
                return source is IPoint;
            }
            return false;
        }

        public static List<IFigure> CreateReflectedFigure(Drawing drawing, IFigure source, IFigure mirror)
        {
            Check.NotNull(source, "source");
            Check.NotNull(mirror, "mirror");

            List<IFigure> result = new List<IFigure>();
            if (source is IPoint)
            {
                var reflectedPoint = Factory.CreateReflectedPoint(drawing, new [] { source, mirror });
                if (reflectedPoint == null)
                {
                    throw "reflectedPoint is null. source = {0}, mirror = {1}"
                        .Format(source, mirror)
                        .AsException();
                }
                reflectedPoint.Visible = source.Visible;
                reflectedPoint.Name = source.Name + "'";
                result.Add(reflectedPoint);
            }
            else if ((source is ILine || source is IEllipse || source is IPolygonalChain) && !(mirror is ICircle))
            {
                var dependencies = new List<IFigure>();
                foreach (var dependency in source.Dependencies)
                {
                    var reflectedDependency = CreateReflectedFigure(drawing, dependency, mirror);
                    if (reflectedDependency == null)
                    {
                        throw "reflectedDependency is null. dependency = {0}, mirror = {1}"
                            .Format(dependency, mirror)
                            .AsException();
                    }
                    if (reflectedDependency.IsEmpty())
                    {
                        throw "reflectedDependency is empty. dependency = {0}, mirror = {1}"
                            .Format(dependency, mirror)
                            .AsException();
                    }
                    result.AddRange(reflectedDependency);
                    var last = reflectedDependency.Last();
                    if (last == null)
                    {
                        throw "last = null".AsException();
                    }
                    dependencies.Add(last);
                }
                var reflected = source.Clone();
                if (reflected == null)
                {
                    throw "reflected = null".AsException();
                }
                reflected.UnregisterFromDependencies();
                reflected.Dependencies.SetItems(dependencies);
                result.Add(reflected);

                // Flip the wind of arcs.
                var arc = reflected as IArc;
                if (arc != null && mirror is ILine)
                {
                    arc.Clockwise = !arc.Clockwise;
                }

            }
            return result;
        }

        public static List<IFigure> CreateDilatedFigure(Drawing drawing, IFigure source, IFigure center, IFigure lengthProvider1, IFigure lengthProvider2, double factor)
        {
            Check.NotNull(source, "source");
            Check.NotNull(center, "center");

            var list = new List<IFigure>() { source, center };

            var lp1 = lengthProvider1 as ILengthProvider;
            var lp2 = lengthProvider2 as ILengthProvider;
            if (lp1 != null)
            {
                if (lp2 != null)
                {
                    factor = lp1.Length / lp2.Length;
                    list.Add(lp1, lp2);
                }
                else
                {
                    factor = lp1.Length;
                    list.Add(lp1);
                }
            }

            List<IFigure> result = new List<IFigure>();
            if (source is IPoint)
            {
                var dilatedPoint = Factory.CreateDilatedPoint(drawing, list, factor);
                if (dilatedPoint == null)
                {
                    throw "dilatedPoint is null. source = {0}, center = {1}, segment1 = {2}, segment2 = {3}, factor = {4}"
                        .Format(source, center, lengthProvider1, lengthProvider2, factor)
                        .AsException();
                }
                dilatedPoint.Visible = source.Visible;
                result.Add(dilatedPoint);
            }
            else if (source is ILine || source is IEllipse || source is IPolygonalChain)
            {
                var dependencies = new List<IFigure>();
                foreach (var dependency in source.Dependencies)
                {
                    var dilatedDependency = CreateDilatedFigure(drawing, dependency, center, lengthProvider1, lengthProvider2, factor);
                    if (dilatedDependency == null)
                    {
                        throw "dilatedDependency is null. dependency = {0}, center = {1}, segment1 = {2}, segment2 = {3} factor = {2}"
                            .Format(dependency, center, lengthProvider1, lengthProvider2, factor)
                            .AsException();
                    }
                    if (dilatedDependency.IsEmpty())
                    {
                        throw "dilatedDependency is empty. dependency = {0}, center = {1}, segment1 = {2}, segment2 = {3} factor = {2}"
                            .Format(dependency, center, lengthProvider1, lengthProvider2, factor)
                            .AsException();
                    }
                    result.AddRange(dilatedDependency);
                    var last = dilatedDependency.Last();
                    if (last == null)
                    {
                        throw "last = null".AsException();
                    }
                    dependencies.Add(last);
                }
                var dilated = source.Clone();
                if (dilated == null)
                {
                    throw "dilated = null".AsException();
                }
                dilated.UnregisterFromDependencies();
                dilated.Dependencies.SetItems(dependencies);
                result.Add(dilated);
            }
            return result;
        }

        public static List<IFigure> CreateRotatedFigure(Drawing drawing, IFigure source, IFigure center, IFigure angleProvider, double angle)
        {
            Check.NotNull(source, "source");
            Check.NotNull(center, "center");

            var list = new List<IFigure>() { source, center };

            var safeAngleProvider = angleProvider as IAngleProvider;
            if (safeAngleProvider != null)
            {
                angle = safeAngleProvider.Angle;
                list.Add(angleProvider);
            }

            List<IFigure> result = new List<IFigure>();
            if (source is IPoint)
            {
                var rotatedPoint = Factory.CreateRotatedPoint(drawing, list, angle);
                if (rotatedPoint == null)
                {
                    throw "rotatedPoint is null. source = {0}, center = {1}, angleArc = {2}, angle = {3}"
                        .Format(source, center, angleProvider, angle)
                        .AsException();
                }
                rotatedPoint.Visible = source.Visible;
                result.Add(rotatedPoint);
            }
            else if (source is ILine || source is IEllipse || source is IPolygonalChain)
            {
                var dependencies = new List<IFigure>();
                foreach (var dependency in source.Dependencies)
                {
                    var rotatedDependency = CreateRotatedFigure(drawing, dependency, center, angleProvider, angle);
                    if (rotatedDependency == null)
                    {
                        throw "rotatedDependency is null. dependency = {0}, center = {1}, angleArc = {2}, angle = {3}"
                            .Format(dependency, center, angleProvider, angle)
                            .AsException();
                    }
                    if (rotatedDependency.IsEmpty())
                    {
                        throw "rotatedDependency is empty. dependency = {0}, center = {1}, angleArc = {2} angle = {3}"
                            .Format(dependency, center, angleProvider, angle)
                            .AsException();
                    }
                    result.AddRange(rotatedDependency);
                    var last = rotatedDependency.Last();
                    if (last == null)
                    {
                        throw "last = null".AsException();
                    }
                    dependencies.Add(last);
                }
                var rotated = source.Clone();
                if (rotated == null)
                {
                    throw "rotated = null".AsException();
                }
                rotated.UnregisterFromDependencies();
                rotated.Dependencies = dependencies;
                result.Add(rotated);
            }
            return result;
        }

        public static List<IFigure> CreateTranslatedFigure(Drawing drawing, IFigure source, List<IFigure> dependenciesSubset, double magnitude, double direction)
        {
            Check.NotNull(source, "source");
            List<IFigure> result = new List<IFigure>();
            if (source is IPoint)
            {
                var list = new List<IFigure>() { source };
                list.AddRange(dependenciesSubset);
                var translatedPoint = Factory.CreateTranslatedPoint(drawing, list, magnitude, direction);
                if (translatedPoint == null)
                {
                    throw "translatedPoint is null. source = {0}, magnitude = {1}, direction = {2}"
                        .Format(source, magnitude, direction)
                        .AsException();
                }
                translatedPoint.Visible = source.Visible;
                result.Add(translatedPoint);
            }
            else if (source is ILine || source is IEllipse || source is IPolygonalChain)
            {
                var dependencies = new List<IFigure>();
                foreach (var dependency in source.Dependencies)
                {
                    var translatedDependency = CreateTranslatedFigure(drawing, dependency, dependenciesSubset, magnitude, direction);
                    if (translatedDependency == null)
                    {
                        throw "translatedDependency is null. dependency = {0}, center = {1}, direction = {2}"
                            .Format(dependency, magnitude, direction)
                            .AsException();
                    }
                    if (translatedDependency.IsEmpty())
                    {
                        throw "translatedDependency is empty. dependency = {0}, center = {1}, direction = {2}"
                            .Format(dependency, magnitude, direction)
                            .AsException();
                    }
                    result.AddRange(translatedDependency);
                    var last = translatedDependency.Last();
                    if (last == null)
                    {
                        throw "last = null".AsException();
                    }
                    dependencies.Add(last);
                }
                var translated = source.Clone();
                if (translated == null)
                {
                    throw "translated = null".AsException();
                }
                translated.UnregisterFromDependencies();
                translated.Dependencies = dependencies;
                result.Add(translated);
            }
            return result;
        }
    }
}