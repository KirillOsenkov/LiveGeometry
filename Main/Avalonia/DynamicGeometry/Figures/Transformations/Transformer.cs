using System;
using System.Collections.Generic;
using System.Linq;

namespace DynamicGeometry
{
    public class Transformer
    {
        /// <summary>
        /// What a click on the figure transforms: the figure, or the regular polygon whose
        /// side or inside it is. A part is not a figure of the drawing (it has no name), and
        /// its copy would be one: an unnamed segment, which a file can't refer to. A vertex
        /// stays a point like any other. Null for what can't be transformed.
        /// </summary>
        /// <param name="keepsLengths">False for a dilation, see <see cref="CanBeTransformSource"/></param>
        public static IFigure FindTransformSource(IFigure figure, bool keepsLengths = true)
        {
            if (figure != null && !(figure is IPoint))
            {
                var owner = figure.Drawing?.Figures
                    .OfType<IFigureParts>()
                    .FirstOrDefault(candidate => candidate.GetPartName(figure) != null);
                if (owner != null)
                {
                    figure = owner;
                }
            }

            return CanBeSource(figure, keepsLengths) ? figure : null;
        }

        /// <summary>
        /// What a transformation takes: a figure it transforms through its points
        /// (<see cref="CanBeTransformSource"/>), or one it traces (<see cref="CanBeTraced"/>)
        /// </summary>
        /// <param name="keepsLengths">False for a dilation, see <see cref="CanBeTransformSource"/></param>
        public static bool CanBeSource(IFigure figure, bool keepsLengths = true)
        {
            return CanBeTransformSource(figure, keepsLengths) || CanBeTraced(figure);
        }

        /// <summary>
        /// Whether the image of the figure can be drawn as the path of the image of a point
        /// that slides along it (<see cref="CreateTracedImage"/>): any figure a point can be
        /// on. That is how a figure that is not transformed through its points (a locus, a
        /// graph, a line or circle given by an equation, a circle whose radius is a number
        /// under a dilation) gets an image, and how anything gets one in a circle: the image
        /// of a line there is a circle, not a line through the images of its points.
        /// </summary>
        public static bool CanBeTraced(IFigure figure)
        {
            return !(figure is IPoint) && PointOnFigure.CanBeOnFigure(figure);
        }

        /// <summary>
        /// The image as a <see cref="Locus"/>: a hidden point slides along the source, and
        /// the locus is the path of its image. Both points are auxiliary, so they go with
        /// the locus. The locus takes the source's line style (a graph's, a locus's).
        /// </summary>
        /// <param name="transformPoint">Makes the image of the sliding point, the image last</param>
        static List<IFigure> CreateTracedImage(Drawing drawing, IFigure source, Func<IPoint, List<IFigure>> transformPoint)
        {
            var domain = ((ILinearFigure)source).GetParameterDomain();
            var sliding = Factory.CreatePointOnFigure(drawing, source, parameter: (domain.Item1 + domain.Item2) / 2);
            sliding.Visible = false;
            sliding.Auxiliary = true;
            var result = new List<IFigure>() { sliding };

            var image = transformPoint(sliding);
            if (image.IsEmpty() || !(image.Last() is IPoint))
            {
                throw "the image of the sliding point is missing. source = {0}"
                    .Format(source)
                    .AsException();
            }

            foreach (var figure in image)
            {
                figure.Visible = false;
                figure.Auxiliary = true;
            }

            result.AddRange(image);
            var locus = Factory.CreateLocus(drawing, new IFigure[] { image.Last(), sliding });
            locus.Visible = source.Visible;
            if (source.Style is LineStyle)
            {
                locus.Style = source.Style;
            }

            result.Add(locus);
            return result;
        }

        /// <summary>
        /// A figure is transformed by transforming what it is built on, down to its points,
        /// so everything it is built on must be something to transform - or a length, which
        /// a reflection, a rotation and a translation leave as it is: a circle by radius
        /// whose radius is a slider, a Number or a measurement keeps that radius. (Asked
        /// only about the figure itself, such a circle was taken, and the tool threw
        /// "dependency is empty" when the last click was made - By Radius makes a slider
        /// for the radius whenever its first click is on empty paper.)
        /// </summary>
        /// <param name="keepsLengths">False for a dilation: a radius that is a number can't be scaled</param>
        public static bool CanBeTransformSource(IFigure figure, bool keepsLengths = true)
        {
            // Not through their points (a line at an angle would transform its angle as if it
            // were a point): these are traced (CanBeTraced)
            if (figure is CircleByEquation || figure is LineByEquation || figure is LineAtAngle || figure is FunctionGraph || figure is Locus)
            {
                return false;
            }

            if (figure is IPoint)
            {
                return true;
            }

            if (!(figure is ILine || figure is IEllipse || figure is IPolygonalChain))
            {
                return false;
            }

            foreach (var dependency in figure.Dependencies)
            {
                if (IsLength(figure, dependency))
                {
                    if (!keepsLengths)
                    {
                        return false;
                    }
                }
                else if (!CanBeTransformSource(dependency, keepsLengths))
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>A radius given by a number, not by a segment or two points: the image has the same</summary>
        static bool IsLength(IFigure figure, IFigure dependency)
        {
            return figure is CircleByRadius
                && dependency is ILengthProvider
                && !(dependency is IPoint || dependency is ILine || dependency is IEllipse || dependency is IPolygonalChain);
        }

        /// <summary>
        /// The segments drawn along the sides of a polygon or polyline (the Triangle,
        /// Polygon and Square tools draw its sides as segments of their own: the polygon's
        /// outline is transparent in the default style) carried over to the sides of its
        /// image, in their style and with their marks. Without them the image was a shape
        /// with no outline. For the transformations that map segments to segments (not a
        /// circle, which doesn't take a polygon: <see cref="CanFigureBeMirrorForSource"/>).
        /// </summary>
        static void AddSideSegments(Drawing drawing, IFigure source, IFigure image, List<IFigure> result)
        {
            // (built on its vertices; a regular polygon is built on its center and draws its
            // sides itself)
            if (!(source is Polygon || source is Polyline) || drawing == null)
            {
                return;
            }

            var vertices = source.Dependencies.ToList();
            var images = image.Dependencies.ToList();
            if (vertices.Count != images.Count || !vertices.All(v => v is IPoint) || !images.All(v => v is IPoint))
            {
                return;
            }

            int count = vertices.Count;
            int sides = source is Polygon ? count : count - 1;
            for (int i = 0; i < sides; i++)
            {
                var segment = FindSideSegment(drawing, vertices[i], vertices[(i + 1) % count]);
                if (segment == null)
                {
                    continue;
                }

                var copy = Factory.CreateSegment(drawing, (IPoint)images[i], (IPoint)images[(i + 1) % count]);
                copy.Style = segment.Style;
                if (segment.Decoration != SegmentDecoration.None)
                {
                    copy.Decoration = segment.Decoration;
                }

                // before the image, which callers take to be the last figure (segments are
                // drawn over polygons whatever the order)
                result.Insert(result.IndexOf(image), copy);
            }
        }

        /// <summary>
        /// The parts of a composite's image (a regular polygon's sides and inside) go over to
        /// the image's dependencies. The clone made them when it was read, on what the
        /// source is built on, and setting the image's dependencies doesn't reach them: two
        /// sides and the inside of a reflected regular pentagon ran to the source's first
        /// vertex. Not in the drawing yet, the parts are registered with nothing (see
        /// DependentPolygonBase.RegisterPart), so their lists are all there is to change.
        /// </summary>
        static void MovePartsOver(IFigure image, IList<IFigure> sourceDependencies)
        {
            if (!(image is CompositeFigure composite))
            {
                return;
            }

            var imageDependencies = image.Dependencies;
            foreach (var part in composite.Children)
            {
                var rewired = part.Dependencies
                    .Select(dependency =>
                    {
                        int index = sourceDependencies.IndexOf(dependency);
                        return index >= 0 && index < imageDependencies.Count ? imageDependencies[index] : dependency;
                    })
                    .ToList();
                if (!rewired.SequenceEqual(part.Dependencies))
                {
                    part.Dependencies = rewired;
                }
            }
        }

        /// <summary>A visible segment from one vertex to the other, either way round</summary>
        static Segment FindSideSegment(Drawing drawing, IFigure vertex1, IFigure vertex2)
        {
            return drawing.Figures
                .OfType<Segment>()
                .FirstOrDefault(segment => segment.GetType() == typeof(Segment)
                    && segment.Visible
                    && segment.Dependencies.Count == 2
                    && segment.Dependencies.Contains(vertex1)
                    && segment.Dependencies.Contains(vertex2));
        }

        /// <summary>
        /// A circle reflects (inverts) a point, and traces anything a point can be on
        /// (<see cref="CanBeTraced"/>): not a polygon, whose image would not be a polygon
        /// </summary>
        public static bool CanFigureBeMirrorForSource(IFigure figure, IFigure source)
        {
            if (figure is IPoint || figure is ILine)
            {
                return true;
            }
            else if (figure is ICircle)
            {
                return source is IPoint || CanBeTraced(source);
            }
            return false;
        }

        /// <summary>Whether the image is a <see cref="Locus"/> (<see cref="CreateTracedImage"/>) rather than a figure of the source's kind</summary>
        static bool IsTraced(IFigure source, bool keepsLengths = true)
        {
            return !(source is IPoint) && !CanBeTransformSource(source, keepsLengths) && CanBeTraced(source);
        }

        /// <param name="sideSegments">Whether the segments along a polygon's sides come along (<see cref="AddSideSegments"/>)</param>
        public static List<IFigure> CreateReflectedFigure(Drawing drawing, IFigure source, IFigure mirror, bool sideSegments = true)
        {
            Check.NotNull(source, "source");
            Check.NotNull(mirror, "mirror");

            if (IsTraced(source) || (mirror is ICircle && CanBeTraced(source)))
            {
                return CreateTracedImage(drawing, source, point => CreateReflectedFigure(drawing, point, mirror));
            }

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

                // (a point made just now for a traced image has no name yet)
                if (!string.IsNullOrEmpty(source.Name))
                {
                    reflectedPoint.Name = source.Name + "'";
                }

                result.Add(reflectedPoint);
            }
            else if ((source is ILine || source is IEllipse || source is IPolygonalChain) && !(mirror is ICircle))
            {
                var dependencies = new List<IFigure>();
                foreach (var dependency in source.Dependencies)
                {
                    if (IsLength(source, dependency))
                    {
                        dependencies.Add(dependency);
                        continue;
                    }

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
                MovePartsOver(reflected, source.Dependencies);
                result.Add(reflected);
                if (sideSegments)
                {
                    AddSideSegments(drawing, source, reflected, result);
                }

                // Flip the wind of arcs.
                var arc = reflected as IArc;
                if (arc != null && mirror is ILine)
                {
                    arc.Clockwise = !arc.Clockwise;
                }

            }
            return result;
        }

        /// <summary>
        /// The factor is a figure the points depend on: a Number holding a typed one, shared by
        /// every point, or anything with a length; with a second length it is their ratio.
        /// </summary>
        /// <param name="sideSegments">Whether the segments along a polygon's sides come along (<see cref="AddSideSegments"/>)</param>
        public static List<IFigure> CreateDilatedFigure(
            Drawing drawing,
            IFigure source,
            IFigure center,
            IFigure lengthProvider1,
            IFigure lengthProvider2,
            bool sideSegments = true)
        {
            Check.NotNull(source, "source");
            Check.NotNull(center, "center");
            Check.NotNull(lengthProvider1, "lengthProvider1");

            if (IsTraced(source, keepsLengths: false))
            {
                return CreateTracedImage(drawing, source, point => CreateDilatedFigure(drawing, point, center, lengthProvider1, lengthProvider2));
            }

            var list = new List<IFigure>() { source, center, lengthProvider1 };
            if (lengthProvider2 != null)
            {
                list.Add(lengthProvider2);
            }

            List<IFigure> result = new List<IFigure>();
            if (source is IPoint)
            {
                var dilatedPoint = Factory.CreateDilatedPoint(drawing, list);
                if (dilatedPoint == null)
                {
                    throw "dilatedPoint is null. source = {0}, center = {1}, segment1 = {2}, segment2 = {3}"
                        .Format(source, center, lengthProvider1, lengthProvider2)
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
                    var dilatedDependency = CreateDilatedFigure(drawing, dependency, center, lengthProvider1, lengthProvider2);
                    if (dilatedDependency == null)
                    {
                        throw "dilatedDependency is null. dependency = {0}, center = {1}, segment1 = {2}, segment2 = {3}"
                            .Format(dependency, center, lengthProvider1, lengthProvider2)
                            .AsException();
                    }
                    if (dilatedDependency.IsEmpty())
                    {
                        throw "dilatedDependency is empty. dependency = {0}, center = {1}, segment1 = {2}, segment2 = {3}"
                            .Format(dependency, center, lengthProvider1, lengthProvider2)
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
                MovePartsOver(dilated, source.Dependencies);
                result.Add(dilated);
                if (sideSegments)
                {
                    AddSideSegments(drawing, source, dilated, result);
                }
            }
            return result;
        }

        /// <summary>
        /// The angle is a figure the points depend on: a Number holding a typed one, shared by
        /// every point of the rotated figure, or anything with an angle.
        /// </summary>
        /// <param name="sideSegments">Whether the segments along a polygon's sides come along (<see cref="AddSideSegments"/>)</param>
        public static List<IFigure> CreateRotatedFigure(
            Drawing drawing,
            IFigure source,
            IFigure center,
            IFigure angleProvider,
            bool sideSegments = true)
        {
            Check.NotNull(source, "source");
            Check.NotNull(center, "center");
            Check.NotNull(angleProvider, "angleProvider");

            if (IsTraced(source))
            {
                return CreateTracedImage(drawing, source, point => CreateRotatedFigure(drawing, point, center, angleProvider));
            }

            var list = new List<IFigure>() { source, center, angleProvider };

            List<IFigure> result = new List<IFigure>();
            if (source is IPoint)
            {
                var rotatedPoint = Factory.CreateRotatedPoint(drawing, list);
                if (rotatedPoint == null)
                {
                    throw "rotatedPoint is null. source = {0}, center = {1}, angle = {2}"
                        .Format(source, center, angleProvider)
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
                    if (IsLength(source, dependency))
                    {
                        dependencies.Add(dependency);
                        continue;
                    }

                    var rotatedDependency = CreateRotatedFigure(drawing, dependency, center, angleProvider);
                    if (rotatedDependency == null)
                    {
                        throw "rotatedDependency is null. dependency = {0}, center = {1}, angle = {2}"
                            .Format(dependency, center, angleProvider)
                            .AsException();
                    }
                    if (rotatedDependency.IsEmpty())
                    {
                        throw "rotatedDependency is empty. dependency = {0}, center = {1}, angle = {2}"
                            .Format(dependency, center, angleProvider)
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
                MovePartsOver(rotated, source.Dependencies);
                result.Add(rotated);
                if (sideSegments)
                {
                    AddSideSegments(drawing, source, rotated, result);
                }
            }
            return result;
        }

        /// <summary>
        /// The sources are shared by every point of a translated figure: one Number for a
        /// typed distance, whatever the source is.
        /// </summary>
        /// <param name="sideSegments">Whether the segments along a polygon's sides come along (<see cref="AddSideSegments"/>)</param>
        public static List<IFigure> CreateTranslatedFigure(
            Drawing drawing,
            IFigure source,
            IFigure distanceSource,
            IFigure directionSource,
            bool sideSegments = true)
        {
            Check.NotNull(source, "source");
            if (IsTraced(source))
            {
                return CreateTracedImage(drawing, source, point => CreateTranslatedFigure(drawing, point, distanceSource, directionSource));
            }

            List<IFigure> result = new List<IFigure>();
            if (source is IPoint)
            {
                var translatedPoint = Factory.CreateTranslatedPoint(drawing, (IPoint)source, distanceSource, directionSource);
                translatedPoint.Visible = source.Visible;
                result.Add(translatedPoint);
            }
            else if (source is ILine || source is IEllipse || source is IPolygonalChain)
            {
                var dependencies = new List<IFigure>();
                foreach (var dependency in source.Dependencies)
                {
                    if (IsLength(source, dependency))
                    {
                        dependencies.Add(dependency);
                        continue;
                    }

                    var translatedDependency = CreateTranslatedFigure(drawing, dependency, distanceSource, directionSource);
                    if (translatedDependency.IsEmpty())
                    {
                        throw "translatedDependency is empty. dependency = {0}, distance = {1}, direction = {2}"
                            .Format(dependency, distanceSource, directionSource)
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
                MovePartsOver(translated, source.Dependencies);
                result.Add(translated);
                if (sideSegments)
                {
                    AddSideSegments(drawing, source, translated, result);
                }
            }
            return result;
        }
    }
}