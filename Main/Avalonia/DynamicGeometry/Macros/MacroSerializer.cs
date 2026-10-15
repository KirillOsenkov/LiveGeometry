using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace DynamicGeometry
{
    public class MacroSerializer : DrawingSerializer
    {
        public IList<IFigure> Inputs { get; set; }
        public IList<IFigure> Results { get; set; }

        public string WriteMacroToString()
        {
            return WriteUsingXmlWriter(Write);
        }

        /// <summary>The name of the tool, on its button</summary>
        public string Name { get; set; } = "Custom tool";

        public static string WriteMacroToString(IList<IFigure> inputs, IList<IFigure> results, string name)
        {
            return new MacroSerializer()
            {
                Inputs = inputs,
                Results = results,
                Name = name
            }.WriteMacroToString();
        }

        public void Write(XmlWriter writer)
        {
            writer.WriteStartDocument();
            writer.WriteStartElement("Macro");

            // the figures are written as a drawing's: the same version says how to read them
            writer.WriteAttributeDouble("Version", Settings.CurrentDrawingVersion);
            writer.WriteAttributeString("Name", Name);
            WriteInputs(writer);
            MacroIcon.Write(writer, Inputs, Results);
            WriteResults(writer);
            writer.WriteEndElement();
            writer.WriteEndDocument();
        }

        void WriteInputs(XmlWriter writer)
        {
            writer.WriteStartElement("Inputs");
            foreach (var input in Inputs)
            {
                WriteInput(input, writer);
            }
            writer.WriteEndElement();
        }

        void WriteInput(IFigure input, XmlWriter writer)
        {
            writer.WriteStartElement("Input");
            writer.WriteAttributeString("Name", InputNames[input]);
            writer.WriteAttributeString("Type", GetInputType(input));
            writer.WriteEndElement();
        }

        Dictionary<IFigure, string> inputNames;

        /// <summary>
        /// The name each input goes by in the macro: its own, or for a part of a figure (a
        /// vertex a regular polygon works out), which has none, what it is ("vertex 2 of
        /// regular pentagon p", which the tool's hint shows). A part was an input without a
        /// name, and what was built on it named it by its polygon, which the tool was never
        /// given: the tool made nothing.
        /// </summary>
        Dictionary<IFigure, string> InputNames
        {
            get
            {
                if (inputNames == null)
                {
                    inputNames = new Dictionary<IFigure, string>();
                    var taken = new HashSet<string>(Results.Select(result => result.Name));
                    foreach (var input in Inputs)
                    {
                        var name = input.Name;
                        if (name.IsEmpty() || taken.Contains(name))
                        {
                            var description = input.ToString();
                            description = description.IsEmpty() ? "input" : char.ToLowerInvariant(description[0]) + description.Substring(1);
                            name = description;
                            for (int number = 2; taken.Contains(name); number++)
                            {
                                name = description + " " + number;
                            }
                        }

                        taken.Add(name);
                        inputNames[input] = name;
                    }
                }

                return inputNames;
            }
        }

        /// <summary>An input by the name it has in the macro, a part of a figure too</summary>
        protected override void WriteDependency(IFigure dependency, XmlWriter writer)
        {
            if (Inputs.Contains(dependency))
            {
                writer.WriteStartElement("Dependency");
                writer.WriteAttributeString("Name", InputNames[dependency]);
                writer.WriteEndElement();
                return;
            }

            base.WriteDependency(dependency, writer);
        }

        static List<Type> commonTypes = new List<Type>()
        {
            typeof(IPoint), typeof(ILine), typeof(IEllipse)
        };
        
        /// <summary>
        /// What the tool asks for: a point, a line or an ellipse of any kind - unless a result
        /// takes more from the input than its kind gives (<see cref="TakesMoreThanKind"/>), or
        /// an expression reads a number off it that not every figure of the kind has (u.X of a
        /// vector: given a segment, the point at (CD.X, CD.Y) was nowhere; a point's X and Y
        /// every point has), when only the input's own type will do
        /// </summary>
        string GetInputType(IFigure input)
        {
            Type inputType = input.GetType();
            if (TakesMoreThanKind(input))
            {
                return inputType.Name;
            }

            foreach (var commonType in commonTypes)
            {
                if (commonType.IsAssignableFrom(inputType))
                {
                    // what every figure of the kind has: a point's X and Y (PointBase), nothing more for a line or an ellipse
                    var sharedType = commonType == typeof(IPoint) ? typeof(PointBase) : commonType;
                    return input.IsReadByExpressionsBeyond(sharedType, Results.Contains)
                        ? inputType.Name
                        : commonType.Name;
                }
            }
            return inputType.Name;
        }

        /// <summary>
        /// Whether a result is built on the line or ellipse for more than being one: its
        /// length (a radius, a distance, a factor, a translation's distance), its area or
        /// perimeter, its inside, a circle to reflect in. A tool defined on a circle with a
        /// segment for its radius asked for any line, and a line given made a circle with no
        /// radius, which threw and stayed in the drawing. A point is a point to every figure.
        /// </summary>
        bool TakesMoreThanKind(IFigure input)
        {
            return !(input is IPoint) && input.Dependents.Any(result => Results.Contains(result)
                && (result is LengthMeasurement
                    || result is CircleByRadius
                    || result is DilatedPoint
                    || result is TranslatedPoint translated && translated.DistanceSource == input
                    || result is ReflectedPoint && input is IEllipse
                    || result is AreaMeasurement
                    || result is PointInRegion));
        }

        /// <summary>
        /// The figures with the styles of their own that they name, as for the clipboard: used
        /// in another drawing (or in a later run), a figure would look its style up there by
        /// name, and a catenary made by the tool lost its rope's look
        /// </summary>
        void WriteResults(XmlWriter writer)
        {
            var drawing = Inputs.Concat(Results).Select(f => f.Drawing).FirstOrDefault(d => d != null);
            if (drawing == null)
            {
                WriteFigureList(Results, writer);
                return;
            }

            var withStyles = new XDocument();
            using (var aside = withStyles.CreateWriter())
            {
                WriteFiguresWithStyles(drawing, Results, aside, DefaultStyleNames.Changed);
            }

            // A name that reads like the default is the default, which says nothing once the
            // figure is built on other points: segment AB on P and Q kept AB as if typed. The
            // figures that had their default names say so, and get the new ones.
            var figures = withStyles.Root.Element("Figures");
            foreach (var element in figures?.Elements() ?? Enumerable.Empty<XElement>())
            {
                var result = Results.FirstOrDefault(r => r.Name == element.ReadString("Name"));
                if (result != null && result.HasDefaultName)
                {
                    element.SetAttributeValue("DefaultName", "true");
                }
            }

            foreach (var element in withStyles.Root.Elements())
            {
                element.WriteTo(writer);
            }
        }
    }
}
