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
            writer.WriteAttributeString("Name", input.Name);
            writer.WriteAttributeString("Type", GetInputType(input));
            writer.WriteEndElement();
        }

        static List<Type> commonTypes = new List<Type>()
        {
            typeof(IPoint), typeof(ILine), typeof(IEllipse)
        };
        
        string GetInputType(IFigure input)
        {
            Type inputType = input.GetType();
            foreach (var commonType in commonTypes)
            {
                if (commonType.IsAssignableFrom(inputType))
                {
                    return commonType.Name;
                }
            }
            return inputType.Name;
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
                WriteFiguresWithStyles(drawing, Results, aside);
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
