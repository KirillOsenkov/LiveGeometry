using System;
using System.Collections.Generic;
using System.Xml;

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

        public static string WriteMacroToString(IList<IFigure> inputs, IList<IFigure> results)
        {
            return new MacroSerializer()
            {
                Inputs = inputs,
                Results = results
            }.WriteMacroToString();
        }

        public void Write(XmlWriter writer)
        {
            writer.WriteStartDocument();
            writer.WriteStartElement("Macro");
            writer.WriteAttributeString("Name", "Custom tool");
            WriteInputs(writer);
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

        void WriteResults(XmlWriter writer)
        {
            WriteFigureList(Results, writer);
        }
    }
}
