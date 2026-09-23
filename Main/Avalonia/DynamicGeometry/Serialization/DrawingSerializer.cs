using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using Avalonia.Media;

namespace DynamicGeometry
{
    public partial class DrawingSerializer
    {
#if !SILVERLIGHT
        public static void Save(Drawing drawing, string fileName)
        {
            string serialized = SaveDrawing(drawing);
            File.WriteAllText(fileName, serialized);
        }
#endif
        public static string SaveDrawing(Drawing drawing)
        {
            return WriteUsingXmlWriter(w => SaveDrawing(drawing, w));
        }

        public static void SaveDrawing(Drawing drawing, XmlWriter writer)
        {
            DrawingSerializer serializer = new DrawingSerializer();
            serializer.Write(drawing, writer);
        }

        public static void SaveDrawing(Drawing drawing, Stream stream)
        {
            using (var writer = XmlWriter.Create(stream, XmlSettings))
            {
                SaveDrawing(drawing, writer);
            }
        }

        public static void SaveFigures(IEnumerable<IFigure> figures, XmlWriter writer)
        {
            DrawingSerializer serializer = new DrawingSerializer();
            serializer.WriteFigures(figures, writer);
        }

        static XmlWriterSettings XmlSettings
        {
            get
            {
                return new XmlWriterSettings()
                {
                    Indent = true,
                    Encoding = Encoding.UTF8,
                    CloseOutput = true
                };
            }
        }

        public static string WriteUsingXmlWriter(Action<XmlWriter> writerConsumer)
        {
            // through a writer that says UTF-8: with a plain StringBuilder the declaration
            // would say utf-16, and the file it ends up in is UTF-8 - strict XML parsers refuse that
            var text = new Utf8StringWriter();
            using (var w = XmlWriter.Create(text, XmlSettings))
            {
                writerConsumer(w);
            }

            return text.ToString();
        }

        class Utf8StringWriter : StringWriter
        {
            public override Encoding Encoding => Encoding.UTF8;
        }

        void Write(Drawing drawing, XmlWriter writer)
        {
            var figures = drawing.GetSerializableFigures();
            writer.WriteStartDocument();
            writer.WriteStartElement("Drawing");
            writer.WriteAttributeDouble("Version", drawing.Version);
            writer.WriteAttributeString("Creator", Avalonia.Application.Current.ToString());
            WriteCoordinateSystem(drawing, writer);
            foreach (var scene in drawing.Scenes)
            {
                writer.WriteStartElement("Scene");
                writer.WriteAttributeDouble("Left", scene.X);
                writer.WriteAttributeDouble("Top", scene.Bottom);
                writer.WriteAttributeDouble("Right", scene.Right);
                writer.WriteAttributeDouble("Bottom", scene.Y);
                writer.WriteEndElement();
            }

            WriteStyles(drawing, writer);
            WriteFigureList(figures, writer);
            writer.WriteEndElement();
            writer.WriteEndDocument();
        }

        void WriteCoordinateSystem(Drawing drawing, XmlWriter writer)
        {
            writer.WriteStartElement("Viewport");
            writer.WriteAttributeDouble("Left", drawing.CoordinateSystem.MinimalVisibleX);
            writer.WriteAttributeDouble("Top", drawing.CoordinateSystem.MaximalVisibleY);
            writer.WriteAttributeDouble("Right", drawing.CoordinateSystem.MaximalVisibleX);
            writer.WriteAttributeDouble("Bottom", drawing.CoordinateSystem.MinimalVisibleY);

            // the paper: a solid color is the Color attribute (white, the default, is left out),
            // a gradient is a Background child element, as a gradient fill of a style is
            var background = Drawing.IsWhite(drawing.Background) ? null : BrushSerializer.WriteBrush(drawing.Background);
            if (background is string color)
            {
                writer.WriteAttributeString("Color", color);
            }

            if (drawing.CoordinateGrid.Locked)
            {
                writer.WriteAttributeBool("Locked", true);
            }

            if (drawing.CoordinateGrid.Visible)
            {
                writer.WriteAttributeBool("Grid", true);
                writer.WriteAttributeBool("Axes", drawing.CoordinateGrid.ShowAxes);
            }

            if (background is System.Xml.Linq.XElement gradient)
            {
                writer.WriteStartElement("Background");
                gradient.WriteTo(writer);
                writer.WriteEndElement();
            }

            writer.WriteEndElement();
        }

        public virtual void WriteStyles(Drawing drawing, XmlWriter writer)
        {
            writer.WriteStartElement("Styles");
            foreach (var style in drawing.StyleManager.GetAllStyles())
            {
                WriteStyle(style, writer);
            }

            writer.WriteEndElement();
        }

        public virtual void WriteStyle(IFigureStyle style, XmlWriter writer)
        {
            writer.WriteStartElement(GetStyleElementName(style));
            var values = valueDiscovery.GetValues(style);

            // Simple values are attributes; structured ones (a gradient brush) are child
            // elements named after the property, and have to come after all attributes.
            var elements = new List<KeyValuePair<string, System.Xml.Linq.XElement>>();
            foreach (var value in values)
            {
                var serialized = SerializationService.Instance.Write(value);
                if (serialized is System.Xml.Linq.XElement element)
                {
                    elements.Add(new KeyValuePair<string, System.Xml.Linq.XElement>(value.Name, element));
                }
                else if (serialized != null)
                {
                    writer.WriteAttributeString(value.Name, serialized.ToString());
                }
            }

            foreach (var pair in elements)
            {
                writer.WriteStartElement(pair.Key);
                pair.Value.WriteTo(writer);
                writer.WriteEndElement();
            }

            writer.WriteEndElement();
        }

        IValueDiscoveryStrategy valueDiscovery = new IncludeByDefaultValueDiscoveryStrategy();

        string GetStyleElementName(IFigureStyle style)
        {
            return style.GetType().Name;
        }

        public virtual void WriteFigureList(IEnumerable<IFigure> list, XmlWriter writer)
        {
            writer.WriteStartElement("Figures");
            WriteFigures(list, writer);
            writer.WriteEndElement();
        }

        protected virtual void WriteFigures(IEnumerable<IFigure> list, XmlWriter writer)
        {
            foreach (var figure in list)
            {
                WriteFigure(figure, writer);
            }
        }

        protected virtual void WriteFigure(IFigure figure, XmlWriter writer)
        {
            writer.WriteStartElement(GetTagNameForFigure(figure));
            writer.WriteAttributeString("Name", figure.Name);
            figure.WriteXml(writer);
            WriteDependencies(figure, writer);
            writer.WriteEndElement();
        }

        protected virtual void WriteDependencies(IFigure figure, XmlWriter writer)
        {
            if (figure.Dependencies.IsEmpty()) return;

            foreach (var dependency in figure.Dependencies)
            {
                WriteDependency(dependency, writer);
            }
        }

        protected virtual void WriteDependency(IFigure dependency, XmlWriter writer)
        {
            writer.WriteStartElement("Dependency");
            writer.WriteAttributeString("Name", dependency.Name);
            writer.WriteEndElement();
        }

        protected virtual string GetTagNameForFigure(IFigure figure)
        {
            return figure.GetType().Name;
        }
    }
}
