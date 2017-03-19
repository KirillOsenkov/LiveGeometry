using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;
using System.Xml;
using System.Windows.Media;

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
            var sb = new StringBuilder();
            using (var w = XmlWriter.Create(sb, XmlSettings))
            {
                writerConsumer(w);
            }

            return sb.ToString();
        }

        void Write(Drawing drawing, XmlWriter writer)
        {
            var figures = drawing.GetSerializableFigures();
            writer.WriteStartDocument();
            writer.WriteStartElement("Drawing");
            writer.WriteAttributeDouble("Version", drawing.Version);
            writer.WriteAttributeString("Creator", System.Windows.Application.Current.ToString());
            WriteCoordinateSystem(drawing, writer);
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

            var backgroundBrush = drawing.Canvas.Background as SolidColorBrush;
            if (backgroundBrush != null && backgroundBrush.Color != Colors.White)
            {
                writer.WriteAttributeString("Color", backgroundBrush.Color.ToString());
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
            foreach (var value in values)
            {
                var serialized = MEFHost.Instance.SerializationService.Write(value);
                if (serialized != null)
                {
                    writer.WriteAttributeString(value.Name, serialized.ToString());
                }
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
