using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
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

            // the figures first, aside, to see which styles they name: only those are written
            // (a figure may name another's style - a vector its arrow's); loading adds back the
            // default ones (StyleManager.AddWithDefaults)
            var figureList = new XDocument();
            using (var figureWriter = figureList.CreateWriter())
            {
                WriteFigureList(figures, figureWriter);
            }

            var usedStyles = new HashSet<string>(figureList
                .Descendants()
                .Attributes("Style")
                .Select(a => a.Value));
            WriteStyles(drawing, usedStyles, writer);
            figureList.Root.WriteTo(writer);
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

            // the paper: a solid color is the Color attribute (the theme's paper, the default,
            // is left out), a gradient is a Background child element, as a gradient fill of a
            // style is; the paper chosen for another theme is a child element named after the
            // theme, with the same Color or Background inside (empty: that theme's paper)
            var background = drawing.OwnBackground == null ? null : BrushSerializer.WriteBrush(drawing.OwnBackground);
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

            if (drawing.CoordinateSystem.GridStep > 0)
            {
                writer.WriteAttributeDouble("GridStep", drawing.CoordinateSystem.GridStep);
            }

            if (background is System.Xml.Linq.XElement gradient)
            {
                writer.WriteStartElement("Background");
                gradient.WriteTo(writer);
                writer.WriteEndElement();
            }

            foreach (var theme in AppTheme.All)
            {
                if (drawing.Overrides.TryGetValue(theme.Name, out var values) && values.TryGetValue(nameof(Drawing.Background), out var paper))
                {
                    writer.WriteStartElement(theme.Name);
                    var themed = paper == null ? null : BrushSerializer.WriteBrush((Brush)paper);
                    if (themed is string themedColor)
                    {
                        writer.WriteAttributeString("Color", themedColor);
                    }
                    else if (themed is System.Xml.Linq.XElement themedGradient)
                    {
                        writer.WriteStartElement("Background");
                        themedGradient.WriteTo(writer);
                        writer.WriteEndElement();
                    }

                    writer.WriteEndElement();
                }
            }

            writer.WriteEndElement();
        }

        /// <summary>
        /// The styles the figures name, and a default style the drawing changed (a figure on
        /// the default style of its kind doesn't name it); a default as a new drawing has it
        /// is left out, the loader adds it back
        /// </summary>
        public virtual void WriteStyles(Drawing drawing, ISet<string> usedStyles, XmlWriter writer)
        {
            var defaults = StyleManager.CreateDefaultStyles();
            writer.WriteStartElement("Styles");
            foreach (var style in drawing.StyleManager.GetAllStyles())
            {
                bool isDefaultName = defaults.Any(candidate => candidate.Name == style.Name);
                bool write = isDefaultName
                    ? !StyleManager.IsUnchangedDefault(style, defaults)
                    : usedStyles.Contains(style.Name);
                if (write)
                {
                    WriteStyle(style, writer, defaults.FirstOrDefault(candidate => candidate.Name == style.Name));
                }
            }

            writer.WriteEndElement();
        }

        /// <param name="original">The default style of the same name as a new drawing has it, if the style is one</param>
        public virtual void WriteStyle(IFigureStyle style, XmlWriter writer, IFigureStyle original = null)
        {
            writer.WriteStartElement(GetStyleElementName(style));

            // the name first, so a file reads as a list of named styles; the rest stay in the
            // order reflection gives (the style's own properties, then its base classes'), and
            // a value a fresh style has anyway (a solid dash, a filled shape) is left out
            var fresh = (IFigureStyle)Activator.CreateInstance(style.GetType());
            var values = valueDiscovery.GetValues(style)
                .OrderBy(v => v.Name == "Name" ? 0 : 1)
                .Where(v => v.Name == "Name" || !IsFreshValue(v, fresh));
            WriteValues(values, writer);

            // what differs under another theme, in a child element named after it
            if (style is FigureStyle figureStyle)
            {
                foreach (var theme in AppTheme.All)
                {
                    var overrides = figureStyle.OverrideValues(theme.Name).ToArray();
                    if (overrides.Length > 0)
                    {
                        writer.WriteStartElement(theme.Name);
                        WriteValues(overrides, writer);
                        writer.WriteEndElement();
                    }
                    else if (original is FigureStyle originalStyle
                        && originalStyle.OverrideValues(theme.Name).Any()
                        && originalStyle.GetBaseSignature() == figureStyle.GetBaseSignature())
                    {
                        // A default style made to look under the theme as it does under
                        // the base one ("Same as in Light"), and otherwise unchanged, says
                        // so with an empty element: without any, the loader takes it for
                        // an old file's copy of the default and puts the theme's look back.
                        writer.WriteStartElement(theme.Name);
                        writer.WriteEndElement();
                    }
                }
            }

            writer.WriteEndElement();
        }

        /// <summary>Simple values as attributes; structured ones (a gradient brush) as child elements named after the property, after all attributes</summary>
        void WriteValues(IEnumerable<IValueProvider> values, XmlWriter writer)
        {
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
        }

        static bool IsFreshValue(IValueProvider value, IFigureStyle fresh)
        {
            var freshValue = PropertyDiscoveryStrategy.CreateValueProvider(fresh, value.Name);
            return Equals(
                SerializationService.Instance.Write(value)?.ToString(),
                SerializationService.Instance.Write(freshValue)?.ToString());
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

            // a part of a figure (a vertex a regular polygon works out) has no name of its
            // own: the figure's name, and which part
            var owner = FindPartOwner(dependency);
            if (owner != null)
            {
                writer.WriteAttributeString("Name", owner.Name);
                writer.WriteAttributeString("Part", owner.GetPartName(dependency));
            }
            else
            {
                writer.WriteAttributeString("Name", dependency.Name);
            }

            writer.WriteEndElement();
        }

        HashSet<IFigure> drawingFigures;

        /// <summary>The figure the dependency is a part of, if it is not a figure of the drawing itself</summary>
        IFigureParts FindPartOwner(IFigure dependency)
        {
            var drawing = dependency.Drawing;
            if (drawing == null)
            {
                return null;
            }

            if (drawingFigures == null)
            {
                drawingFigures = new HashSet<IFigure>(drawing.Figures);
            }

            if (drawingFigures.Contains(dependency))
            {
                return null;
            }

            return drawing.Figures
                .OfType<IFigureParts>()
                .FirstOrDefault(figure => figure.GetPartName(dependency) != null);
        }

        protected virtual string GetTagNameForFigure(IFigure figure)
        {
            return figure.GetType().Name;
        }
    }
}
