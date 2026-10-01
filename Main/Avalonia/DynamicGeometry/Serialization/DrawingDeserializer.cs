using System;
using System.Collections.Generic;
using System.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using System.Xml.Linq;

namespace DynamicGeometry
{
    public partial class DrawingDeserializer
    {
        public static Drawing OpenDrawing(Canvas canvas, XElement element)
        {
            Drawing drawing = new Drawing(canvas);
            DrawingDeserializer deserializer = new DrawingDeserializer();
            deserializer.ReadDrawing(drawing, element);

            if (!deserializer.IsSuccess)
            {
                throw new Exception(deserializer.GetErrorReport());
            }

            return drawing;
        }

        public static Drawing OpenDrawing(Canvas canvas, string savedDrawing)
        {
            Check.NotEmpty(savedDrawing);

            XElement element = XElement.Parse(savedDrawing);
            return OpenDrawing(canvas, element);
        }

        public virtual void ReadDrawing(Drawing drawing, XElement element)
        {
            Check.NotNull(drawing, "drawing");
            Check.NotNull(element, "element");
            drawing.Version = element.ReadDouble("Version");    // Defaults to 0 if Version attribute does not exist.
            ReadStyles(drawing, element);
            var figuresNode = element.Element("Figures");
            if (figuresNode == null)
            {
                // Perhaps notify user that no figures were found.
            }
            else
            {
                // under the names the file gives them until all are in: a label's expression
                // is compiled by those names as the label comes in
                drawing.KeepsNamesAsRead = true;
                try
                {
                    var figures = ReadFigures(figuresNode, drawing);
                    foreach (var figure in figures)
                    {
                        Actions.Add(drawing, figure);
                    }
                }
                finally
                {
                    drawing.KeepsNamesAsRead = false;
                }
            }
            ReadViewport(drawing, element);
            ReadScenes(drawing, element);
            DrawingUpdater updater;
#if TABULA
            updater = new TABDrawingUpdater();
#else
            updater = new DrawingUpdater();
#endif
            updater.UpdateIfNecessary(drawing);
            drawing.Recalculate();

            // Files don't record which version of the library wrote them, so a drawing that is
            // known to be that old has to say so (the gallery drawings from the phone do). Saving it
            // writes the upgraded algorithms and not the attribute.
            if (element.ReadString("IntersectionOrder") == "Legacy")
            {
                // in the order of construction: what comes later is built on what was fixed
                foreach (var intersection in drawing.Figures.OfType<IntersectionPoint>().ToArray())
                {
                    if (intersection.UpgradeLegacyCircleAndLineOrder())
                    {
                        drawing.Recalculate();
                    }
                }
            }

            // Every figure used to be numbered by its type (Segment1); one nobody renamed takes
            // the name of its points now (2026-09-27), and of two named after the same points
            // the first in the file is AB, the next AB2
            foreach (var figure in drawing.Figures.ToArray())
            {
                (figure as FigureBase)?.UpdateDefaultName();
                FigureBase.SettleDefaultNames(drawing, figure);
            }

            // A typed distance or direction of a translated point used to be an attribute of
            // the point; it is a Number the point depends on now (2026-09-25)
            foreach (var translated in drawing.Figures.OfType<TranslatedPoint>().ToArray())
            {
                translated.UpgradeLegacyValues();
            }

            // Version 1 (2026-09-23): the offset of a label from what it labels is in pixels,
            // not units of the plane. Here, after the viewport, so that the conversion happens
            // at the zoom the file opens at. Saving writes the pixels and the new version.
            if (drawing.Version < 1)
            {
                foreach (var label in drawing.Figures.OfType<LabelWithOffset>())
                {
                    label.UpgradeOffsetFromUnits();
                }

                drawing.Version = Settings.CurrentDrawingVersion;
                drawing.Recalculate();
            }
            //drawing.CoordinateSystem.MoveTo(drawing.Figures.OfType<IPoint>().Midpoint().Minus());
        }

        List<string> errors = new List<string>();

        public void ReportError(string error)
        {
            if (System.Diagnostics.Debugger.IsAttached)
            {
                throw new Exception(error);
            }

            errors.Add(error);
        }

        public bool IsSuccess
        {
            get
            {
                return errors.Count == 0;
            }
        }

        public string GetErrorReport()
        {
            return string.Join("\n", errors.ToArray());
        }

        private void ReadViewport(Drawing drawing, XElement element)
        {
            var viewportNode = element.Element("Viewport");
            if (viewportNode == null)
            {
                return;
            }

            double minX = viewportNode.ReadDouble("Left");
            double maxX = viewportNode.ReadDouble("Right");
            double minY = viewportNode.ReadDouble("Bottom");
            double maxY = viewportNode.ReadDouble("Top");
            drawing.CoordinateGrid.Locked = viewportNode.ReadBool("Locked", false);
            // a file that doesn't say has no grid: the serializer only writes Grid when it is
            // on. (It used to default to the global setting, which the last drawing shown
            // had set - so a gallery tile loaded after a graph got a grid.)
            drawing.CoordinateGrid.Visible = viewportNode.ReadBool("Grid", false);
            drawing.CoordinateGrid.ShowAxes = viewportNode.ReadBool("Axes", true);
            drawing.CoordinateSystem.GridStep = viewportNode.ReadDouble("GridStep");
            drawing.CoordinateSystem.SetViewport(minX, maxX, minY, maxY);
            string styleName = viewportNode.ReadString("Style");    // Don't know who uses this.  I don't. - David
            if (!styleName.IsEmpty() && drawing.StyleManager != null)
            {
                var style = drawing.StyleManager[styleName];
                if (style != null)
                {
                    var wpfStyle = style.GetWpfStyle();
                    drawing.Canvas.Apply(wpfStyle);
                }
            }
            else if (viewportNode.Element("Background")?.Elements().FirstOrDefault() is XElement gradient)
            {
                drawing.Background = BrushSerializer.ParseBrush(gradient);
            }
            else if (viewportNode.ReadString("Color") != null)
            {
                drawing.Background = new SolidColorBrush(viewportNode.ReadString("Color").ToColor());
            }
            else
            {
                drawing.Background = null; // the theme's paper
            }

            // the paper chosen for another theme: a child element named after the theme, with
            // the same Color or Background inside; empty for that theme's own paper
            foreach (var themeNode in viewportNode.Elements())
            {
                if (AppTheme.ByName(themeNode.Name.LocalName) == null)
                {
                    continue;
                }

                Brush paper = null;
                if (themeNode.Element("Background")?.Elements().FirstOrDefault() is XElement themedGradient)
                {
                    paper = BrushSerializer.ParseBrush(themedGradient);
                }
                else if (themeNode.ReadString("Color") != null)
                {
                    paper = new SolidColorBrush(themeNode.ReadString("Color").ToColor());
                }

                drawing.SetOverride(themeNode.Name.LocalName, nameof(Drawing.Background), paper);
            }
        }

        /// <summary>The suggested views, in the same Left/Top/Right/Bottom form as the viewport</summary>
        static void ReadScenes(Drawing drawing, XElement element)
        {
            drawing.Scenes.Clear();
            foreach (var sceneNode in element.Elements("Scene"))
            {
                double left = sceneNode.ReadDouble("Left");
                double right = sceneNode.ReadDouble("Right");
                double bottom = sceneNode.ReadDouble("Bottom");
                double top = sceneNode.ReadDouble("Top");
                if (right > left && top > bottom)
                {
                    drawing.Scenes.Add(new Rect(left, bottom, right - left, top - bottom));
                }
            }
        }

        public virtual void ReadStyles(Drawing drawing, XElement element)
        {
            var stylesNode = element.Element("Styles");
            drawing.StyleManager.Clear();
            if (stylesNode == null)
            {
                drawing.StyleManager.AddDefaultStyles();
                return;
            }

            // A file has only the styles its figures use; the rest are the defaults. A kind of
            // style this version doesn't have (PolylineStyle of an old file) is left out, and
            // what uses it gets the default of its kind: reading it threw, and nothing of the
            // drawing came in.
            var own = new List<IFigureStyle>();
            foreach (var styleNode in stylesNode.Elements())
            {
                if (!StyleTypes.Contains(styleNode.Name.LocalName))
                {
                    ReportError(string.Format("The style {0} is left out: this version has no style of the kind {1}.", styleNode.ReadString("Name"), styleNode.Name.LocalName));
                    continue;
                }

                var style = ReadStyle(styleNode);
                if (style != null)
                {
                    own.Add(style);
                }
            }

            drawing.StyleManager.AddWithDefaults(own);
        }

        static HashSet<string> styleTypes;

        /// <summary>The names of the kinds of style a file may hold</summary>
        static HashSet<string> StyleTypes
        {
            get
            {
                return styleTypes ??= new HashSet<string>(typeof(DrawingDeserializer).Assembly
                    .GetTypes()
                    .Where(t => typeof(IFigureStyle).IsAssignableFrom(t) && !t.IsAbstract && !t.IsInterface)
                    .Select(t => t.Name));
            }
        }

        private IFigureStyle ReadStyle(XElement styleNode)
        {
            var style = SerializationService.Instance.Read<IFigureStyle>(styleNode);

            // what differs under another theme: a child element named after the theme, its
            // attributes and elements the properties as in the style's own element
            if (style is FigureStyle figureStyle)
            {
                var type = figureStyle.GetType();
                foreach (var themeNode in styleNode.Elements())
                {
                    string theme = themeNode.Name.LocalName;
                    if (AppTheme.ByName(theme) == null)
                    {
                        continue;
                    }

                    // even an empty one: the style is as under the base theme, on purpose
                    figureStyle.SaysThemes = true;
                    foreach (var attribute in themeNode.Attributes())
                    {
                        var value = OverrideValue(figureStyle, type, theme, attribute.Name.LocalName);
                        if (value != null)
                        {
                            SerializationService.Instance.Read(value, attribute.Value);
                        }
                    }

                    foreach (var propertyNode in themeNode.Elements())
                    {
                        var value = OverrideValue(figureStyle, type, theme, propertyNode.Name.LocalName);
                        if (value != null)
                        {
                            SerializationService.Instance.Read(value, propertyNode);
                        }
                    }
                }
            }

            return style;
        }

        static ThemedValue OverrideValue(FigureStyle style, Type type, string theme, string property)
        {
            var propertyInfo = type.GetProperty(property);
            return propertyInfo == null ? null : new ThemedValue(new PropertyValue(propertyInfo, style), style, theme);
        }

        private IValueDiscoveryStrategy valueDiscovery = new IncludeByDefaultValueDiscoveryStrategy();

        public virtual void ReadDrawing(Drawing drawing, string savedDrawing)
        {
            XElement element = XElement.Parse(savedDrawing);
            ReadDrawing(drawing, element);
        }

        public virtual void ReadFigureList(IList<IFigure> figureList, XElement element, Drawing drawing)
        {
            ReadFigureList(figureList, element, drawing, new Dictionary<string, IFigure>());
        }

        /// <param name="byFileName">Filled with the figures read, by the names the file gives them</param>
        public void ReadFigureList(IList<IFigure> figureList, XElement element, Drawing drawing, Dictionary<string, IFigure> byFileName)
        {
            if (element.Name == "Drawing")
            {
                element = element.Element("Figures");
            }

            var figures = ReadFigures(element, drawing, byFileName);
            figureList.AddRange(figures);
        }

        protected virtual IEnumerable<IFigure> ReadFigures(XElement figuresNode)
        {
            return ReadFigures(figuresNode, null);
        }

        public virtual IEnumerable<IFigure> ReadFigures(XElement figuresNode, Drawing drawing)
        {
            Dictionary<string, IFigure> figures = new Dictionary<string, IFigure>();
            return ReadFigures(figuresNode, drawing, figures);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="figuresNode"></param>
        /// <param name="drawing"></param>
        /// <param name="figures">The dictionary is on purpose - we pass an existing dictionary during macro creation</param>
        /// <returns></returns>
        public virtual IEnumerable<IFigure> ReadFigures(XElement figuresNode, Drawing drawing, Dictionary<string, IFigure> figures)
        {
            List<IFigure> result = new List<IFigure>();
            List<string> nameBlacklist = new List<string>();
            Dictionary<string, XElement> nodeMap = new Dictionary<string, XElement>();
            foreach (var figureNode in figuresNode.Elements())
            {
                string name = figureNode.ReadString("Name");
                if (string.IsNullOrEmpty(name))
                {
                    ReportError(figureNode.Name.LocalName + " without a name is left out.");
                }
                else if (nodeMap.ContainsKey(name))
                {
                    ReportError(string.Format("Two figures are called {0}: only the first is read.", name));
                }
                else
                {
                    nodeMap.Add(name, figureNode);
                }
            }

            foreach (var figureName in nodeMap.Keys)
            {
                ReadFigure(figureName, figures, nameBlacklist, nodeMap, drawing, result.Add);
            }

            return result;
        }

        private static Dictionary<string, Type> mFigureTypes;
        public static Dictionary<string, Type> FigureTypes
        {
            get
            {
                if (mFigureTypes == null)
                {
                    mFigureTypes = new Dictionary<string, Type>();
                    var assembly = typeof(DrawingDeserializer).Assembly;
                    foreach (var type in assembly.GetTypes()
                        .Where(t => typeof(IFigure).IsAssignableFrom(t)))
                    {
                        mFigureTypes.Add(type.Name, type);
                    }
                }

                return mFigureTypes;
            }
        }

        public static Type FindType(string typeName)
        {
            if (FigureTypes.ContainsKey(typeName))
            {
                return FigureTypes[typeName];
            }

            return null;
        }

        public virtual void ReadFigure(
            string figureName,
            Dictionary<string, IFigure> alreadyDeserializedFigures,
            List<string> nameBlacklist,
            Dictionary<string, XElement> nodeMap,
            Drawing drawing,
            Action<IFigure> callbackWhenCreated)
        {
            if (alreadyDeserializedFigures.ContainsKey(figureName))
            {
                return;
            }

            // A file that is not whole - a figure of a kind this version doesn't have, one
            // built on a figure the file lacks, two built on each other - is read as far as
            // it goes, and what is left out is said in words (ReportError). It used to throw
            // half way: the message was a bare name, or "The given key was not present in
            // the dictionary", and nothing of the drawing was shown.
            if (!nodeMap.TryGetValue(figureName, out var figureNode) || !beingRead.Add(figureName))
            {
                return;
            }

            try
            {
                ReadFigure(
                    figureName,
                    figureNode,
                    alreadyDeserializedFigures,
                    nameBlacklist,
                    nodeMap,
                    drawing,
                    callbackWhenCreated);
            }
            finally
            {
                beingRead.Remove(figureName);
            }
        }

        // the figures whose dependencies are being read: one that comes up again is built on itself
        readonly HashSet<string> beingRead = new HashSet<string>();

        void ReadFigure(
            string figureName,
            XElement figureNode,
            Dictionary<string, IFigure> alreadyDeserializedFigures,
            List<string> nameBlacklist,
            Dictionary<string, XElement> nodeMap,
            Drawing drawing,
            Action<IFigure> callbackWhenCreated)
        {
            Type type = FindType(figureNode.Name.LocalName);
            if (type == null)
            {
                ReportError(string.Format("{0} is left out: this version has no figure of the kind {1}.", figureName, figureNode.Name.LocalName));
                return;
            }

            var dependencyNodes = figureNode.Elements("Dependency").ToArray();
            var dependencyNames = dependencyNodes.Select(e => e.ReadString("Name")).ToArray();
            foreach (var dependencyName in dependencyNames)
            {
                if (dependencyName != null)
                {
                    ReadFigure(dependencyName, alreadyDeserializedFigures, nameBlacklist, nodeMap, drawing, callbackWhenCreated);
                }
            }

            List<IFigure> dependencies = new List<IFigure>();
            for (int i = 0; i < dependencyNames.Length; i++)
            {
                string dependencyName = dependencyNames[i];
                IFigure existingDependency = null;
                if (dependencyName == null || !alreadyDeserializedFigures.TryGetValue(dependencyName, out existingDependency))
                {
                    ReportError(string.Format("{0} is left out: it is built on {1}, which could not be read.", figureName, dependencyName ?? "a figure without a name"));
                    return;
                }

                // a part of the figure (a vertex a regular polygon works out), not the figure
                string partName = dependencyNodes[i].ReadString("Part");
                if (partName != null)
                {
                    existingDependency = (existingDependency as IFigureParts)?.GetPart(partName);
                    if (existingDependency == null)
                    {
                        ReportError(string.Format("{0} is left out: {1} has no part {2}.", figureName, dependencyName, partName));
                        return;
                    }
                }

                dependencies.Add(existingDependency);
            }

            IFigure instance = InstantiateFigure(type, drawing, dependencies);
            if (instance == null)
            {
                ReportError(string.Format("{0} is left out: a {1} can't be read from a file.", figureName, type.Name));
                return;
            }

            if (!GenerateNewNames)
            {
                instance.Name = figureName;
            }
            if (drawing.Figures[instance.Name] != null)
            {
                instance.GenerateNewNameIfNecessary(drawing, nameBlacklist);
            }

            nameBlacklist.Add(instance.Name);
            alreadyDeserializedFigures.Add(figureName, instance);

            try
            {
                instance.ReadXml(figureNode);
            }
            catch (Exception ex)
            {
                ReportError(string.Format("{0} was not read in full: {1}", figureName, ex.Message));
                callbackWhenCreated(instance);
                return;
            }

            callbackWhenCreated(instance);
        }

        IFigure InstantiateFigure(Type type, Drawing drawing, IList<IFigure> dependencies)
        {
            IFigure instance = null;

            var defaultCtor = type.GetConstructor(Type.EmptyTypes);
            if (defaultCtor != null)
            {
                instance = Activator.CreateInstance(type) as IFigure;
                instance.Drawing = drawing;
                instance.Dependencies = dependencies;
            }
            else
            {
                var ctorWithDrawingAndDependencies = type.GetConstructor(new Type[] { typeof(Drawing), typeof(IList<IFigure>) });
                if (ctorWithDrawingAndDependencies != null)
                {
                    instance = Activator.CreateInstance(type, drawing, dependencies) as IFigure;
                }
            }

            return instance;
        }

        public bool GenerateNewNames { get; set; }
    }
}
