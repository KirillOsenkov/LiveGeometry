using System;
using System.Collections.Generic;
using System.Linq;
using Avalonia.Media;
using GuiLabs.Undo;

namespace DynamicGeometry
{
    public partial class StyleManager
    {
        static IEnumerable<Type> StyleTypes = Reflector.DiscoverTypes<IFigureStyle>();

        protected readonly ListWithEvents<IFigureStyle> list = new ListWithEvents<IFigureStyle>();

        public Drawing Drawing { get; set; }

        public StyleManager(Drawing drawing)
        {
            list.ItemAdded += OnStyleAdded;
            AddDefaultStyles();
            Drawing = drawing;
        }

        private void OnStyleAdded(IFigureStyle style)
        {
            if (style.Name.IsEmpty())
            {
                style.Name = GenerateUniqueName();
            }
            style.StyleManager = this;
        }

        protected virtual string GenerateUniqueName()
        {
            int n = 1;
            while (this[n.ToString()] != null)
            {
                n++;
            }
            return n.ToString();
        }

        public bool NameIsValid(string name)
        {
            return !list.Any(f => f.Name == name);
        }

        public IFigureStyle this[string index]
        {
            get
            {
                index = CanonicalName(index);
                if (aliases.TryGetValue(index, out var target))
                {
                    index = target;
                }

                foreach (var style in list)
                {
                    if (style.Name == index)
                    {
                        return style;
                    }
                }

                return null;
            }
        }

        /// <summary>The name a style goes by now, for the names older files use</summary>
        public static string CanonicalName(string name)
        {
            return name == LegacyDependentPointStyleName ? DependentPointStyleName : name;
        }

        public IEnumerable<IFigureStyle> GetAllStyles()
        {
            return list;
        }

        public virtual IEnumerable<TStyle> GetStyles<TStyle>()
            where TStyle : class, IFigureStyle
        {
            foreach (var style in list)
            {
                TStyle correctStyle = style as TStyle;
                if (correctStyle != null)
                {
                    yield return correctStyle;
                }
            }
        }

        public IEnumerable<IFigureStyle> GetCompatibleStyles(Type styleType)
        {
            foreach (var style in list)
            {
                Type type = style.GetType();
                if (styleType == type)
                {
                    yield return style;
                }
            }
        }

        protected int numDefaultStyles;

        /// <summary>
        /// The styles of a new drawing. Their colors are the theme's (the paper group of
        /// <see cref="AppTheme"/>): the base value from the Light theme, an override from every
        /// other, so the same style draws right on either paper; the chosen colors (a red line,
        /// a blue outline) read on both and are literal, with a Dark override where they don't.
        /// </summary>
        public virtual void AddDefaultStyles()
        {
            // The look tells how a point behaves: the two kinds that can be dragged are full
            // size and warm/bright, the constructed ones are a little smaller and cooler.
            var freePointStyle = ThemedPoint(FreePointStyleName, size: 10, theme => theme.FreePointFill);
            var pointOnFigureStyle = ThemedPoint(PointOnFigureStyleName, size: 10, theme => theme.PointOnFigureFill);
            var intersectionPointStyle = ThemedPoint(IntersectionPointStyleName, size: 8, theme => theme.IntersectionPointFill);
            var midpointStyle = ThemedPoint(MidpointStyleName, size: 8, theme => theme.MidpointFill);
            var dependentPointStyle = ThemedPoint(DependentPointStyleName, size: 8, theme => theme.DependentPointFill);

            var lineStyle = new LineStyle() { Name = LineStyleName };
            lineStyle.BindToTheme(nameof(LineStyle.Color), theme => theme.Line);
            var lineStyle2 = new LineStyle()
            {
                Name = "OtherLine",
                Color = Color.FromArgb(200, 0, 0, 255)
            };
            lineStyle2.SetOverride(AppTheme.Dark.Name, nameof(LineStyle.Color), Color.FromArgb(200, 122, 155, 255));
            var thickLineStyle = new LineStyle()
            {
                Name = "ThickLine",
                StrokeWidth = 2.5
            };
            thickLineStyle.BindToTheme(nameof(LineStyle.Color), theme => AppTheme.WithAlpha(theme.Ink, 230));
            var redLineStyle = new LineStyle()
            {
                Name = "RedLine",
                Color = Color.FromArgb(255, 216, 59, 59),
                StrokeWidth = 1.5
            };
            var greenLineStyle = new LineStyle()
            {
                Name = "GreenLine",
                Color = Color.FromArgb(255, 46, 158, 79),
                StrokeWidth = 1.5
            };

            // auxiliary constructions: there, but stepping back
            var dashedLineStyle = new LineStyle()
            {
                Name = "DashedLine",
                Color = Color.FromArgb(255, 110, 110, 110),
                StrokeWidth = 1.25,
                Dash = LineDash.Dash
            };
            dashedLineStyle.SetOverride(AppTheme.Dark.Name, nameof(LineStyle.Color), Color.FromArgb(255, 158, 158, 158));
            var dottedLineStyle = new LineStyle()
            {
                Name = "DottedLine",
                Color = Color.FromArgb(255, 110, 110, 110),
                StrokeWidth = 1.5,
                Dash = LineDash.Dot
            };
            dottedLineStyle.SetOverride(AppTheme.Dark.Name, nameof(LineStyle.Color), Color.FromArgb(255, 158, 158, 158));

            // a bar rather than a line: what the knob of a slider runs along
            var sliderTrackStyle = new LineStyle()
            {
                Name = SliderTrackStyleName,
                StrokeWidth = 6
            };
            sliderTrackStyle.BindToTheme(nameof(LineStyle.Color), theme => theme.SliderTrack);

            // an outline with a hint of the same color inside: made for circles, fine for polygons
            var blueOutlineStyle = new ShapeStyle()
            {
                Name = "BlueOutline",
                Color = Color.FromArgb(255, 47, 123, 214),
                StrokeWidth = 1.5,
                Fill = new SolidColorBrush(Color.FromArgb(28, 47, 123, 214))
            };
            var orangeOutlineStyle = new ShapeStyle()
            {
                Name = "OrangeOutline",
                Color = Color.FromArgb(255, 224, 138, 0),
                StrokeWidth = 1.5,
                Fill = new SolidColorBrush(Color.FromArgb(28, 224, 138, 0))
            };
            var purpleOutlineStyle = new ShapeStyle()
            {
                Name = "PurpleOutline",
                Color = Color.FromArgb(255, 136, 84, 208),
                StrokeWidth = 1.5,
                Fill = new SolidColorBrush(Color.FromArgb(28, 136, 84, 208))
            };
            var shapeWithLineStyle = new ShapeStyle() { Name = OutlinedShapeStyleName };
            shapeWithLineStyle.BindToTheme(nameof(ShapeStyle.Fill), theme => new SolidColorBrush(theme.ShapeFill));
            shapeWithLineStyle.BindToTheme(nameof(ShapeStyle.Color), theme => AppTheme.WithAlpha(theme.Ink, 100));
            var shapeStyle = new ShapeStyle()
            {
                Name = ShapeStyleName,
                Color = Colors.Transparent
            };
            shapeStyle.BindToTheme(nameof(ShapeStyle.Fill), theme => new SolidColorBrush(theme.ShapeFill));
            var shapeStyle2 = new ShapeStyle()
            {
                Name = "OtherShape",
                Color = Colors.Transparent,
                Fill = new SolidColorBrush(Color.FromArgb(100, 200, 255, 200))
            };
            shapeStyle2.SetOverride(AppTheme.Dark.Name, nameof(ShapeStyle.Fill), new SolidColorBrush(Color.FromArgb(100, 128, 200, 128)));
            var hyperLinkStyle = ThemedText(HyperlinkStyleName, fontSize: 18);
            var textStyle = ThemedText(TextStyleName, fontSize: 18);
            var headerStyle = ThemedText(HeadingStyleName, fontSize: 40);

            // the caption of a drawing of the gallery: a heading in the splash's blue, the
            // explanation in the chrome's text color, the locus of a "drag to here" ring
            var galleryTitleStyle = new TextStyle()
            {
                Name = GalleryTitleStyleName,
                FontSize = 30,
                Color = Color.FromRgb(0x1F, 0x4E, 0x8C),
                FontFamily = new FontFamily("Segoe UI"),
                Bold = true
            };
            galleryTitleStyle.SetOverride(AppTheme.Dark.Name, nameof(TextStyle.Color), Color.FromRgb(0x9C, 0xC4, 0xF0));
            var galleryTextStyle = new TextStyle()
            {
                Name = GalleryTextStyleName,
                FontSize = 16,
                FontFamily = new FontFamily("Segoe UI")
            };
            galleryTextStyle.BindToTheme(nameof(TextStyle.Color), theme => theme.Text);
            var galleryLocusStyle = new LineStyle()
            {
                Name = GalleryLocusStyleName,
                Color = Color.FromRgb(0xE0, 0x36, 0x2B),
                StrokeWidth = 2.5
            };

            var newStyles = new IFigureStyle[]
            {
                freePointStyle,
                pointOnFigureStyle,
                intersectionPointStyle,
                midpointStyle,
                dependentPointStyle,
                lineStyle,
                lineStyle2,
                thickLineStyle,
                redLineStyle,
                greenLineStyle,
                dashedLineStyle,
                dottedLineStyle,
                sliderTrackStyle,
                shapeStyle,
                shapeStyle2,
                shapeWithLineStyle,
                blueOutlineStyle,
                orangeOutlineStyle,
                purpleOutlineStyle,
                textStyle,
                headerStyle,
                hyperLinkStyle,
                galleryTitleStyle,
                galleryTextStyle,
                galleryLocusStyle
            };

            list.AddRange(newStyles);

            numDefaultStyles = newStyles.Length;
        }

        /// <summary>A point style filled with a theme color, rimmed with the theme's ink</summary>
        static PointStyle ThemedPoint(string name, double size, Func<AppTheme, Color> fill)
        {
            var style = new PointStyle()
            {
                Name = name,
                Size = size
            };
            style.BindToTheme(nameof(PointStyle.Fill), theme => new SolidColorBrush(fill(theme)));
            style.BindToTheme(nameof(PointStyle.Color), theme => AppTheme.WithAlpha(theme.Ink, 100));
            return style;
        }

        /// <summary>A text style in the theme's ink</summary>
        static TextStyle ThemedText(string name, double fontSize)
        {
            var style = new TextStyle()
            {
                Name = name,
                FontSize = fontSize,
                FontFamily = new FontFamily("Segoe UI")
            };
            style.BindToTheme(nameof(TextStyle.Color), theme => theme.Ink);
            return style;
        }

        /// <summary>A theme color was tweaked: the styles built from the theme read it again</summary>
        public void RefreshTheme()
        {
            foreach (var style in list)
            {
                (style as FigureStyle)?.RefreshFromTheme();
            }
        }

        public IFigureStyle GetStyle(string name)
        {
            return this[name];
        }

        /// <summary>Whether the style is what a new drawing has under that name, unchanged: a file needn't carry it</summary>
        public static bool IsUnchangedDefault(IFigureStyle style, IList<IFigureStyle> defaults)
        {
            var original = defaults.FirstOrDefault(candidate => candidate.Name == style.Name);
            return original != null && original.GetType() == style.GetType() && original.GetSignature() == style.GetSignature();
        }

        public void SetStyleIfAvailable(IFigure figure, string styleName)
        {
            var style = GetStyle(styleName);
            if (style != null)
            {
                figure.Style = style;
            }
        }

        public IEnumerable<IFigureStyle> GetSupportedStyles(IFigure figure)
        {
            return GetSupportedStyles(figure.GetType());
        }

        public IEnumerable<IFigureStyle> GetSupportedStyles(Type figureType)
        {
            foreach (var style in list)
            {
                if (style.GetType().SupportsFigureType(figureType))
                {
                    yield return style;
                }
            }
        }

        public static Type GetStyleType(Type figureType)
        {
            foreach (var styleType in StyleTypes)
            {
                if (styleType.SupportsFigureType(figureType))
                {
                    return styleType;
                }
            }
            return null;
        }

        // The names of the default styles. A figure whose style is the default of its kind is
        // written without one, and a file names a default it doesn't carry (the loader adds
        // the defaults it lacks): the names are part of the file format.
        public const string FreePointStyleName = "FreePoint";
        public const string PointOnFigureStyleName = "PointOnFigure";
        public const string IntersectionPointStyleName = "IntersectionPoint";
        public const string MidpointStyleName = "Midpoint";
        public const string DependentPointStyleName = "DependentPoint";
        public const string LineStyleName = "Line";
        public const string SliderTrackStyleName = "SliderTrack";
        public const string ShapeStyleName = "Shape";
        public const string OutlinedShapeStyleName = "OutlinedShape";
        public const string TextStyleName = "Text";
        public const string HeadingStyleName = "Heading";
        public const string HyperlinkStyleName = "Hyperlink";
        public const string GalleryTitleStyleName = "GalleryTitle";
        public const string GalleryTextStyleName = "GalleryText";
        public const string GalleryLocusStyleName = "GalleryLocus";

        /// <summary>What files from before 2026-09-29 call the dependent point style</summary>
        public const string LegacyDependentPointStyleName = "DependentPointStyle";

        /// <summary>
        /// What the defaults looked like in files from before the named defaults (the phone
        /// and CD drawings carry numbered copies): a file style that looks like one of these
        /// is taken for the default it stands for, which follows the theme. The old opaque
        /// black line is not among them - the default line is translucent now, and every
        /// old drawing would turn gray.
        /// </summary>
        static IEnumerable<(IFigureStyle Prototype, string Name)> LegacyDefaults()
        {
            var black = Color.FromRgb(0, 0, 0);
            yield return (LegacyPoint(Color.FromRgb(0xFF, 0xFF, 0x00)), FreePointStyleName);
            yield return (LegacyPoint(Color.FromRgb(0x00, 0xFF, 0x00)), PointOnFigureStyleName);
            yield return (LegacyPoint(Color.FromRgb(0xC0, 0xC0, 0xC0)), DependentPointStyleName);
            yield return (new TextStyle() { FontSize = 18, Color = black, FontFamily = new FontFamily("Segoe UI") }, TextStyleName);
            yield return (new TextStyle() { FontSize = 40, Color = black, FontFamily = new FontFamily("Segoe UI") }, HeadingStyleName);

            static PointStyle LegacyPoint(Color fill)
            {
                return new PointStyle()
                {
                    Size = 10,
                    Fill = new SolidColorBrush(fill),
                    Color = Color.FromRgb(0, 0, 0),
                    StrokeWidth = 1
                };
            }
        }

        /// <summary>The names a file's styles went by that stand for a default now (see <see cref="AddWithDefaults"/>)</summary>
        readonly Dictionary<string, string> aliases = new Dictionary<string, string>();

        public virtual IFigureStyle AssignDefaultStyle(IFigure figure)
        {
            // A drawing from a file gets the named styles it lacks (AddWithDefaults); if its own
            // style of that name isn't a point style, the first point style does, as before.
            // Likewise a slider: its track's style, else the first line style.
            var byKind = figure is IPoint ? GetStyle(GetDefaultPointStyleName(figure))
                : figure is Slider ? GetStyle(SliderTrackStyleName)
                : null;
            if (byKind != null && byKind.GetType().SupportsFigureType(figure.GetType()))
            {
                return byKind;
            }

            var supportedStyles = GetSupportedStyles(figure);
            return supportedStyles.FirstOrDefault();
        }

        static string GetDefaultPointStyleName(IFigure point)
        {
            // PointOnFigure is a FreePoint, so it goes first
            if (point is PointOnFigure)
            {
                return PointOnFigureStyleName;
            }

            if (point is FreePoint)
            {
                return FreePointStyleName;
            }

            if (point is IntersectionPoint)
            {
                return IntersectionPointStyleName;
            }

            if (point is MidPoint)
            {
                return MidpointStyleName;
            }

            // draggable along its circle or line, like a point on a figure
            if (point is TranslatedPoint translated && translated.HasFreedom)
            {
                return PointOnFigureStyleName;
            }

            return DependentPointStyleName;
        }

        public IFigureStyle CreateNewStyle(IFigure figure)
        {
            var newStyle = figure.Style.Clone();
            Actions.AddItem(figure.Drawing.ActionManager, list, newStyle);
            return newStyle;
        }

        public IFigureStyle CreateNewStyle(Drawing drawing, IFigureStyle style)
        {
            var newStyle = style.Clone();
            Actions.AddItem(drawing.ActionManager, list, newStyle);
            return newStyle;
        }

        public IFigureStyle FindExistingOrAddNew(IFigureStyle style)
        {
            var found = list
                .Where(s => s.GetSignature() == style.GetSignature())
                .FirstOrDefault();
            if (found != null)
            {
                return found;
            }

            Actions.AddItem(Drawing.ActionManager, list, style);
            return style;
        }

        public void Clear()
        {
            list.Clear();
        }

        public virtual void Add(IFigureStyle style)
        {
            list.Add(style);
        }

        /// <summary>Takes out a style nothing uses (one a paste brought, on undo); not recorded</summary>
        public void Withdraw(IFigureStyle style)
        {
            list.Remove(style);
        }

        /// <summary>The name, or the name with the first number after it that no style goes by</summary>
        public string FreeName(string name)
        {
            if (this[name] == null)
            {
                return name;
            }

            for (int i = 2; ; i++)
            {
                if (this[name + i] == null)
                {
                    return name + i;
                }
            }
        }

        /// <summary>A fresh set of the styles a new drawing starts with, names included</summary>
        public static List<IFigureStyle> CreateDefaultStyles()
        {
            // unnamed defaults get their names ("1", "2"...) on the way into a manager
            return new StyleManager(drawing: null).list.ToList();
        }

        /// <summary>
        /// A drawing's own styles (from a file, which carries only the ones its figures use)
        /// completed with the default styles it lacks by name, in the order of a new drawing:
        /// the style picker, the style of each kind of new point and the first line or shape
        /// style (which new figures take) are then those of a new drawing, unless the drawing
        /// has its own under the same name. Styles with names of their own come last. A file's
        /// copy of a default that looks the same in Light (files from before the themes carry
        /// every default they use) is dropped for the default itself, which follows the theme;
        /// the figures find it under the name. So is a copy under another name of what a
        /// default used to look like (<see cref="LegacyDefaults"/>): the figures find the
        /// default under the old name.
        /// </summary>
        public void AddWithDefaults(IList<IFigureStyle> own)
        {
            var taken = new HashSet<IFigureStyle>();
            aliases.Clear();
            foreach (var style in own)
            {
                if (style.Name == LegacyDependentPointStyleName)
                {
                    style.Name = DependentPointStyleName;
                }
            }

            foreach (var defaultStyle in CreateDefaultStyles())
            {
                var replacement = own.FirstOrDefault(s => s.Name == defaultStyle.Name);
                if (replacement != null)
                {
                    taken.Add(replacement);
                    if (LooksLikeInLight(replacement, defaultStyle))
                    {
                        replacement = null;
                    }
                }

                list.Add(replacement ?? defaultStyle);
            }

            var legacyDefaults = LegacyDefaults().ToArray();
            foreach (var style in own)
            {
                if (taken.Contains(style))
                {
                    continue;
                }

                var legacy = legacyDefaults.FirstOrDefault(candidate => LooksLikeInLight(style, candidate.Prototype));
                if (legacy.Name != null)
                {
                    aliases[style.Name] = legacy.Name;
                }
                else
                {
                    list.Add(style);
                }
            }
        }

        static bool LooksLikeInLight(IFigureStyle style, IFigureStyle other)
        {
            return style is FigureStyle a
                && other is FigureStyle b
                && a.GetType() == b.GetType()
                && a.Overrides.Count == 0
                && !a.SaysThemes
                && a.GetBaseSignature() == b.GetBaseSignature();
        }

        /// <summary>A style the user made; the ones a drawing starts with stay</summary>
        public bool CanRemove(IFigureStyle style)
        {
            return list.Contains(style) && list.IndexOf(style) >= numDefaultStyles;
        }

        public virtual void Remove(IFigureStyle style)
        {
            if (CanRemove(style))
            {
                var transaction = Transaction.Create(Drawing.ActionManager, false);
                foreach (var fig in Drawing.Figures)
                {
                    if (fig.Style == style)
                    {
                        Actions.SetProperty(Drawing.ActionManager, new PropertyValue("Style",fig), AssignDefaultStyle(fig));
                    }
                }
                Actions.RemoveItem(Drawing.ActionManager, list, style);
                transaction.Commit();
            }
        }
    }
}