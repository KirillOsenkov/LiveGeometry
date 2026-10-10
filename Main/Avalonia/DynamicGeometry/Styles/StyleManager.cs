using System;
using System.Collections.Generic;
using System.ComponentModel;
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
            list.ItemRemoved += style => ForgetFreshDefaults();
            AddDefaultStyles();
            Drawing = drawing;
        }

        // The default styles as AddDefaultStyles made them, while the list holds them and
        // nothing else and none of them has changed: a file read into a new drawing takes
        // them as they are (DrawingDeserializer.ReadStyles). It made a second set for every
        // drawing it read, every tile of the gallery among them.
        IFigureStyle[] freshDefaults;

        /// <summary>The defaults of a new drawing, untouched (see <see cref="AddWithDefaults"/>); null once anything was added, taken away or changed</summary>
        public IReadOnlyList<IFigureStyle> FreshDefaults => freshDefaults;

        void ForgetFreshDefaults()
        {
            if (freshDefaults == null)
            {
                return;
            }

            foreach (var style in freshDefaults)
            {
                style.PropertyChanged -= FreshDefault_PropertyChanged;
            }

            freshDefaults = null;
        }

        void FreshDefault_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            ForgetFreshDefaults();
        }

        private void OnStyleAdded(IFigureStyle style)
        {
            ForgetFreshDefaults();
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
            // Each kind of figure has a column per hue in its style picker (StyleHue.All),
            // rows of eight, and its defaults sit in the column of their color: the gray line,
            // the yellow shape. The order of the list is the order of the picker.

            // The look tells how a point behaves: the two kinds that can be dragged are full
            // size and warm/bright, the constructed ones are a little smaller and cooler. The
            // row under them: beads, bigger, for the points that matter.
            var freePointStyle = ThemedPoint(FreePointStyleName, size: 10, theme => theme.FreePointFill);
            var pointOnFigureStyle = ThemedPoint(PointOnFigureStyleName, size: 10, theme => theme.PointOnFigureFill);
            var intersectionPointStyle = ThemedPoint(IntersectionPointStyleName, size: 8, theme => theme.IntersectionPointFill);
            var midpointStyle = ThemedPoint(MidpointStyleName, size: 8, theme => theme.MidpointFill);
            var dependentPointStyle = ThemedPoint(DependentPointStyleName, size: 8, theme => theme.DependentPointFill);
            var pointStyles = new List<IFigureStyle>()
            {
                dependentPointStyle,
                HuePoint("RedPoint", fill: "#FF7B7B", darkFill: "#F07272"),
                midpointStyle,
                freePointStyle,
                pointOnFigureStyle,
                intersectionPointStyle,
                HuePoint("BluePoint", fill: "#7EA8F8", darkFill: "#6E9EF5"),
                HuePoint("PurplePoint", fill: "#B98CF5", darkFill: "#A77EF0")
            };
            pointStyles.AddRange(StyleHue.All.Select(Bead));

            // the handles of a Bezier path: small gray squares, apart from the points (the
            // picker offers it to them alone, IsOffered)
            var handleStyle = ThemedPoint(HandleStyleName, size: 7, theme => theme.DependentPointFill);
            handleStyle.Shape = PointShape.Square;
            pointStyles.Add(handleStyle);

            // Lines: thin, thick, and dashed for auxiliary constructions (there, but stepping
            // back). Gray is the theme's: the line every new figure gets, the ink.
            var lineStyle = new LineStyle() { Name = LineStyleName };
            lineStyle.BindToTheme(nameof(LineStyle.Color), theme => theme.Line);
            var thickLineStyle = new LineStyle()
            {
                Name = "ThickLine",
                StrokeWidth = ThickStrokeWidth
            };
            thickLineStyle.BindToTheme(nameof(LineStyle.Color), theme => AppTheme.WithAlpha(theme.Ink, 230));
            var dashedLineStyle = HueLine("DashedLine", StyleHue.Gray, strokeWidth: 1.25, LineDash.Dash);
            var lineStyles = new List<IFigureStyle>() { lineStyle };
            lineStyles.AddRange(StyleHue.Colors.Select(hue => HueLine(LineStyleNameOf(hue), hue, ThinStrokeWidth, LineDash.Solid)));
            lineStyles.Add(thickLineStyle);
            lineStyles.AddRange(StyleHue.Colors.Select(hue => HueLine(ThickLineStyleNameOf(hue), hue, ThickStrokeWidth, LineDash.Solid)));
            lineStyles.Add(dashedLineStyle);
            lineStyles.AddRange(StyleHue.Colors.Select(hue => HueLine(DashedLineStyleNameOf(hue), hue, ThinStrokeWidth, LineDash.Dash)));

            // a bar rather than a line: what the knob of a slider runs along (the picker
            // offers it to sliders only, IsOffered)
            var sliderTrackStyle = new LineStyle()
            {
                Name = SliderTrackStyleName,
                StrokeWidth = 6
            };
            sliderTrackStyle.BindToTheme(nameof(LineStyle.Color), theme => theme.SliderTrack);
            lineStyles.Add(sliderTrackStyle);

            // Shapes: outlined with a flat fill, outlined with a gradient, and gradients without
            // an outline - a polygon's sides are segments of their own (the shape tools draw
            // them). The fill of a new polygon is among these, yellow and flat, as it always
            // was, and so is the green next to it.
            var shapeStyle = new ShapeStyle()
            {
                Name = ShapeStyleName,
                Color = Colors.Transparent
            };
            shapeStyle.BindToTheme(nameof(ShapeStyle.Fill), theme => new SolidColorBrush(theme.ShapeFill));
            var greenShapeStyle = new ShapeStyle()
            {
                Name = "GreenShape",
                Color = Colors.Transparent,
                Fill = new SolidColorBrush(Color.FromArgb(100, 200, 255, 200))
            };
            greenShapeStyle.SetOverride(AppTheme.Dark.Name, nameof(ShapeStyle.Fill), new SolidColorBrush(Color.FromArgb(100, 128, 200, 128)));
            var shapeStyles = new List<IFigureStyle>();
            shapeStyles.AddRange(StyleHue.All.Select(hue => HueShape(
                OutlineStyleNameOf(hue),
                hue,
                outlined: true,
                gradient: false)));
            shapeStyles.AddRange(StyleHue.All.Select(hue => HueShape(
                GradientOutlineStyleNameOf(hue),
                hue,
                outlined: true,
                gradient: true)));
            shapeStyles.AddRange(StyleHue.All.Select(hue =>
                hue == StyleHue.Brown ? shapeStyle
                : hue == StyleHue.Green ? greenShapeStyle
                : HueShape(ShapeStyleNameOf(hue), hue, outlined: false, gradient: true)));

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

            var newStyles = pointStyles
                .Concat(lineStyles)
                .Concat(shapeStyles)
                .Concat(new IFigureStyle[]
                {
                    textStyle,
                    headerStyle,
                    hyperLinkStyle,
                    galleryTitleStyle,
                    galleryTextStyle,
                    galleryLocusStyle
                })
                .ToArray();

            list.AddRange(newStyles);

            numDefaultStyles = newStyles.Length;
            if (list.Count == newStyles.Length)
            {
                freshDefaults = newStyles;
                foreach (var style in newStyles)
                {
                    style.PropertyChanged += FreshDefault_PropertyChanged;
                }
            }
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

        const double ThinStrokeWidth = 1.5;
        const double ThickStrokeWidth = 2.5;

        /// <summary>
        /// A point style of a color of its own on each paper, rimmed with the theme's ink like
        /// the defaults, and as small as the constructed ones: in the picker's row only the two
        /// kinds that can be dragged, side by side in the middle, are bigger
        /// </summary>
        static PointStyle HuePoint(string name, string fill, string darkFill)
        {
            var style = new PointStyle()
            {
                Name = name,
                Size = 8,
                Fill = new SolidColorBrush(Color.Parse(fill))
            };
            style.SetOverride(AppTheme.Dark.Name, nameof(PointStyle.Fill), new SolidColorBrush(Color.Parse(darkFill)));
            style.BindToTheme(nameof(PointStyle.Color), theme => AppTheme.WithAlpha(theme.Ink, 100));
            return style;
        }

        /// <summary>A big point with a highlight, for the points that matter</summary>
        static PointStyle Bead(StyleHue hue)
        {
            var style = new PointStyle()
            {
                Name = hue.Name + "Bead",
                Size = 12,
                Fill = hue.BeadFill,
                Color = hue.BeadRim
            };
            style.SetOverride(AppTheme.Dark.Name, nameof(PointStyle.Fill), hue.DarkBeadFill);
            style.SetOverride(AppTheme.Dark.Name, nameof(PointStyle.Color), hue.DarkBeadRim);
            return style;
        }

        // The names of a hue's styles, a column of the pickers (StyleHue): the gray ones are
        // the theme's defaults under their own names. One place for the names, so that a
        // style can be traded for another of the same hue (ConvertStyle).

        /// <summary>The thin line of the hue: the default line for gray</summary>
        public static string LineStyleNameOf(StyleHue hue)
        {
            return hue == StyleHue.Gray ? LineStyleName : hue.Name + "Line";
        }

        public static string ThickLineStyleNameOf(StyleHue hue)
        {
            return "Thick" + LineStyleNameOf(hue);
        }

        public static string DashedLineStyleNameOf(StyleHue hue)
        {
            return "Dashed" + LineStyleNameOf(hue);
        }

        /// <summary>The shape outlined in the hue with a flat fill: OutlinedShape for gray</summary>
        public static string OutlineStyleNameOf(StyleHue hue)
        {
            return hue == StyleHue.Gray ? OutlinedShapeStyleName : hue.Name + "Outline";
        }

        public static string GradientOutlineStyleNameOf(StyleHue hue)
        {
            return "Gradient" + hue.Name + "Outline";
        }

        /// <summary>The shape without an outline: the yellow fill of a new polygon in the brown column, the classic green in the green one</summary>
        public static string ShapeStyleNameOf(StyleHue hue)
        {
            return hue == StyleHue.Brown ? ShapeStyleName
                : hue == StyleHue.Green ? "GreenShape"
                : hue.Name + "Shape";
        }

        /// <summary>The hue whose column a default line or shape style of that name is in, or null for any other name</summary>
        public static StyleHue HueOfDefault(string styleName)
        {
            return StyleHue.All.FirstOrDefault(hue =>
                styleName == LineStyleNameOf(hue)
                || styleName == ThickLineStyleNameOf(hue)
                || styleName == DashedLineStyleNameOf(hue)
                || styleName == OutlineStyleNameOf(hue)
                || styleName == GradientOutlineStyleNameOf(hue)
                || styleName == ShapeStyleNameOf(hue));
        }

        static LineStyle HueLine(string name, StyleHue hue, double strokeWidth, LineDash dash)
        {
            var style = new LineStyle()
            {
                Name = name,
                Color = hue.Stroke,
                StrokeWidth = strokeWidth,
                Dash = dash
            };
            style.SetOverride(AppTheme.Dark.Name, nameof(LineStyle.Color), hue.DarkStroke);
            return style;
        }

        /// <summary>
        /// A shape filled with the hue's gradient or with its lighter end alone, outlined in
        /// the hue or not at all
        /// </summary>
        static ShapeStyle HueShape(string name, StyleHue hue, bool outlined, bool gradient)
        {
            var style = new ShapeStyle()
            {
                Name = name,
                Color = outlined ? hue.Stroke : Colors.Transparent,
                StrokeWidth = outlined ? ThinStrokeWidth : 1,
                Fill = gradient ? hue.Fill : hue.SolidFill
            };
            style.SetOverride(AppTheme.Dark.Name, nameof(ShapeStyle.Fill), gradient ? hue.DarkFill : hue.DarkSolidFill);
            if (outlined)
            {
                style.SetOverride(AppTheme.Dark.Name, nameof(ShapeStyle.Color), hue.DarkStroke);
            }

            return style;
        }

        /// <summary>
        /// Whether the style picker offers the style for the figure: not one kept for a single
        /// purpose (a slider's track, the gallery's locus) to any other figure, unless the
        /// figure has it. Every other kind gets full rows of eight (<see cref="StyleHue"/>).
        /// </summary>
        public static bool IsOffered(IFigureStyle style, IFigure figure)
        {
            if (figure == null || figure.Style == style)
            {
                return true;
            }

            return style.Name switch
            {
                SliderTrackStyleName => figure is Slider,
                HandleStyleName => figure is BezierPath.BezierPathHandle,
                GalleryLocusStyleName => false,
                _ => true
            };
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
        public static bool IsUnchangedDefault(IFigureStyle style)
        {
            return NewDrawingDefaults.TryGetValue(style.Name, out var original)
                && original.Style.GetType() == style.GetType()
                && original.Signature == style.GetSignature();
        }

        /// <summary>The default style of that name as a new drawing has it, or null; only to look at</summary>
        public static IFigureStyle GetNewDrawingDefault(string name)
        {
            return NewDrawingDefaults.TryGetValue(name, out var original) ? original.Style : null;
        }

        static Dictionary<string, (IFigureStyle Style, string Signature)> newDrawingDefaults;
        static int newDrawingDefaultsVersion;

        /// <summary>
        /// The default styles of a new drawing by name, with their signatures, which every
        /// save compares the drawing's styles with: made again only when a theme color
        /// changes (<see cref="AppTheme.Version"/>), since the defaults are made from them.
        /// Made for every save, they were most of what a save cost.
        /// </summary>
        static Dictionary<string, (IFigureStyle Style, string Signature)> NewDrawingDefaults
        {
            get
            {
                if (newDrawingDefaults == null || newDrawingDefaultsVersion != AppTheme.Version)
                {
                    newDrawingDefaults = CreateDefaultStyles().ToDictionary(style => style.Name, style => (style, style.GetSignature()));
                    newDrawingDefaultsVersion = AppTheme.Version;
                }

                return newDrawingDefaults;
            }
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
            return GetSupportedStyles(StyledType(figure));
        }

        /// <summary>
        /// The type whose styles the figure takes: its own, but for a Bezier path, whose style
        /// is its inside's (it is a figure a point can be on, which line styles are for)
        /// </summary>
        static Type StyledType(IFigure figure)
        {
            return figure is BezierPath ? typeof(BezierPath.BezierPathInterior) : figure.GetType();
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
        public const string HandleStyleName = "Handle";
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

            return GetDefaultStyle(StyledType(figure));
        }

        /// <summary>
        /// The style a new figure of the type gets, but for points and sliders, which go by
        /// their kind (<see cref="AssignDefaultStyle"/>): the outlined shape for a figure that
        /// draws its own outline around an inside (<see cref="IsOutlinedShape"/>), else the
        /// line, or for a shape the default fill - which is not the first shape style, the
        /// picker shows the outlined ones first - else the first style that fits
        /// </summary>
        public IFigureStyle GetDefaultStyle(Type figureType)
        {
            var supportedStyles = GetSupportedStyles(figureType).ToList();
            if (IsOutlinedShape(figureType))
            {
                var outlined = supportedStyles.FirstOrDefault(style => style.Name == OutlinedShapeStyleName);
                if (outlined != null)
                {
                    return outlined;
                }
            }

            return supportedStyles.FirstOrDefault(style => style.Name == LineStyleName || style.Name == ShapeStyleName)
                ?? supportedStyles.FirstOrDefault();
        }

        /// <summary>
        /// Whether a figure of the type is drawn as an outline around a filled inside: an arc
        /// with an inside, a sector or a circular segment, which the line default left
        /// unfilled (both a line and a shape, it took the line, which comes first in the
        /// list). Not a circle, whose default has always been the line, nor a polygon, whose
        /// outline is its side segments and whose default fill has no stroke.
        /// </summary>
        public static bool IsOutlinedShape(Type figureType)
        {
            return typeof(EllipseArcBase).IsAssignableFrom(figureType)
                && typeof(IShapeWithInterior).IsAssignableFrom(figureType);
        }

        /// <summary>
        /// The style a figure converted into another kind takes (an arc into a sector, a sector
        /// into an arc: <see cref="EllipseArc.Convert"/>). Where the new kind is filled and the
        /// old was not, or the other way round, a default style is traded for the one of the
        /// same hue that fits: a red line becomes the red outline with its flat fill, so that
        /// the sector made of the arc is seen at once, and the outline becomes the line again
        /// (a gradient or an outline-less fill too: the hue is what is kept). Any other style
        /// stays while its kind takes the new figure, else the new figure gets its default.
        /// </summary>
        public IFigureStyle ConvertStyle(IFigure oldFigure, IFigure newFigure)
        {
            var style = oldFigure.Style;
            if (style == null)
            {
                return AssignDefaultStyle(newFigure);
            }

            var newType = StyledType(newFigure);
            if (IsOutlinedShape(oldFigure.GetType()) != IsOutlinedShape(newType))
            {
                var hue = HueOfDefault(style.Name);
                if (hue != null)
                {
                    var counterpart = GetStyle(IsOutlinedShape(newType) ? OutlineStyleNameOf(hue) : LineStyleNameOf(hue));
                    if (counterpart != null && counterpart.GetType().SupportsFigureType(newType))
                    {
                        return counterpart;
                    }
                }
            }

            return style.GetType().SupportsFigureType(newType) ? style : AssignDefaultStyle(newFigure);
        }

        static string GetDefaultPointStyleName(IFigure point)
        {
            if (point is BezierPath.BezierPathHandle)
            {
                return HandleStyleName;
            }

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

        /// <summary>
        /// A copy of the figure's style, added to the drawing's styles, for the figure to take.
        /// A figure that can be filled gets a style with a fill even when it has a line style
        /// now (a circle takes either): the same stroke, and a fill there to switch on.
        /// </summary>
        public IFigureStyle CreateNewStyle(IFigure figure)
        {
            var newStyle = figure.Style is LineStyle line
                && line is not ShapeStyle
                && typeof(ShapeStyle).SupportsFigureType(figure.GetType())
                ? ShapeStyle.WithStrokeOf(line)
                : figure.Style.Clone();
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
            // unnamed defaults get their names ("1", "2"...) on the way into a manager; the
            // manager lets go of them (it would stay subscribed to the styles it hands out)
            var manager = new StyleManager(drawing: null);
            manager.ForgetFreshDefaults();
            return manager.list.ToList();
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
        /// <param name="defaults">Default styles nothing else has (<see cref="FreshDefaults"/>), to take instead of making a new set</param>
        public void AddWithDefaults(IList<IFigureStyle> own, IReadOnlyList<IFigureStyle> defaults = null)
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

            foreach (var defaultStyle in defaults ?? CreateDefaultStyles())
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

            foreach (var style in own)
            {
                if (taken.Contains(style))
                {
                    continue;
                }

                var legacyName = FindLegacyDefault(style);
                if (legacyName != null)
                {
                    aliases[style.Name] = legacyName;
                }
                else
                {
                    list.Add(style);
                }
            }
        }

        // What the legacy defaults look like, worked out once: a signature is every value of
        // a style written out, and a file's styles were compared with each legacy default by
        // working out both signatures again, for every tile of the gallery
        static (Type Type, string Signature, string Name)[] legacySignatures;

        static (Type Type, string Signature, string Name)[] LegacySignatures => legacySignatures ??= LegacyDefaults()
            .Select(legacy => (legacy.Prototype.GetType(), ((FigureStyle)legacy.Prototype).GetBaseSignature(), legacy.Name))
            .ToArray();

        /// <summary>The name of the default the style looks like in Light, as one of <see cref="LegacyDefaults"/>; null when none</summary>
        static string FindLegacyDefault(IFigureStyle style)
        {
            if (!(style is FigureStyle figureStyle) || figureStyle.Overrides.Count != 0 || figureStyle.SaysThemes)
            {
                return null;
            }

            string signature = null;
            foreach (var legacy in LegacySignatures)
            {
                if (legacy.Type != style.GetType())
                {
                    continue;
                }

                signature ??= figureStyle.GetBaseSignature();
                if (signature == legacy.Signature)
                {
                    return legacy.Name;
                }
            }

            return null;
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
                foreach (var fig in Drawing.Figures.SelectMany(FigureParts.WithParts))
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