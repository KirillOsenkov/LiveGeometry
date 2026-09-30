using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using Avalonia.Media.Immutable;
using Avalonia.Styling;
using Avalonia.Threading;

namespace DynamicGeometry;

/// <summary>
/// The colors of the app's chrome (toolbar, ribbon, side panel, Figure List, gallery page):
/// one named color per role, and a theme is a set of values for them. Every theme is a
/// <see cref="ThemeVariant"/> with a resource dictionary of brushes under the role names,
/// registered on the application (<see cref="Register"/>); the chrome binds its brushes to
/// those resources (<see cref="ThemeBinding"/>), so switching the application's requested
/// variant re-skins the chrome and Fluent's own controls (text boxes, scroll bars, menus)
/// in one go, and a color set on a theme while it shows changes the screen at once - the
/// theme can be edited in the property grid, and <see cref="CopyAsCode"/> takes the result
/// back into this file. A new theme is a new instance here with a variant of its own,
/// inheriting Light or Dark so that Fluent has something to fall back on.
/// (Not "Theme": every control has a Theme property, its ControlTheme, which would shadow
/// the class inside the controls that use it most.)
/// </summary>
[PropertyGridNoUndo]
public class AppTheme : INotifyPropertyChanged
{
    /// <summary>The choice that follows the operating system (or the browser) instead of naming a theme</summary>
    public const string SystemChoice = "System";

    public static AppTheme Light { get; } = new AppTheme("Light", ThemeVariant.Light)
    {
        Strip = Color.Parse("#DDE0E3"),
        HeaderRow = Color.Parse("#E9ECF1"),
        Background = Color.Parse("#F6F7F9"),
        GroupBackground = Color.Parse("#FCFCFD"),
        InputBackground = Color.Parse("#FFFFFF"),
        Page = Color.Parse("#FFFFFF"),
        TabLine = Color.Parse("#A9B1BE"),
        Separator = Color.Parse("#D5D9E0"),
        ButtonHover = Color.Parse("#E6EBF2"),
        ButtonPressed = Color.Parse("#D3DCE8"),
        ButtonChecked = Color.Parse("#D2E7FF"),
        ButtonCheckedBorder = Color.Parse("#6FAEEC"),
        Text = Color.Parse("#2B3038"),
        TextEmphasis = Color.Parse("#000000"),
        TextMuted = Color.Parse("#6B7482"),
        TextFaint = Color.Parse("#B4BAC4"),
        IconOutline = Color.Parse("#3A4250"),
        ShapeOutline = Color.Parse("#A87B05"),
        ShapeIconFill = new LinearGradientBrush()
        {
            StartPoint = new RelativePoint(0, 0, RelativeUnit.Relative),
            EndPoint = new RelativePoint(1, 1, RelativeUnit.Relative),
            GradientStops =
            {
                new GradientStop(Color.Parse("#FFFEF0"), 0),
                new GradientStop(Color.Parse("#F1E7A8"), 1)
            }
        },
        ImageFill = new LinearGradientBrush()
        {
            StartPoint = new RelativePoint(0, 0, RelativeUnit.Relative),
            EndPoint = new RelativePoint(1, 1, RelativeUnit.Relative),
            GradientStops =
            {
                new GradientStop(Color.Parse("#E0FEFF"), 0),
                new GradientStop(Color.Parse("#4EE7FF"), 1)
            }
        },
        RulerFill = Color.Parse("#FFFF00"),
        AngleFill = Color.Parse("#33FF33"),
        AreaFill = Color.Parse("#FFD6D6"),
        AreaHatch = Color.Parse("#C8606E"),
        Guide = Color.Parse("#8A94A6"),
        Accent = Color.Parse("#2F7BD6"),
        Destructive = Color.Parse("#B3261E"),
        HintBackground = Color.Parse("#FFFDE8"),
        HintBorder = Color.Parse("#D9D29A"),
        ErrorBackground = Color.Parse("#FDECEC"),
        ErrorBorder = Color.Parse("#D93B3B"),
        ErrorText = Color.Parse("#9B1C1C"),
        Paper = Color.Parse("#FFFFFF"),
        Ink = Color.Parse("#000000"),
        Line = Color.Parse("#64000000"),
        SliderTrack = Color.Parse("#C0C0C0"),
        FreePointFill = Color.Parse("#FFFF64"),
        PointOnFigureFill = Color.Parse("#7CE38B"),
        IntersectionPointFill = Color.Parse("#6FD3F7"),
        MidpointFill = Color.Parse("#FFB45A"),
        DependentPointFill = Color.Parse("#D0D0D0"),
        ShapeFill = Color.Parse("#64FFFFC8"),
        Axis = Color.Parse("#8080FF"),
        GridMajor = Color.Parse("#D3D3D3"),
        GridMinor = Color.Parse("#ECECEC")
    };

    public static AppTheme Dark { get; } = new AppTheme("Dark", ThemeVariant.Dark)
    {
        Strip = Color.Parse("#1B1F26"),
        HeaderRow = Color.Parse("#22272F"),
        Background = Color.Parse("#2A303A"),
        GroupBackground = Color.Parse("#323945"),
        InputBackground = Color.Parse("#1F242B"),
        Page = Color.Parse("#1E2228"),
        TabLine = Color.Parse("#4D5665"),
        Separator = Color.Parse("#3A414D"),
        ButtonHover = Color.Parse("#3A4250"),
        ButtonPressed = Color.Parse("#48525F"),
        ButtonChecked = Color.Parse("#2A4A6E"),
        ButtonCheckedBorder = Color.Parse("#4C8DD6"),
        Text = Color.Parse("#D9DEE6"),
        TextEmphasis = Color.Parse("#FFFFFF"),
        TextMuted = Color.Parse("#98A2B3"),
        TextFaint = Color.Parse("#6B7482"),
        IconOutline = Color.Parse("#D0D6DF"),
        ShapeOutline = Color.Parse("#D9B44A"),
        ShapeIconFill = new LinearGradientBrush()
        {
            StartPoint = new RelativePoint(0, 0, RelativeUnit.Relative),
            EndPoint = new RelativePoint(1, 1, RelativeUnit.Relative),
            GradientStops =
            {
                new GradientStop(Color.Parse("#7D7A60"), 0),
                new GradientStop(Color.Parse("#4A4838"), 1)
            }
        },
        ImageFill = new SolidColorBrush(Color.Parse("#4E9A5E")),
        RulerFill = Color.Parse("#A88E32"),
        AngleFill = Color.Parse("#3F8F4E"),
        AreaFill = Color.Parse("#6E4A50"),
        AreaHatch = Color.Parse("#C88A94"),
        Guide = Color.Parse("#7A8595"),
        Accent = Color.Parse("#5AA0F2"),
        Destructive = Color.Parse("#F0736A"),
        HintBackground = Color.Parse("#3B3A2C"),
        HintBorder = Color.Parse("#6E6A45"),
        ErrorBackground = Color.Parse("#4A2A2A"),
        ErrorBorder = Color.Parse("#D95050"),
        ErrorText = Color.Parse("#F2A0A0"),
        Paper = Color.Parse("#2B2B2B"),
        Ink = Color.Parse("#D0D0D0"),
        Line = Color.Parse("#D3D3D3"),
        SliderTrack = Color.Parse("#606060"),
        FreePointFill = Color.Parse("#F5C542"),
        PointOnFigureFill = Color.Parse("#6BCF7F"),
        IntersectionPointFill = Color.Parse("#4FC3F7"),
        MidpointFill = Color.Parse("#F0A050"),
        DependentPointFill = Color.Parse("#8E949C"),
        ShapeFill = Color.Parse("#46D8CC96"),
        Axis = Color.Parse("#8C8CFF"),
        GridMajor = Color.Parse("#4A4A4A"),
        GridMinor = Color.Parse("#383838")
    };

    /// <summary>Every theme there is, in the order a list offers them</summary>
    public static IReadOnlyList<AppTheme> All { get; } = new[] { Light, Dark };

    /// <summary>
    /// The theme whose values a style and a paper hold as their own; every other theme's are
    /// overrides on top of them
    /// </summary>
    public static AppTheme Base => Light;

    public static bool IsBase(AppTheme theme)
    {
        return theme == Base;
    }

    public static readonly CornerRadius ButtonCornerRadius = new CornerRadius(5);
    public const double ButtonMinWidth = 52;

    public AppTheme(string name, ThemeVariant variant)
    {
        Name = name;
        Variant = variant;
    }

    public string Name { get; }

    public ThemeVariant Variant { get; }

    /// <summary>The brushes of this theme under the names of the colors, kept current by the setters</summary>
    public ResourceDictionary Resources { get; } = new ResourceDictionary();

    public event PropertyChangedEventHandler PropertyChanged;

    /// <summary>Puts every theme's dictionary under its variant, once, before the first window</summary>
    public static void Register(Application application)
    {
        foreach (var theme in All)
        {
            application.Resources.ThemeDictionaries[theme.Variant] = theme.Resources;
        }

        registered = true;
    }

    /// <summary>The theme on screen</summary>
    public static AppTheme Current
    {
        get
        {
            var variant = Application.Current?.ActualThemeVariant;
            return All.FirstOrDefault(theme => theme.Variant == variant) ?? Light;
        }
    }

    /// <summary>What the operating system (or the browser) asks for: Light or Dark</summary>
    public static AppTheme System
    {
        get
        {
            var settings = Application.Current?.PlatformSettings;
            return settings != null && settings.GetColorValues().ThemeVariant == Avalonia.Platform.PlatformThemeVariant.Dark
                ? Dark
                : Light;
        }
    }

    /// <summary>Raised after the theme on screen changed, for whatever can't be bound to a resource</summary>
    public static event Action CurrentChanged;

    /// <summary>
    /// Raised after a color of any theme was set (tweaked in the property grid), for what
    /// holds copies of theme colors (the default styles of a drawing). Only for a color
    /// drawings take (<see cref="IsDrawingColor"/>), and once per dispatcher tick however
    /// many sets a drag in the color picker made (<see cref="SetResource"/>).
    /// </summary>
    public static event Action ColorsChanged;

    /// <summary>
    /// Counts what changes the look of drawings: a switch of theme, a tweaked color. A
    /// drawing that is off screen skips the refresh and compares against this when it
    /// comes back (<see cref="Drawing.RefreshThemeIfStale"/>).
    /// </summary>
    public static int Version { get; private set; }

    static bool listening;
    static bool registered;

    /// <summary>
    /// Shows the named theme, or with <see cref="SystemChoice"/> whichever the system asks for
    /// (and follows the system from then on)
    /// </summary>
    public static void Apply(string choice)
    {
        var application = Application.Current;
        if (!listening)
        {
            listening = true;
            application.ActualThemeVariantChanged += (s, e) =>
            {
                Version++;
                CurrentChanged?.Invoke();
            };
        }

        var theme = ByName(choice);
        application.RequestedThemeVariant = theme?.Variant ?? ThemeVariant.Default;
    }

    public static AppTheme ByName(string name)
    {
        return All.FirstOrDefault(theme => theme.Name == name);
    }

    public override string ToString()
    {
        return Name + " theme";
    }

    #region Colors

    // The chrome's surfaces, from the top of the window down: the toolbar strip, the row of
    // ribbon headers, the tools and the side panel, a boxed group inside the panel.

    Color strip;
    [PropertyGridVisible]
    [PropertyGridGroup("Surfaces")]
    public Color Strip { get => strip; set => Set(ref strip, value); }

    Color headerRow;
    [PropertyGridVisible]
    [PropertyGridGroup("Surfaces")]
    public Color HeaderRow { get => headerRow; set => Set(ref headerRow, value); }

    Color background;
    [PropertyGridVisible]
    [PropertyGridGroup("Surfaces")]
    public Color Background { get => background; set => Set(ref background, value); }

    Color groupBackground;
    [PropertyGridVisible]
    [PropertyGridGroup("Surfaces")]
    public Color GroupBackground { get => groupBackground; set => Set(ref groupBackground, value); }

    Color inputBackground;
    [PropertyGridVisible]
    [PropertyGridGroup("Surfaces")]
    public Color InputBackground { get => inputBackground; set => Set(ref inputBackground, value); }

    /// <summary>The gallery page</summary>
    Color page;
    [PropertyGridVisible]
    [PropertyGridGroup("Surfaces")]
    public Color Page { get => page; set => Set(ref page, value); }

    /// <summary>The line along the ribbon's header row, the outline of a tab, the border of a box</summary>
    Color tabLine;
    [PropertyGridVisible]
    [PropertyGridGroup("Lines")]
    public Color TabLine { get => tabLine; set => Set(ref tabLine, value); }

    Color separator;
    [PropertyGridVisible]
    [PropertyGridGroup("Lines")]
    public Color Separator { get => separator; set => Set(ref separator, value); }

    Color buttonHover;
    [PropertyGridVisible]
    [PropertyGridGroup("Buttons")]
    public Color ButtonHover { get => buttonHover; set => Set(ref buttonHover, value); }

    Color buttonPressed;
    [PropertyGridVisible]
    [PropertyGridGroup("Buttons")]
    public Color ButtonPressed { get => buttonPressed; set => Set(ref buttonPressed, value); }

    Color buttonChecked;
    [PropertyGridVisible]
    [PropertyGridGroup("Buttons")]
    public Color ButtonChecked { get => buttonChecked; set => Set(ref buttonChecked, value); }

    Color buttonCheckedBorder;
    [PropertyGridVisible]
    [PropertyGridGroup("Buttons")]
    public Color ButtonCheckedBorder { get => buttonCheckedBorder; set => Set(ref buttonCheckedBorder, value); }

    Color text;
    [PropertyGridVisible]
    [PropertyGridGroup("Text")]
    public Color Text { get => text; set => Set(ref text, value); }

    /// <summary>The selected tab header</summary>
    Color textEmphasis;
    [PropertyGridVisible]
    [PropertyGridGroup("Text")]
    public Color TextEmphasis { get => textEmphasis; set => Set(ref textEmphasis, value); }

    /// <summary>Text that is beside the point: the name of the app on the gallery page, a credit</summary>
    Color textMuted;
    [PropertyGridVisible]
    [PropertyGridGroup("Text")]
    public Color TextMuted { get => textMuted; set => Set(ref textMuted, value); }

    /// <summary>Barely there until hovered: the cross that closes the side panel</summary>
    Color textFaint;
    [PropertyGridVisible]
    [PropertyGridGroup("Text")]
    public Color TextFaint { get => textFaint; set => Set(ref textFaint, value); }

    /// <summary>The outlines of the drawn chrome icons (toolbar, property grid buttons)</summary>
    Color iconOutline;
    [PropertyGridVisible]
    [PropertyGridGroup("Icons")]
    public Color IconOutline { get => iconOutline; set => Set(ref iconOutline, value); }

    /// <summary>The outline of every shape in the tool icons of shapes (triangle, square, polygon)...</summary>
    Color shapeOutline;
    [PropertyGridVisible]
    [PropertyGridGroup("Icons")]
    public Color ShapeOutline { get => shapeOutline; set => Set(ref shapeOutline, value); }

    /// <summary>
    /// ...and what fills it (a brush: it may be a gradient). The icon's own: a new polygon on
    /// the paper is filled with <see cref="ShapeFill"/>, which is translucent and would sink
    /// into the ribbon. Also the figure that is transformed, in the icons of the
    /// transformations: it is a shape like any other...
    /// </summary>
    Brush shapeIconFill;
    [PropertyGridVisible]
    [PropertyGridGroup("Icons")]
    public Brush ShapeIconFill { get => shapeIconFill; set => Set(ref shapeIconFill, value); }

    /// <summary>...and its image, which is what stands out</summary>
    Brush imageFill;
    [PropertyGridVisible]
    [PropertyGridGroup("Icons")]
    public Brush ImageFill { get => imageFill; set => Set(ref imageFill, value); }

    /// <summary>The ruler of the Distance tool (and so the Measure tab)</summary>
    Color rulerFill;
    [PropertyGridVisible]
    [PropertyGridGroup("Icons")]
    public Color RulerFill { get => rulerFill; set => Set(ref rulerFill, value); }

    /// <summary>The protractor of the Angle tool</summary>
    Color angleFill;
    [PropertyGridVisible]
    [PropertyGridGroup("Icons")]
    public Color AngleFill { get => angleFill; set => Set(ref angleFill, value); }

    /// <summary>The hatched pentagon of the Area tool...</summary>
    Color areaFill;
    [PropertyGridVisible]
    [PropertyGridGroup("Icons")]
    public Color AreaFill { get => areaFill; set => Set(ref areaFill, value); }

    /// <summary>...and its hatching</summary>
    Color areaHatch;
    [PropertyGridVisible]
    [PropertyGridGroup("Icons")]
    public Color AreaHatch { get => areaHatch; set => Set(ref areaHatch, value); }

    /// <summary>Faint construction lines in an icon (a grid, an axis)</summary>
    Color guide;
    [PropertyGridVisible]
    [PropertyGridGroup("Icons")]
    public Color Guide { get => guide; set => Set(ref guide, value); }

    /// <summary>The blue of arrows and of what is pointed out</summary>
    Color accent;
    [PropertyGridVisible]
    [PropertyGridGroup("Icons")]
    public Color Accent { get => accent; set => Set(ref accent, value); }

    Color destructive;
    [PropertyGridVisible]
    [PropertyGridGroup("Icons")]
    public Color Destructive { get => destructive; set => Set(ref destructive, value); }

    Color hintBackground;
    [PropertyGridVisible]
    [PropertyGridGroup("Notes")]
    public Color HintBackground { get => hintBackground; set => Set(ref hintBackground, value); }

    Color hintBorder;
    [PropertyGridVisible]
    [PropertyGridGroup("Notes")]
    public Color HintBorder { get => hintBorder; set => Set(ref hintBorder, value); }

    Color errorBackground;
    [PropertyGridVisible]
    [PropertyGridGroup("Notes")]
    public Color ErrorBackground { get => errorBackground; set => Set(ref errorBackground, value); }

    Color errorBorder;
    [PropertyGridVisible]
    [PropertyGridGroup("Notes")]
    public Color ErrorBorder { get => errorBorder; set => Set(ref errorBorder, value); }

    Color errorText;
    [PropertyGridVisible]
    [PropertyGridGroup("Notes")]
    public Color ErrorText { get => errorText; set => Set(ref errorText, value); }

    // The paper and what a new drawing draws on it: the default styles are built from these
    // (StyleManager.AddDefaultStyles, CartesianGrid), and so are the tool icons, which show
    // the figures as they would look.

    Color paper;
    [PropertyGridVisible]
    [PropertyGridGroup("Paper")]
    public Color Paper { get => paper; set => Set(ref paper, value); }

    /// <summary>Lines, text and the rims of points, on the paper and in the tool icons</summary>
    Color ink;
    [PropertyGridVisible]
    [PropertyGridGroup("Paper")]
    public Color Ink { get => ink; set => Set(ref ink, value); }

    /// <summary>The line of a new drawing (the Line style): segments, lines, circles</summary>
    Color line;
    [PropertyGridVisible]
    [PropertyGridGroup("Paper")]
    public Color Line { get => line; set => Set(ref line, value); }

    /// <summary>The bar a slider's knob runs along (the SliderTrack style), and in the Slider tool's icon</summary>
    Color sliderTrack;
    [PropertyGridVisible]
    [PropertyGridGroup("Paper")]
    public Color SliderTrack { get => sliderTrack; set => Set(ref sliderTrack, value); }

    Color freePointFill;
    [PropertyGridVisible]
    [PropertyGridGroup("Paper")]
    public Color FreePointFill { get => freePointFill; set => Set(ref freePointFill, value); }

    Color pointOnFigureFill;
    [PropertyGridVisible]
    [PropertyGridGroup("Paper")]
    public Color PointOnFigureFill { get => pointOnFigureFill; set => Set(ref pointOnFigureFill, value); }

    Color intersectionPointFill;
    [PropertyGridVisible]
    [PropertyGridGroup("Paper")]
    public Color IntersectionPointFill { get => intersectionPointFill; set => Set(ref intersectionPointFill, value); }

    Color midpointFill;
    [PropertyGridVisible]
    [PropertyGridGroup("Paper")]
    public Color MidpointFill { get => midpointFill; set => Set(ref midpointFill, value); }

    Color dependentPointFill;
    [PropertyGridVisible]
    [PropertyGridGroup("Paper")]
    public Color DependentPointFill { get => dependentPointFill; set => Set(ref dependentPointFill, value); }

    /// <summary>The translucent fill of a new polygon or circle</summary>
    Color shapeFill;
    [PropertyGridVisible]
    [PropertyGridGroup("Paper")]
    public Color ShapeFill { get => shapeFill; set => Set(ref shapeFill, value); }

    Color axis;
    [PropertyGridVisible]
    [PropertyGridGroup("Paper")]
    public Color Axis { get => axis; set => Set(ref axis, value); }

    Color gridMajor;
    [PropertyGridVisible]
    [PropertyGridGroup("Paper")]
    public Color GridMajor { get => gridMajor; set => Set(ref gridMajor, value); }

    Color gridMinor;
    [PropertyGridVisible]
    [PropertyGridGroup("Paper")]
    public Color GridMinor { get => gridMinor; set => Set(ref gridMinor, value); }

    #endregion

    void Set(ref Color field, Color value, [CallerMemberName] string key = null)
    {
        field = value;
        SetResource(key, new ImmutableSolidColorBrush(value));
    }

    /// <summary>A fill that may be a gradient: the resource is the brush itself, which every icon bound to the key shares</summary>
    void Set(ref Brush field, Brush value, [CallerMemberName] string key = null)
    {
        field = value;
        SetResource(key, value.ToImmutable());
    }

    readonly Dictionary<string, object> pendingResources = new Dictionary<string, object>();

    /// <summary>
    /// Puts the brush under the key. Once the dictionary is registered, a write makes the
    /// application tell the whole tree that its resources changed, and every binding
    /// anywhere (the chrome's, and the ones inside Fluent's templates) reads its resource
    /// again, whatever the key; so the writes of a drag in the color picker - a set per
    /// pointer move - are gathered and made once the dispatcher is idle, followed by one
    /// <see cref="ColorsChanged"/>. The property itself changed already: the grid's row
    /// shows the value, <see cref="CopyAsCode"/> reads it.
    /// </summary>
    void SetResource(string key, object brush)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(key));
        if (!registered)
        {
            Resources[key] = brush;
            return;
        }

        bool scheduled = pendingResources.Count > 0;
        pendingResources[key] = brush;
        if (!scheduled)
        {
            Dispatcher.UIThread.Post(FlushResources, DispatcherPriority.Background);
        }
    }

    void FlushResources()
    {
        bool drawingsChanged = false;
        foreach (var pair in pendingResources)
        {
            Resources[pair.Key] = pair.Value;
            drawingsChanged |= IsDrawingColor(pair.Key);
        }

        pendingResources.Clear();
        if (drawingsChanged)
        {
            Version++;
            ColorsChanged?.Invoke();
        }
    }

    static HashSet<string> drawingColors;

    /// <summary>
    /// Whether drawings take the color: the default styles are built from the Paper group
    /// (<c>StyleManager.AddDefaultStyles</c>, <c>CartesianGrid</c>) and the gallery's text
    /// style from <see cref="Text"/>. A change of any other color repaints the chrome alone.
    /// </summary>
    public static bool IsDrawingColor(string key)
    {
        if (drawingColors == null)
        {
            drawingColors = new HashSet<string>(
                ColorProperties
                    .Where(property => property.GetCustomAttributes<PropertyGridGroupAttribute>().Any(group => group.Name == "Paper"))
                    .Select(property => property.Name))
            {
                nameof(Text)
            };
        }

        return drawingColors.Contains(key);
    }

    /// <summary>The color with another alpha: a rim or a line that lets the paper through</summary>
    public static Color WithAlpha(Color color, byte alpha)
    {
        return Color.FromArgb(alpha, color.R, color.G, color.B);
    }

    /// <summary>The color properties (and the fills that are brushes), in the order declared</summary>
    public static IEnumerable<PropertyInfo> ColorProperties
    {
        get
        {
            return typeof(AppTheme)
                .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(property => property.PropertyType == typeof(Color) || property.PropertyType == typeof(Brush));
        }
    }

    /// <summary>
    /// The colors as the initializer above, to paste back after tweaking them in the property
    /// grid (also printed to the console)
    /// </summary>
    [PropertyGridVisible]
    [PropertyGridName("Copy as code")]
    [PropertyGridIcon(PropertyGridIcon.Copy)]
    public void CopyAsCode()
    {
        var code = ToCode();
        Clipboard.SetText(code);
        Console.WriteLine(code);
    }

    public string ToCode()
    {
        var sb = new StringBuilder();
        sb.AppendLine($"public static AppTheme {Name} {{ get; }} = new AppTheme(\"{Name}\", ThemeVariant.{Variant})");
        sb.AppendLine("{");
        var properties = ColorProperties.ToArray();
        for (int i = 0; i < properties.Length; i++)
        {
            string value = ToCode(properties[i].GetValue(this));
            string separator = i < properties.Length - 1 ? "," : "";
            sb.AppendLine($"    {properties[i].Name} = {value}{separator}");
        }

        sb.AppendLine("};");
        return sb.ToString();
    }

    /// <summary>A color or a brush as the expression that makes it, indented as a line of the initializer</summary>
    static string ToCode(object value)
    {
        if (value is Color color)
        {
            return ToCode(color);
        }

        if (value is ISolidColorBrush solid)
        {
            return $"new SolidColorBrush({ToCode(solid.Color)})";
        }

        if (value is not ILinearGradientBrush gradient)
        {
            return "null";
        }

        var sb = new StringBuilder();
        sb.AppendLine("new LinearGradientBrush()");
        sb.AppendLine("    {");
        sb.AppendLine($"        StartPoint = {ToCode(gradient.StartPoint)},");
        sb.AppendLine($"        EndPoint = {ToCode(gradient.EndPoint)},");
        sb.AppendLine("        GradientStops =");
        sb.AppendLine("        {");
        for (int i = 0; i < gradient.GradientStops.Count; i++)
        {
            var stop = gradient.GradientStops[i];
            string separator = i < gradient.GradientStops.Count - 1 ? "," : "";
            sb.AppendLine($"            new GradientStop({ToCode(stop.Color)}, {ToCode(stop.Offset)}){separator}");
        }

        sb.AppendLine("        }");
        sb.Append("    }");
        return sb.ToString();
    }

    static string ToCode(Color color)
    {
        return $"Color.Parse(\"{ColorText.ToHex(color)}\")";
    }

    static string ToCode(RelativePoint point)
    {
        return $"new RelativePoint({ToCode(point.Point.X)}, {ToCode(point.Point.Y)}, RelativeUnit.{point.Unit})";
    }

    static string ToCode(double number)
    {
        // adding zero turns a negative zero, which would print as -0, into zero
        // (global: System is the theme the system asks for, in here)
        return (global::System.Math.Round(number, digits: 4) + 0.0).ToString("0.####", CultureInfo.InvariantCulture);
    }
}
