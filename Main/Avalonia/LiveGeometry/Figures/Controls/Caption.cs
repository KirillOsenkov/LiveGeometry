using System;
using System.Linq;
using GuiLabs.Undo;

namespace LiveGeometry;

/// <summary>
/// A drawing's caption, as the drawings of the gallery have: a heading and, under it or not,
/// an explanation - labels named Title and Description, or the first in the Heading and the
/// Explanation style - pinned to the screen beside or under the figure and wrapped to a
/// column. The layout itself is the app's (<see cref="Drawing.LayOutCaption"/>). A user
/// makes a caption by giving a label the style (<see cref="CaptionStyleValue"/>: the first
/// time, the label is named as the gallery expects, and with a heading there the caption is
/// laid out) or by typing on the drawing's page (<see cref="Drawing.Title"/>,
/// <see cref="Drawing.Description"/>).
/// </summary>
public static class Caption
{
    public const string TitleName = "Title";
    public const string DescriptionName = "Description";
    public const string HintName = "Hint";

    public static Label FindTitle(Drawing drawing)
    {
        return Find(drawing, TitleName, StyleManager.HeadingStyleName);
    }

    public static Label FindDescription(Drawing drawing)
    {
        return Find(drawing, DescriptionName, StyleManager.ExplanationStyleName);
    }

    /// <summary>A caption is a heading, with an explanation under it or not</summary>
    public static bool HasCaption(Drawing drawing)
    {
        return FindTitle(drawing) != null;
    }

    /// <summary>
    /// The label named so, else the first in the style that isn't named as another part of
    /// the caption (2026-10-16: a label given the style and not the name was no caption, and
    /// a user had to do both)
    /// </summary>
    public static Label Find(Drawing drawing, string name, string styleName)
    {
        if (drawing.Figures[name] is Label named)
        {
            return named;
        }

        return drawing.Figures
            .OfType<Label>()
            .FirstOrDefault(label => label.Name is not (TitleName or DescriptionName or HintName)
                && label.Style?.Name == styleName);
    }

    /// <summary>The name the gallery expects of a label in the style: Title for a heading, Description for an explanation; null for any other style</summary>
    public static string NameForStyle(string styleName)
    {
        return styleName switch
        {
            StyleManager.HeadingStyleName => TitleName,
            StyleManager.ExplanationStyleName => DescriptionName,
            _ => null
        };
    }

    /// <summary>
    /// The heading's or the explanation's text, set from the drawing's page: the label is
    /// made if there is none (named and styled as the caption expects, on the screen until
    /// the layout pins it), and the caption is laid out once there is a heading. Recorded
    /// through the action manager, or changed directly without one.
    /// </summary>
    public static void SetText(Drawing drawing, bool title, string text, ActionManager actionManager)
    {
        var label = title ? FindTitle(drawing) : FindDescription(drawing);
        if (label != null)
        {
            if (actionManager == null)
            {
                label.Text = text;
            }
            else
            {
                actionManager.SetProperty(label, nameof(Label.Text), text);
            }

            return;
        }

        label = Factory.CreateLabel(drawing);
        label.Name = title ? TitleName : DescriptionName;
        label.Style = drawing.StyleManager.GetStyle(title ? StyleManager.HeadingStyleName : StyleManager.ExplanationStyleName);
        label.Text = text;
        label.MoveTo(drawing.CoordinateSystem.ViewCenter);
        label.CaptionMade = true;
        if (actionManager == null)
        {
            drawing.Figures.Add(label);
            return;
        }

        Actions.Add(drawing, label);
        RecordLayout(drawing);
    }

    /// <summary>
    /// Records the caption's layout as an undo step, inside the transaction of what made
    /// the caption: the labels pinned, wrapped and on a plate, the figure in the room they
    /// leave, and from then on Zoom to fit lays it out the same way. Undo puts the labels,
    /// the view and Zoom to fit back as they were. Nothing without both labels, and nothing
    /// where nobody lays captions out (a tool, the player).
    /// </summary>
    public static void RecordLayout(Drawing drawing)
    {
        var layOut = drawing.LayOutCaption;
        if (layOut == null || !HasCaption(drawing))
        {
            return;
        }

        var saved = new[] { FindTitle(drawing), FindDescription(drawing), drawing.Figures[HintName] as Label }
            .Where(label => label != null)
            .Select(label => (label, label.Pin, label.PinOffset, label.Coordinates, label.WrapWidth, label.Backdrop))
            .ToArray();
        var coordinateSystem = drawing.CoordinateSystem;
        var center = coordinateSystem.ViewCenter;
        double unitLength = coordinateSystem.UnitLength;
        var fitToWindow = drawing.FitToWindow;
        drawing.ActionManager.RecordAction(new CallMethodAction(
            () =>
            {
                drawing.FitToWindow = layOut;
                layOut();
            },
            () =>
            {
                drawing.FitToWindow = fitToWindow;
                foreach (var (label, pin, offset, coordinates, width, backdrop) in saved)
                {
                    label.WrapWidth = width;
                    label.Backdrop = backdrop;
                    label.Pin = pin;
                    if (pin == LabelPin.None)
                    {
                        label.MoveTo(coordinates);
                    }
                    else
                    {
                        label.PinOffset = offset;
                        if (label.Drawing != null)
                        {
                            label.UpdateVisual();
                        }
                    }
                }

                coordinateSystem.SetView(center, unitLength);
            }));
    }
}

/// <summary>
/// The Style row of a label: the first time the label is given the Heading or the Explanation
/// style, it is named as the caption expects (Title, Description) unless it has a name of its
/// own or the name is taken, and once the drawing has both labels the caption is laid out
/// (<see cref="Caption.RecordLayout"/>), all in the one undo step of the pick. Later changes
/// of the style are plain (<see cref="Label.CaptionMade"/>).
/// </summary>
public class CaptionStyleValue : StylePropertyValueProvider, ISelfRecordingValue
{
    public bool TryRecord(ActionManager actionManager, object value)
    {
        var label = (Label)Parent;
        var drawing = label.Drawing;
        string expected = value is IFigureStyle style ? Caption.NameForStyle(style.Name) : null;
        if (label.CaptionMade || expected == null || drawing == null)
        {
            return false;
        }

        label.CaptionMade = true;
        using (Transaction.Create(actionManager, delayed: false))
        {
            Actions.SetProperty(actionManager, new PropertyValue(nameof(FigureBase.StyleDisplay), label), value);
            if (label.HasDefaultName && drawing.Figures[expected] == null)
            {
                actionManager.SetProperty(label, nameof(FigureBase.Name), expected);
            }

            Caption.RecordLayout(drawing);
        }

        // the page again: its Pinned to, Wrap width and Background rows changed under it
        drawing.RaiseDisplayProperties(label);
        return true;
    }
}

/// <summary>The Title and Description rows of the drawing's page (<see cref="Caption.SetText"/>)</summary>
public class CaptionTextValue : PropertyValue, ISelfRecordingValue
{
    public bool TryRecord(ActionManager actionManager, object value)
    {
        using (Transaction.Create(actionManager, delayed: false))
        {
            Caption.SetText((Drawing)Parent, title: Name == nameof(Drawing.Title), (string)value ?? "", actionManager);
        }

        return true;
    }
}
