using System.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Presenters;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.VisualTree;

namespace LiveGeometry;

/// <summary>
/// The text of a label edited where it is on the canvas (a double click on it, the Drag
/// tool): a box in the label's own font, size and width over the label, which hides while
/// the box is up. Enter or a click elsewhere puts the text in, one undo step; Shift+Enter
/// starts a new line; Escape leaves the label as it was. The side panel is put away
/// meanwhile, since what it would say is right there, and comes back with the label after.
/// </summary>
public class InPlaceTextEditor
{
    readonly Label label;
    readonly Drawing drawing;
    readonly TextBox box;
    readonly Border frame;
    readonly double shapeOpacity;
    bool isClosed;

    // the box's own text sits this far inside the frame: its outline and a gap
    const double FrameThickness = 1.5;
    const double FrameGap = 1.5;

    InPlaceTextEditor(Label label)
    {
        this.label = label;
        drawing = label.Drawing;
        var text = label.TextBlock;
        box = new TextBox()
        {
            Text = label.Text,
            AcceptsReturn = true,
            TextWrapping = text.TextWrapping,
            FontFamily = text.FontFamily,
            FontSize = text.FontSize,
            FontWeight = text.FontWeight,
            FontStyle = text.FontStyle,
            LineHeight = text.LineHeight,
            Foreground = text.Foreground,
            Padding = text.Padding,
            MinHeight = 0,
            MinWidth = 24,
            BorderThickness = new Thickness(0),
            CornerRadius = new CornerRadius(2)
        };
        // (a width includes the padding, the label's as the box's: the lines break where they did)
        if (label.WrapWidth > 0)
        {
            box.Width = label.WrapWidth;
        }

        // the box looks the same whatever the pointer and the keyboard do: Fluent's states
        // would paint it white and draw a line under it
        IBrush background = drawing.Background;
        foreach (var state in new[] { "", "PointerOver", "Focused" })
        {
            box.Resources["TextControlBackground" + state] = background;
            box.Resources["TextControlForeground" + state] = text.Foreground;
            box.Resources["TextControlBorderBrush" + state] = Brushes.Transparent;
        }

        box.Resources["TextControlBorderThemeThicknessFocused"] = new Thickness(0);
        box.AddHandler(InputElement.KeyDownEvent, Box_KeyDown, RoutingStrategies.Tunnel);
        box.LostFocus += (s, e) => Commit();

        frame = new Border()
        {
            BorderThickness = new Thickness(FrameThickness),
            Padding = new Thickness(FrameGap),
            CornerRadius = new CornerRadius(4),
            Child = box,
            ZIndex = label.Shape.ZIndex + 1
        };
        frame.BindTheme(Border.BorderBrushProperty, nameof(AppTheme.Accent));

        shapeOpacity = label.Shape.Opacity;
        label.Shape.Opacity = 0;
        label.Shape.PropertyChanged += Shape_PropertyChanged;
        drawing.Canvas.Children.Add(frame);
        Place();
    }

    /// <summary>The editor open now, if any: one at a time</summary>
    public static InPlaceTextEditor Current { get; private set; }

    /// <summary>Whether the label's text can be edited where it is: a text, not a measurement, and one that can be changed</summary>
    public static bool CanEdit(IFigure figure)
    {
        return figure is Label label
            && !label.Locked
            && label.Drawing?.Canvas != null
            && !label.Drawing.FixedLabels.Contains(label)
            && !label.Drawing.IsRecordingTransaction;
    }

    /// <param name="caret">Where the caret goes, on the canvas (the double click); null for the end of the text</param>
    public static void Open(Label label, Point? caret = null)
    {
        Current?.Commit();
        if (!CanEdit(label))
        {
            return;
        }

        // the panel would say the same text again, over the drawing (put away once the
        // editor is up: the view and a caption stay where the panel left them)
        var editor = new InPlaceTextEditor(label);
        Current = editor;
        label.Drawing.RaiseDisplayProperties(null);
        editor.box.CaretIndex = editor.box.Text?.Length ?? 0;
        editor.box.Focus();
        if (caret != null)
        {
            editor.PutCaretAt(caret.Value);
        }
    }

    /// <summary>The caret where the text is under this place of the canvas, as a click in the box would put it</summary>
    void PutCaretAt(Point onCanvas)
    {
        frame.UpdateLayout();
        var presenter = box.GetVisualDescendants().OfType<TextPresenter>().FirstOrDefault();
        var point = presenter != null ? drawing.Canvas.TranslatePoint(onCanvas, presenter) : null;
        if (point == null)
        {
            return;
        }

        presenter.MoveCaretToPoint(point.Value);
        box.CaretIndex = presenter.CaretIndex;
    }

    // the box follows the label when the view moves under it (a pinned one stays)
    void Shape_PropertyChanged(object sender, AvaloniaPropertyChangedEventArgs e)
    {
        if (e.Property == Canvas.LeftProperty || e.Property == Canvas.TopProperty)
        {
            Place();
        }
    }

    void Place()
    {
        double inset = FrameThickness + FrameGap;
        Canvas.SetLeft(frame, Canvas.GetLeft(label.Shape) - inset);
        Canvas.SetTop(frame, Canvas.GetTop(label.Shape) - inset);
    }

    void Box_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter && e.KeyModifiers == KeyModifiers.None)
        {
            e.Handled = true;
            Commit();
        }
    }

    /// <summary>The text into the label, an undo step if it changed</summary>
    public void Commit()
    {
        if (isClosed)
        {
            return;
        }

        var text = box.Text ?? "";
        Close();
        if (text != label.Text && drawing.Figures.Contains(label))
        {
            Actions.SetProperty(drawing.ActionManager, new PropertyValue(nameof(Label.Text), label), text);
        }

        drawing.RaiseSelectionChanged(drawing.GetSelectedFigures());
    }

    /// <summary>The label as it was (Escape)</summary>
    public void Cancel()
    {
        if (isClosed)
        {
            return;
        }

        Close();
        drawing.RaiseSelectionChanged(drawing.GetSelectedFigures());
    }

    void Close()
    {
        isClosed = true;
        if (Current == this)
        {
            Current = null;
        }

        label.Shape.PropertyChanged -= Shape_PropertyChanged;
        label.Shape.Opacity = shapeOpacity;
        bool hadKeyboard = box.IsKeyboardFocusWithin;
        drawing.Canvas?.Children.Remove(frame);

        // the keys go back to the canvas (the box leaving the tree gives them to no one);
        // not when they went elsewhere already, which is what closed the box
        if (hadKeyboard)
        {
            drawing.Canvas?.Focus();
        }
    }
}
