using System;
using System.Collections.Generic;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Controls.Primitives;
using Avalonia.Controls.Primitives.PopupPositioning;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Threading;

namespace LiveGeometry;

/// <summary>
/// The completion list of a <see cref="ColoredTextBox"/> (docs/completion-plan.md, after the
/// Visual Studio editor's): what can go at the caret, in a popup under the word, while the
/// keyboard stays in the box. What the list holds is <see cref="ExpressionCompletion"/>'s; this
/// is when it opens and closes, the keys, and the popup. It opens on the first letter of a
/// word, after a dot behind a figure's name, and on Ctrl+Space; it follows the word as it is
/// typed or deleted back, and closes when the caret leaves the word, the text changes any other
/// way, or the box loses the keyboard. Selection is firm (a filled row: Tab, Enter, a click and
/// any character that can't go on a name commit) once the word has a letter, and only
/// outlined while it is empty (after a dot: only Tab and a click commit, and Enter and the
/// rest do what they always do), so that "A." and then "+" types a plus.
/// </summary>
public class CompletionList
{
    const int RowsShown = 8;
    const double RowHeight = 22;
    const double ListFontSize = 13;
    const double MinListWidth = 180;
    const double MaxListWidth = 320;
    const double IconSize = 14;

    readonly ColoredTextBox box;
    readonly Border root;
    readonly StackPanel rowPanel;
    readonly TextBlock heading;
    readonly TextBlock help;
    Popup popup;
    ColoredTextPresenter presenter;

    // the word the open list is for (null: closed), what can go there, and the rows for what
    // is typed of it so far
    CompletionSpot spot;
    IReadOnlyList<CompletionItem> items = [];
    IReadOnlyList<CompletionItem> shown = [];
    int selected;
    bool hard;

    // A key of the user's is changing the text (a character typed, Backspace, Delete): the
    // list follows the word. Any other change of the text - the program, a click putting a
    // reference in, undo, a paste - closes it.
    bool typing;
    bool committing;

    public CompletionList(ColoredTextBox box)
    {
        this.box = box;
        box.AddHandler(InputElement.KeyDownEvent, KeyDown, RoutingStrategies.Tunnel);
        box.AddHandler(InputElement.KeyDownEvent, (s, e) => typing = false, RoutingStrategies.Bubble, handledEventsToo: true);
        box.AddHandler(InputElement.TextInputEvent, BeforeTextInput, RoutingStrategies.Tunnel);
        box.AddHandler(InputElement.TextInputEvent, AfterTextInput, RoutingStrategies.Bubble, handledEventsToo: true);
        box.PropertyChanged += Box_PropertyChanged;
        box.LostFocus += (s, e) => Dispatcher.UIThread.Post(CloseIfUnfocused, DispatcherPriority.Input);

        rowPanel = new StackPanel();
        heading = new TextBlock()
        {
            FontWeight = FontWeight.SemiBold,
            TextWrapping = TextWrapping.Wrap
        };
        help = new TextBlock()
        {
            TextWrapping = TextWrapping.Wrap
        };
        help.BindTheme(TextBlock.ForegroundProperty, nameof(AppTheme.TextMuted));
        var helpStrip = new Border()
        {
            BorderThickness = new Thickness(0, 1, 0, 0),
            Padding = new Thickness(6, 4, 6, 5),
            Margin = new Thickness(0, 2, 0, 0),
            Child = new StackPanel() { Children = { heading, help } }
        };
        helpStrip.BindTheme(Border.BorderBrushProperty, nameof(AppTheme.Separator));
        DockPanel.SetDock(helpStrip, Dock.Bottom);
        root = new Border()
        {
            MinWidth = MinListWidth,
            MaxWidth = MaxListWidth,
            Padding = new Thickness(2),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(4),
            Child = new DockPanel()
            {
                Children =
                {
                    helpStrip,
                    new ScrollViewer()
                    {
                        MaxHeight = RowsShown * RowHeight,
                        HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                        VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                        Content = rowPanel
                    }
                }
            }
        };
        root.BindTheme(Border.BackgroundProperty, nameof(AppTheme.InputBackground));
        root.BindTheme(Border.BorderBrushProperty, nameof(AppTheme.TabLine));
        root.BindTheme(TextElement.ForegroundProperty, nameof(AppTheme.Text));

        // the chrome's size, whatever the box's (a label's text is written larger), which the
        // popup would take after it
        TextElement.SetFontSize(root, ListFontSize);
        TextElement.SetFontWeight(root, FontWeight.Normal);
        TextElement.SetFontStyle(root, FontStyle.Normal);
    }

    public bool IsOpen => spot != null;

    /// <summary>
    /// The popup of the box's template, and the presenter it goes under: drawn in the
    /// window's overlay, so that it never takes the keyboard and is the same on the desktop
    /// and in the browser; below the word, above when there is no room, slid to stay in the
    /// window
    /// </summary>
    public void Attach(Popup popup, ColoredTextPresenter presenter)
    {
        Close();
        if (this.popup != null)
        {
            this.popup.Child = null;
        }

        this.popup = popup;
        this.presenter = presenter;
        if (popup == null)
        {
            return;
        }

        popup.ShouldUseOverlayLayer = true;
        popup.IsLightDismissEnabled = false;
        popup.Placement = PlacementMode.BottomEdgeAlignedLeft;
        popup.PlacementConstraintAdjustment = PopupPositionerConstraintAdjustment.FlipY | PopupPositionerConstraintAdjustment.SlideX;
        popup.PlacementTarget = presenter;
        popup.Child = root;
    }

    #region Keys

    void KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Space && e.KeyModifiers == KeyModifiers.Control)
        {
            Invoke();
            e.Handled = true;
            return;
        }

        if (e.Key == Key.Back || e.Key == Key.Delete)
        {
            typing = true;
            return;
        }

        if (!IsOpen || e.KeyModifiers != KeyModifiers.None)
        {
            return;
        }

        switch (e.Key)
        {
            case Key.Up:
                Move(-1);
                break;
            case Key.Down:
                Move(1);
                break;
            case Key.PageUp:
                Move(-RowsShown);
                break;
            case Key.PageDown:
                Move(RowsShown);
                break;
            case Key.Tab:
                Commit();
                break;
            case Key.Enter:
                // outlined: the box's Enter, as if there were no list (done, a new line, the
                // panel's command); firm: the commit alone, and the next Enter is the box's
                if (!hard)
                {
                    Close();
                    return;
                }

                Commit();
                break;
            case Key.Escape:
                Close();
                break;
            default:
                return;
        }

        e.Handled = true;
    }

    // whatever can't go on a name ends it: the selected name goes in (firm), or the list
    // goes (outlined), and then the character is typed
    void BeforeTextInput(object sender, TextInputEventArgs e)
    {
        typing = true;
        if (IsOpen && !string.IsNullOrEmpty(e.Text) && !Scanner.IsNamePart(e.Text[0]))
        {
            if (hard)
            {
                Commit();
            }
            else
            {
                Close();
            }
        }
    }

    // the character is in: a list may open on it (a commit just now included: "ci" and "."
    // puts in "circle." and opens circle's quantities)
    void AfterTextInput(object sender, TextInputEventArgs e)
    {
        typing = false;
        if (!IsOpen && e.Text?.Length == 1)
        {
            OpenAfterTyping(e.Text[0]);
        }
    }

    /// <summary>The first letter of a word, or a dot after a figure's name</summary>
    void OpenAfterTyping(char typed)
    {
        var found = Locate();
        if (found == null)
        {
            return;
        }

        int caret = box.CaretIndex;
        bool startsWord = Scanner.IsNameStart(typed) && found.Start == caret - 1;
        bool afterDot = typed == '.' && found.MemberOf != null && found.Start == caret;
        if (startsWord || afterDot)
        {
            Open(found, commitIfUnique: false);
        }
    }

    /// <summary>Ctrl+Space: the list for wherever the caret is; a word only one row matches is completed at once</summary>
    void Invoke()
    {
        if (IsOpen)
        {
            Follow();
            return;
        }

        var found = Locate();
        if (found != null)
        {
            Open(found, commitIfUnique: true);
        }
    }

    /// <summary>The first Up or Down on an outlined row only makes it firm, as it is; then the selection moves, stopping at the ends</summary>
    void Move(int delta)
    {
        if (hard || System.Math.Abs(delta) > 1)
        {
            selected = System.Math.Clamp(selected + delta, 0, shown.Count - 1);
        }

        hard = true;
        ShowSelection();
    }

    #endregion

    void Box_PropertyChanged(object sender, AvaloniaPropertyChangedEventArgs e)
    {
        if (!IsOpen || committing)
        {
            return;
        }

        if (e.Property == TextBox.TextProperty && !typing)
        {
            Close();
        }
        else if (e.Property == TextBox.TextProperty || e.Property == TextBox.CaretIndexProperty)
        {
            Follow();
        }
    }

    /// <summary>
    /// The word typed on or deleted back: the rows for it now. The caret gone from the word
    /// (an arrow, a click, Backspace past its start) closes the list. Deleted back to
    /// nothing, the word shows every row, outlined; one more Backspace leaves it.
    /// </summary>
    void Follow()
    {
        var found = Locate();
        if (!spot.IsSameWord(found))
        {
            Close();
            return;
        }

        spot = found;
        Refilter();
    }

    CompletionSpot Locate()
    {
        return box.Drawing == null ? null : ExpressionCompletion.Locate(box.Text, box.CaretIndex, box.Context);
    }

    void Open(CompletionSpot found, bool commitIfUnique)
    {
        items = ExpressionCompletion.Items(found, box.Drawing, box.Owner, box.Context);
        if (items.Count == 0 || popup == null)
        {
            return;
        }

        spot = found;
        if (!Refilter())
        {
            return;
        }

        if (commitIfUnique && spot.Word.Length > 0 && shown.Count == 1)
        {
            Commit();
            return;
        }

        var start = presenter.TextLayout.HitTestTextPosition(spot.Start);
        popup.PlacementRect = new Rect(start.X, start.Y, 1, start.Height);

        // the chrome's font, not the box's (the in-place editor's is the label's)
        var window = TopLevel.GetTopLevel(box);
        if (window != null)
        {
            TextElement.SetFontFamily(root, window.FontFamily);
        }

        popup.IsOpen = true;
    }

    /// <summary>The rows for the word as it is; false, the list closed, when none matches</summary>
    bool Refilter()
    {
        var filter = ExpressionCompletion.Filter(items, spot, box.Drawing, box.Owner);
        if (filter.Shown.Count == 0)
        {
            Close();
            return false;
        }

        shown = filter.Shown;
        selected = filter.Selected;
        hard = filter.Hard;

        // never narrower while open: the list doesn't jump as it is filtered
        root.MinWidth = System.Math.Max(root.MinWidth, root.Bounds.Width);
        rowPanel.Children.Clear();
        for (int i = 0; i < shown.Count; i++)
        {
            rowPanel.Children.Add(CreateRow(shown[i], i));
        }

        ShowSelection();
        return true;
    }

    /// <summary>The selected item's name in place of the word (all of it, also past the caret), the caret after it</summary>
    void Commit()
    {
        if (selected < 0 || selected >= shown.Count)
        {
            Close();
            return;
        }

        var item = shown[selected];
        var word = Locate() ?? spot;
        committing = true;
        try
        {
            if (word.Word != item.InsertText)
            {
                box.SelectionStart = word.Start;
                box.SelectionEnd = word.End;
                box.SelectedText = item.InsertText;
            }

            box.CaretIndex = word.Start + item.InsertText.Length;
        }
        finally
        {
            committing = false;
        }

        Close();
    }

    public void Close()
    {
        spot = null;
        items = [];
        shown = [];
        if (popup != null)
        {
            popup.IsOpen = false;
        }

        root.MinWidth = MinListWidth;
    }

    void CloseIfUnfocused()
    {
        if (!box.IsKeyboardFocusWithin)
        {
            Close();
        }
    }

    #region Rows

    /// <summary>
    /// A row: an icon, the name in the color its kind has in the box with the part typed
    /// bold, and a faint word for what it is. A click commits it, and the keyboard stays in
    /// the box (nothing in the popup can take it).
    /// </summary>
    Control CreateRow(CompletionItem item, int index)
    {
        var name = new TextBlock()
        {
            VerticalAlignment = VerticalAlignment.Center,
            TextTrimming = TextTrimming.CharacterEllipsis,
            Inlines = NameInlines(item.Text, spot.Word)
        };
        name.BindTheme(TextBlock.ForegroundProperty, ColorOf(item.Kind));
        var kind = new TextBlock()
        {
            Text = item.KindText,
            FontSize = ListFontSize - 2,
            Margin = new Thickness(12, 0, 0, 0),
            VerticalAlignment = VerticalAlignment.Center
        };
        kind.BindTheme(TextBlock.ForegroundProperty, nameof(AppTheme.TextMuted));
        var icon = CreateIcon(item);
        icon.Margin = new Thickness(0, 0, 6, 0);
        icon.VerticalAlignment = VerticalAlignment.Center;
        Grid.SetColumn(name, 1);
        Grid.SetColumn(kind, 2);
        var row = new Border()
        {
            Height = RowHeight,
            Padding = new Thickness(4, 0, 6, 0),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(3),
            Background = Brushes.Transparent,
            BorderBrush = Brushes.Transparent,
            Child = new Grid()
            {
                ColumnDefinitions = new ColumnDefinitions("Auto,*,Auto"),
                Children = { icon, name, kind }
            }
        };
        row.PointerPressed += (s, e) =>
        {
            if (e.GetCurrentPoint(row).Properties.IsLeftButtonPressed)
            {
                selected = index;
                hard = true;
                Commit();
                e.Handled = true;
            }
        };
        return row;
    }

    static InlineCollection NameInlines(string text, string word)
    {
        int match = word.Length == 0 ? -1 : text.IndexOf(word, StringComparison.OrdinalIgnoreCase);
        if (match < 0)
        {
            return new InlineCollection() { new Run(text) };
        }

        return new InlineCollection()
        {
            new Run(text.Substring(0, match)),
            new Run(text.Substring(match, word.Length)) { FontWeight = FontWeight.Bold },
            new Run(text.Substring(match + word.Length))
        };
    }

    /// <summary>The figure's ribbon icon; for what has none, a dot in the color its kind has in the box</summary>
    static Control CreateIcon(CompletionItem item)
    {
        if (item.Kind == CompletionItemKind.Figure && item.Figure != null)
        {
            var figureIcon = FigureIcons.Create(FigureIcons.GetKey(item.Figure), IconSize);
            if (figureIcon != null)
            {
                return figureIcon;
            }
        }

        var dot = new Border()
        {
            Width = IconSize,
            Height = IconSize,
            Child = new Border()
            {
                Width = 8,
                Height = 8,
                CornerRadius = new CornerRadius(2),
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            }
        };
        ((Border)dot.Child).BindTheme(Border.BackgroundProperty, ColorOf(item.Kind));
        return dot;
    }

    static string ColorOf(CompletionItemKind kind)
    {
        return kind switch
        {
            CompletionItemKind.Member => nameof(AppTheme.ExpressionProperty),
            CompletionItemKind.Function => nameof(AppTheme.ExpressionFunction),
            CompletionItemKind.Reserved => nameof(AppTheme.ExpressionReserved),
            _ => nameof(AppTheme.ExpressionFigure)
        };
    }

    /// <summary>The selected row filled (firm) or outlined, in view, and its help under the list</summary>
    void ShowSelection()
    {
        for (int i = 0; i < rowPanel.Children.Count; i++)
        {
            var row = (Border)rowPanel.Children[i];
            row.BindTheme(Border.BackgroundProperty, i == selected && hard ? nameof(AppTheme.ButtonChecked) : null, whenNone: Brushes.Transparent);
            row.BindTheme(Border.BorderBrushProperty, i == selected && !hard ? nameof(AppTheme.ButtonCheckedBorder) : null, whenNone: Brushes.Transparent);
        }

        if (selected < 0 || selected >= shown.Count)
        {
            return;
        }

        var item = shown[selected];
        heading.Text = item.Heading;
        help.Text = item.Help;
        help.IsVisible = !string.IsNullOrEmpty(help.Text);

        // once the rows are laid out (a list just filled has none yet)
        var selectedRow = rowPanel.Children[selected];
        Dispatcher.UIThread.Post(() => selectedRow.BringIntoView(), DispatcherPriority.Loaded);
    }

    #endregion
}
