using System.ComponentModel;
using System.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;
using GuiLabs.Undo;

namespace DynamicGeometry;

/// <summary>
/// Picked by <c>[PropertyGridPreferredEditor("Emoji")]</c> on a string (a point style's
/// character): a search box, the emoji it finds as swatches, and the chosen one as big as
/// the points will show it, with its name. A single character pasted into the box is offered
/// as it is. Rows of its own, without the label column: the tab already says what it is.
/// </summary>
public class EmojiEditorFactory : BaseValueEditorFactory<EmojiEditor, string>
{
    public EmojiEditorFactory()
    {
        // after the plain string editor: only the attribute chooses this one
        LoadOrder = 1;
    }
}

public class EmojiEditor : StackPanel, IValueEditor
{
    const int MaxResults = 240;
    const double SwatchSize = 30;
    const double PreviewMaxSize = 96;

    readonly TextBox searchBox;
    readonly ListBox results;
    readonly EmojiGlyph preview;
    readonly TextBlock previewName;
    bool guard;

    public EmojiEditor()
    {
        searchBox = new TextBox()
        {
            Watermark = "Search, or paste a character",
            MaxWidth = 340,
            HorizontalAlignment = HorizontalAlignment.Left,
            Width = 340
        };
        searchBox.TextChanged += (s, e) => ShowResults();
        searchBox.KeyDown += SearchBox_KeyDown;

        results = new ListBox()
        {
            MaxHeight = 210,
            Margin = new Thickness(0, 6, 0, 6)
        };
        results.SelectionChanged += Results_SelectionChanged;

        preview = new EmojiGlyph()
        {
            HorizontalAlignment = HorizontalAlignment.Left,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(2, 4, 12, 4)
        };
        previewName = new TextBlock()
        {
            VerticalAlignment = VerticalAlignment.Center,
            TextWrapping = TextWrapping.Wrap,
            MaxWidth = 220
        };
        var previewRow = new StackPanel() { Orientation = Orientation.Horizontal };
        previewRow.Children.Add(preview);
        previewRow.Children.Add(previewName);

        Children.Add(searchBox);
        Children.Add(results);
        Children.Add(previewRow);
        Children.Add(new TextBlock()
        {
            Text = "Emoji art: Twemoji by Twitter, CC-BY 4.0",
            FontSize = 10,
            Foreground = RibbonTheme.TabHeaderText,
            Margin = new Thickness(0, 6, 0, 6)
        });

        ShowResults();
    }

    IValueProvider value;
    public IValueProvider Value
    {
        get
        {
            return value;
        }
        set
        {
            if (this.value != null)
            {
                this.value.ValueChanged -= ValueChanged;
            }

            if (OwnerStyle != null)
            {
                OwnerStyle.PropertyChanged -= OwnerStyle_PropertyChanged;
            }

            this.value = value;
            if (this.value != null)
            {
                this.value.ValueChanged += ValueChanged;
            }

            if (OwnerStyle != null)
            {
                OwnerStyle.PropertyChanged += OwnerStyle_PropertyChanged;
            }

            ValueChanged();
        }
    }

    public ActionManager ActionManager { get; set; }

    /// <summary>What the character belongs to: the preview takes its size and color</summary>
    PointStyle OwnerStyle => value?.Parent as PointStyle;

    string Character => value?.GetValue<string>();

    void OwnerStyle_PropertyChanged(object sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName is "Size" or "Fill")
        {
            UpdatePreview();
        }
    }

    void ValueChanged()
    {
        UpdatePreview();
        SelectCurrent();
    }

    void UpdatePreview()
    {
        string character = Character;
        preview.Text = character;
        preview.Foreground = OwnerStyle?.Fill;
        double size = System.Math.Min(OwnerStyle?.Size ?? PointStyle.DefaultCharacterSize, PreviewMaxSize);
        preview.Width = size;
        preview.Height = size;
        preview.IsVisible = character != null;
        if (character == null)
        {
            previewName.Text = "No emoji yet: pick one above";
            previewName.Foreground = RibbonTheme.TabHeaderText;
            return;
        }

        var emoji = EmojiList.Find(character);
        previewName.Text = (emoji?.Name ?? "") + "\n" + EmojiList.CodePoints(character);
        previewName.Foreground = RibbonTheme.Text;
    }

    void ShowResults()
    {
        string query = searchBox.Text?.Trim();
        var found = string.IsNullOrEmpty(query)
            ? EmojiList.Suggestions
            : EmojiList.Search(query);
        var texts = found.Take(MaxResults).Select(e => (e.Text, e.Name)).ToList();
        // a character no font of ours has would be a box in the browser, whatever the desktop finds
        if (EmojiList.IsSingleCharacter(query) && !texts.Any(t => t.Text == query) && EmojiFont.CanDraw(query))
        {
            texts.Insert(0, (query, EmojiList.Find(query)?.Name ?? EmojiList.CodePoints(query)));
        }

        guard = true;
        results.Items.Clear();
        foreach (var (text, name) in texts)
        {
            var swatch = new EmojiGlyph()
            {
                Text = text,
                Width = SwatchSize,
                Height = SwatchSize,
                Inset = 4,
                Tag = text
            };
            ToolTip.SetTip(swatch, name);
            results.Items.Add(swatch);
        }

        results.IsVisible = texts.Count > 0;
        guard = false;
        SelectCurrent();
    }

    void SelectCurrent()
    {
        guard = true;
        results.SelectedItem = results.Items
            .OfType<EmojiGlyph>()
            .FirstOrDefault(g => (string)g.Tag == Character);
        guard = false;
    }

    void Results_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (!guard && results.SelectedItem is EmojiGlyph glyph)
        {
            Pick((string)glyph.Tag);
        }
    }

    // Enter takes the first thing found (not the first suggestion, with nothing typed)
    void SearchBox_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter
            && !string.IsNullOrWhiteSpace(searchBox.Text)
            && results.Items.OfType<EmojiGlyph>().FirstOrDefault() is EmojiGlyph first)
        {
            Pick((string)first.Tag);
            e.Handled = true;
        }
    }

    void Pick(string character)
    {
        if (value == null || !value.CanSetValue || character == Character)
        {
            return;
        }

        if (ActionManager != null)
        {
            Actions.SetProperty(ActionManager, value, character);
        }
        else
        {
            value.SetValue(character);
        }
    }
}
