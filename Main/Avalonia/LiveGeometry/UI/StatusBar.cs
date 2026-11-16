using System;
using System.Collections.Generic;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Layout;
using Avalonia.Media;

namespace LiveGeometry;

public class StatusBar : WrapPanel
{
    public TextBlock TextBlock { get; set; }
    public Border border = new Border();

    public StatusBar()
    {
        HorizontalAlignment = HorizontalAlignment.Left;
        VerticalAlignment = VerticalAlignment.Bottom;
        Margin = new Thickness(8);

        border.Padding = new Thickness(10, 5, 10, 5);
        border.CornerRadius = new CornerRadius(6);
        border.BindTheme(Border.BackgroundProperty, nameof(AppTheme.HintBackground));
        border.BindTheme(Border.BorderBrushProperty, nameof(AppTheme.HintBorder));
        border.BorderThickness = new Thickness(1);
        border.BoxShadow = new BoxShadows(new BoxShadow()
        {
            OffsetY = 1,
            Blur = 4,
            Color = Color.FromArgb(40, 0, 0, 0)
        });
        border.PointerPressed += border_MouseLeftButtonDown;

        TextBlock = new TextBlock()
        {
            FontSize = 12,
            TextWrapping = TextWrapping.Wrap,
            MaxWidth = 640
        };
        TextBlock.BindTheme(TextBlock.ForegroundProperty, nameof(AppTheme.Text));

        targetText.BindTheme(TextBlock.ForegroundProperty, nameof(AppTheme.Text));
        DockPanel.SetDock(targetIcon, Dock.Left);
        targetRow.Children.Add(targetIcon);
        targetRow.Children.Add(targetText);

        var lines = new StackPanel();
        lines.Children.Add(targetRow);
        lines.Children.Add(TextBlock);
        border.Child = lines;

        this.Children.Add(border);
    }

    const double IconSize = 16;
    const double IconGap = 6;

    // What a click at the cursor takes, a line over the hint: the figure's icon, its name
    // bold and its construction faded, as a row of the Figure List has them. The plate is
    // at the bottom of the window, so the hint under the line is not moved by it.
    readonly DockPanel targetRow = new DockPanel() { IsVisible = false };

    readonly Decorator targetIcon = new Decorator()
    {
        Width = IconSize,
        Height = IconSize,
        Margin = new Thickness(0, 0, IconGap, 0),
        VerticalAlignment = VerticalAlignment.Top
    };

    readonly TextBlock targetText = new TextBlock()
    {
        FontSize = 12,
        TextWrapping = TextWrapping.Wrap,
        MaxWidth = 640 - IconSize - IconGap,
        VerticalAlignment = VerticalAlignment.Center
    };

    ClickStatus target;
    readonly List<IDisposable> fadedBindings = new();

    /// <summary>What a click at the cursor takes, shown over the <see cref="Text"/>; null for nothing</summary>
    public ClickStatus Target
    {
        get => target;
        set
        {
            if (value == null ? target == null : value.SameAs(target))
            {
                return;
            }

            bool sameIcon = value != null && target != null && value.IconKey == target.IconKey;
            target = value;
            targetRow.IsVisible = value != null;
            foreach (var binding in fadedBindings)
            {
                binding.Dispose();
            }

            fadedBindings.Clear();
            targetText.Inlines ??= new InlineCollection();
            targetText.Inlines.Clear();
            if (value == null)
            {
                targetIcon.Child = null;
                return;
            }

            if (!sameIcon)
            {
                targetIcon.Child = value.IconKey == null ? null : FigureIcons.Create(value.IconKey, IconSize);
            }

            targetIcon.IsVisible = targetIcon.Child != null;
            foreach (var (text, style) in value.Pieces)
            {
                var run = new Run(text);
                if (style == ClickStatusStyle.Name)
                {
                    run.FontWeight = FontWeight.Bold;
                }
                else if (style == ClickStatusStyle.Faded)
                {
                    fadedBindings.Add(run.Bind(TextElement.ForegroundProperty, this.GetResourceObservable(nameof(AppTheme.TextMuted))));
                }

                targetText.Inlines.Add(run);
            }

            PlaceLines();
        }
    }

    // the hint is not there while it says nothing; a gap between the two lines
    void PlaceLines()
    {
        bool hasText = !string.IsNullOrEmpty(TextBlock.Text) || TextBlock.Inlines?.Count > 0;
        TextBlock.IsVisible = hasText;
        targetRow.Margin = new Thickness(0, 0, 0, hasText ? 4 : 0);
    }

    public string Text
    {
        get => TextBlock.Text;
        set
        {
            TextBlock.Inlines?.Clear();
            TextBlock.Text = value;
            PlaceLines();
        }
    }

    /// <summary>
    /// The text with a link in it: the first occurrence of linkText is underlined and opens
    /// the address in the browser. A click on the rest puts the bar away, as ever.
    /// </summary>
    public void Show(string text, string linkText, Uri link)
    {
        int at = text.IndexOf(linkText, StringComparison.Ordinal);
        if (at < 0)
        {
            Text = text;
            return;
        }

        var anchor = new HyperlinkButton() { Text = linkText };
        anchor.BindTheme(TextBlock.ForegroundProperty, nameof(AppTheme.Accent));
        // the press is handled here, otherwise the bar would be hidden by its own
        // click-to-dismiss handler before the release that is the click
        anchor.PointerPressed += (sender, e) => e.Handled = true;
        anchor.Click += (sender, e) => TopLevel.GetTopLevel(this)?.Launcher.LaunchUriAsync(link);

        TextBlock.Text = null;
        TextBlock.Inlines.Clear();
        TextBlock.Inlines.Add(new Run(text.Substring(0, at)));
        // the same font and size as the text around it, so the bottom of the line is the
        // bottom of the anchor (aligned by baseline, the anchor's bottom edge is put on the
        // baseline and the link is drawn a few pixels too high)
        TextBlock.Inlines.Add(new InlineUIContainer(anchor) { BaselineAlignment = BaselineAlignment.Bottom });
        TextBlock.Inlines.Add(new Run(text.Substring(at + linkText.Length)));
        PlaceLines();
    }

    new 
    public bool IsVisible
    {
        get => this.Visibility == Visibility.Visible; set => this.Visibility = value ? Visibility.Visible : Visibility.Collapsed;
    }

    void border_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        IsVisible = false;
    }
}
