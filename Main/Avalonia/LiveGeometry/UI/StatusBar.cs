using System;
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
        border.Child = TextBlock;

        this.Children.Add(border);
    }

    public string Text
    {
        get => TextBlock.Text;
        set
        {
            TextBlock.Inlines?.Clear();
            TextBlock.Text = value;
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
