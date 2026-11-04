using System;
using System.Collections.Generic;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Layout;
using Avalonia.Media;

namespace LiveGeometry;

/// <summary>
/// Stepping through the construction (<see cref="StepPlayback"/>) as the host shows it: the
/// Step by Step toggle on the ribbon; while stepping, the step bar in the ribbon's place
/// and at its height (First, Previous, Next, Last, "Step 4 of 9: what the step says", and
/// the toggle to leave), the Figure List shown, the keys.
/// </summary>
public partial class DrawingHost
{
    public Command CommandToggleSteps { get; set; }
    public Command CommandStepFirst { get; set; }
    public Command CommandStepPrevious { get; set; }
    public Command CommandStepNext { get; set; }
    public Command CommandStepLast { get; set; }

    /// <summary>The bar that takes the ribbon's place while stepping</summary>
    public Border StepBar { get; private set; }

    readonly TextBlock stepText = new TextBlock()
    {
        VerticalAlignment = VerticalAlignment.Center,
        TextWrapping = TextWrapping.Wrap,
        FontSize = 16,
        Margin = new Thickness(14, 0, 14, 0)
    };

    readonly Run stepNumberRun = new Run() { FontWeight = FontWeight.SemiBold };
    readonly Run stepTextRun = new Run();

    /// <summary>Whether the Figure List was up before the stepping showed it</summary>
    bool figureExplorerShownBeforeSteps;

    /// <summary>Whether the bar is up in the ribbon's place</summary>
    public bool IsStepping { get; private set; }

    /// <summary>The stepping started or stopped: the ribbon goes or comes back (the application shows and hides it)</summary>
    public event Action StepsStateChanged;

    public StepPlayback Playback => CurrentDrawing?.Playback;

    void CreateStepCommands()
    {
        CommandToggleSteps = new Command(ToggleSteps, ToggleIcons.StepByStep(), "Step by Step", BehaviorCategories.Selection)
        {
            IsChecked = () => Playback?.IsOn == true,
            HintText = "Step through the construction, a figure or a group of figures at a time. "
                + "Space, the arrow keys and Page Up and Down go from step to step, Escape leaves. "
                + "A figure's page says which step it is in and what the step says."
        };
        CommandStepFirst = new Command(() => Playback?.First(), ToggleIcons.StepFirst(), "First", BehaviorCategories.Selection)
        {
            Shortcut = "Home"
        };
        CommandStepPrevious = new Command(() => Playback?.Previous(), ToggleIcons.StepPrevious(), "Previous", BehaviorCategories.Selection)
        {
            Shortcut = "Left"
        };
        CommandStepNext = new Command(() => Playback?.Next(), ToggleIcons.StepNext(), "Next", BehaviorCategories.Selection)
        {
            Shortcut = "Space"
        };
        CommandStepLast = new Command(() => Playback?.Last(), ToggleIcons.StepLast(), "Last", BehaviorCategories.Selection)
        {
            Shortcut = "End"
        };
    }

    /// <summary>
    /// The bar: the toggle to leave, the step buttons, the step's number and text filling
    /// the rest. The same surface as the ribbon's tools strip, with its bottom line, at
    /// its own height (lower than the ribbon's two rows).
    /// </summary>
    Border CreateStepBar()
    {
        // the buttons at the tools' own height, in the middle of the bar's: the two toggles
        // and the four step buttons, each group a unit, so that in a window too narrow for
        // all of them in a row the step buttons go onto a second row together. The step
        // buttons are set closer together than the tools, so that a phone (390 wide) takes
        // all six in one row.
        var buttons = new WrapPanel()
        {
            Orientation = Orientation.Horizontal,
            VerticalAlignment = VerticalAlignment.Center
        };
        var navigation = new StackPanel() { Orientation = Orientation.Horizontal };
        foreach (var command in new[] { CommandStepFirst, CommandStepPrevious, CommandStepNext, CommandStepLast })
        {
            navigation.Children.Add(new CommandToolButton(command) { DrawingHost = this, SideRoom = StepButtonSideRoom });
        }

        // the bar's own copy of the toggle (an icon can be in one button only)
        var leaveCommand = new Command(ToggleSteps, ToggleIcons.StepByStep(), CommandToggleSteps.Name, BehaviorCategories.Selection)
        {
            IsChecked = CommandToggleSteps.IsChecked,
            HintText = "Back to the whole drawing and all the tools (Escape)."
        };
        var leave = new CommandToolButton(leaveCommand) { DrawingHost = this };

        // and of the Figure List toggle: the ribbon's is away with the ribbon
        var figureListCommand = new Command(ToggleFigureExplorer, ToggleIcons.FigureList(), "Figure List", BehaviorCategories.Selection)
        {
            IsChecked = () => FigureExplorer.IsVisible,
            HintText = "The figures with their step numbers at the left; a row goes to its step."
        };
        var figureList = new CommandToolButton(figureListCommand) { DrawingHost = this };
        stepText.BindTheme(TextBlock.ForegroundProperty, nameof(AppTheme.Text));
        stepText.Inlines = new InlineCollection() { stepNumberRun, stepTextRun };

        // the toggles first, as on the ribbon, a divider, the step buttons, the text
        var divider = new Avalonia.Controls.Shapes.Rectangle()
        {
            Width = 1,
            Height = 44,
            Margin = new Thickness(6, 0, 6, 0),
            VerticalAlignment = VerticalAlignment.Center
        };
        divider.BindTheme(Avalonia.Controls.Shapes.Shape.FillProperty, nameof(AppTheme.Separator));
        var toggles = new StackPanel() { Orientation = Orientation.Horizontal };
        toggles.Children.Add(leave);
        toggles.Children.Add(figureList);
        toggles.Children.Add(divider);
        buttons.Children.Add(toggles);
        buttons.Children.Add(navigation);

        var layout = new DockPanel();
        DockPanel.SetDock(buttons, Dock.Left);
        layout.Children.Add(buttons);
        layout.Children.Add(stepText);

        StepBar = new Border()
        {
            BorderThickness = new Thickness(0, 0, 0, 1),
            Padding = new Thickness(6, 3, 6, 3),
            IsVisible = false,
            Child = layout
        };
        StepBar.BindTheme(Border.BackgroundProperty, nameof(AppTheme.Background));
        StepBar.BindTheme(Border.BorderBrushProperty, nameof(AppTheme.Separator));

        // In a narrow window (a phone) the text beside the buttons would be a word a line:
        // the buttons go on a row of their own and the text under them, as wide as the
        // window. Decided by the bar's width alone, so the layout settles in one pass.
        StepBar.SizeChanged += (s, e) =>
        {
            double room = e.NewSize.Width - StepBar.Padding.Left - StepBar.Padding.Right - buttons.DesiredSize.Width;
            bool beside = room >= MinimumStepTextWidth;
            if ((DockPanel.GetDock(buttons) == Dock.Left) != beside)
            {
                DockPanel.SetDock(buttons, beside ? Dock.Left : Dock.Top);
                stepText.Margin = beside ? new Thickness(14, 0, 14, 0) : new Thickness(8, 2, 8, 6);
            }
        };
        return StepBar;
    }

    /// <summary>The least width the step's text gets beside the buttons before it goes under them</summary>
    const double MinimumStepTextWidth = 420;

    /// <summary>The room at the sides of a step button's icon and text (a tool's is 6)</summary>
    const double StepButtonSideRoom = 2;

    /// <summary>The window width from which the stepping opens the Figure List by itself (the ribbon's own small-screen limit)</summary>
    const double FigureExplorerWhileSteppingMinimumWidth = 700;

    /// <summary>The Step by Step toggle after the Figure List on the ribbon. Called once the tools are on it.</summary>
    public void AddStepControls()
    {
        AddToolbarButton(CommandToggleSteps);
    }

    public void ToggleSteps()
    {
        if (Playback?.IsOn == true)
        {
            ExitSteps();
        }
        else
        {
            EnterSteps();
        }
    }

    /// <summary>Starts at the first step with the Drag tool, nothing in the side panel and the Figure List up</summary>
    public void EnterSteps()
    {
        var drawing = CurrentDrawing;
        if (drawing == null)
        {
            return;
        }

        CloseSidePanel();
        drawing.SetDefaultBehavior();
        if (!drawing.Playback.Start())
        {
            ShowHint("Nothing to step through: every figure is hidden or not a step.");
            CommandToolButton.UpdateToggles();
            return;
        }

        DrawingControl.Focus();
    }

    /// <summary>Back to the whole drawing and every tool (the stepping's Changed puts the ribbon back)</summary>
    public void ExitSteps()
    {
        Playback?.Stop();
    }

    /// <summary>The bar follows the stepping: on, off, or at another step</summary>
    void Playback_Changed()
    {
        var playback = Playback;
        bool on = playback?.IsOn == true;
        if (on != IsStepping)
        {
            ShowStepsState(on);
        }

        if (!on)
        {
            return;
        }

        stepNumberRun.Text = "Step " + (playback.Current + 1) + " of " + playback.Count + ": ";
        stepTextRun.Text = playback.CurrentStep?.Text ?? "";
        CommandStepFirst.Enabled = playback.Current > 0;
        CommandStepPrevious.Enabled = playback.Current > 0;
        CommandStepNext.Enabled = playback.Current < playback.Count - 1;
        CommandStepLast.Enabled = playback.Current < playback.Count - 1;
    }

    /// <summary>
    /// Stepping: the bar in the ribbon's place (the application hides the ribbon on
    /// <see cref="StepsStateChanged"/>), the Figure List up. Not stepping: the ribbon
    /// back, and the Figure List as it was before.
    /// </summary>
    void ShowStepsState(bool on)
    {
        IsStepping = on;
        if (on)
        {
            // the Figure List up where there is room for it beside the drawing, away on a
            // phone, where it would take half the width (the bar's toggle brings it back
            // either way); as it was before, on leaving
            figureExplorerShownBeforeSteps = FigureExplorer.IsVisible;
            bool room = Bounds.Width >= FigureExplorerWhileSteppingMinimumWidth;
            if (FigureExplorer.IsVisible != room)
            {
                ShowFigureExplorer(room);
            }
        }
        else if (FigureExplorer.IsVisible != figureExplorerShownBeforeSteps)
        {
            ShowFigureExplorer(figureExplorerShownBeforeSteps);
        }

        StepBar.IsVisible = on;
        StepsStateChanged?.Invoke();
        CommandToolButton.UpdateToggles();
    }

    /// <summary>The keys of the stepping, from the canvas; false for any other key</summary>
    public bool HandleStepKey(Avalonia.Input.Key key)
    {
        var playback = Playback;
        if (playback?.IsOn != true)
        {
            return false;
        }

        switch (key)
        {
            case Avalonia.Input.Key.Right:
            case Avalonia.Input.Key.Down:
            case Avalonia.Input.Key.Space:
            case Avalonia.Input.Key.PageDown:
                playback.Next();
                return true;
            case Avalonia.Input.Key.Left:
            case Avalonia.Input.Key.Up:
            case Avalonia.Input.Key.PageUp:
                playback.Previous();
                return true;
            case Avalonia.Input.Key.Home:
                playback.First();
                return true;
            case Avalonia.Input.Key.End:
                playback.Last();
                return true;
        }

        return false;
    }

    void Playback_CollectionChanged(object sender, EventArgs e)
    {
        Playback?.Refresh();
    }
}
