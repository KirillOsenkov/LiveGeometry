using System;
using System.Collections.Generic;
using System.Linq;

namespace LiveGeometry;

/// <summary>
/// Stepping through the construction (<see cref="ConstructionSteps"/>): the figures of the
/// steps after the current one are not drawn, those of the steps before it are dimmed, and
/// the current step's are selected and flashed with what they are built on. Nothing of the
/// drawing changes: a figure's Visible stays as it is, so nothing is recorded for undo and
/// nothing is saved. Hit testing, snapping and dragging leave a figure that isn't drawn
/// alone (<see cref="IsHidden"/>).
/// </summary>
public class StepPlayback
{
    readonly Drawing drawing;
    readonly List<ConstructionStep> steps = new List<ConstructionStep>();
    readonly HashSet<IFigure> hidden = new HashSet<IFigure>();
    readonly HashSet<IFigure> dimmed = new HashSet<IFigure>();
    int current;

    /// <summary>How much of the figures of the steps already taken shows; 1 for no dimming</summary>
    public static double DimOpacity = 0.7;

    public StepPlayback(Drawing drawing)
    {
        this.drawing = drawing;
    }

    public bool IsOn { get; private set; }

    public int Count => steps.Count;

    /// <summary>The current step, from 0; -1 while off</summary>
    public int Current => IsOn ? current : -1;

    public ConstructionStep CurrentStep => IsOn && current < steps.Count ? steps[current] : null;

    /// <summary>The current step changed, or the stepping started or stopped</summary>
    public event Action Changed;

    /// <summary>Starts at the first step. False when the drawing has no step to show.</summary>
    public bool Start()
    {
        Partition();
        if (steps.Count == 0)
        {
            return false;
        }

        // nothing selected while stepping: a halo left from before would read as part of a step
        drawing.Figures.ClearSelection();
        drawing.RaiseSelectionChanged();
        IsOn = true;
        GoTo(0);
        return true;
    }

    public void Stop()
    {
        if (!IsOn)
        {
            return;
        }

        IsOn = false;
        hidden.Clear();
        dimmed.Clear();
        steps.Clear();
        RefreshFigures();
        SelectionFlash.StopAll(drawing.Canvas);
        Changed?.Invoke();
    }

    public void Next()
    {
        GoTo(current + 1);
    }

    public void Previous()
    {
        GoTo(current - 1);
    }

    public void First()
    {
        GoTo(0);
    }

    public void Last()
    {
        GoTo(steps.Count - 1);
    }

    public void GoTo(int index)
    {
        if (!IsOn || steps.Count == 0)
        {
            return;
        }

        current = System.Math.Clamp(index, 0, steps.Count - 1);
        Apply();
        FlashCurrent();
        Changed?.Invoke();
    }

    /// <summary>
    /// The steps read off the list again (a figure added, deleted, grouped or excluded):
    /// the view stays at the same step where the steps are as they were. Stops when no
    /// step is left.
    /// </summary>
    public void Refresh()
    {
        if (!IsOn)
        {
            return;
        }

        var before = steps.Select(step => step.Figures.ToList()).ToList();
        Partition();
        if (steps.Count == 0)
        {
            Stop();
            return;
        }

        bool same = before.Count == steps.Count
            && before.Zip(steps, (old, step) => old.SequenceEqual(step.Figures)).All(equal => equal);
        if (same)
        {
            return;
        }

        current = System.Math.Clamp(current, 0, steps.Count - 1);
        Apply();
        Changed?.Invoke();
    }

    /// <summary>Whether the figure is of a step after the current one, and so not drawn</summary>
    public bool IsHidden(IFigure figure)
    {
        return IsOn && hidden.Contains(figure);
    }

    /// <summary>How much of the figure shows: dimmed when its step is already taken</summary>
    public double OpacityOf(IFigure figure)
    {
        return IsOn && dimmed.Contains(figure) ? DimOpacity : 1;
    }

    /// <summary>The step the figure is in, or -1</summary>
    public int StepOf(IFigure figure)
    {
        var whole = drawing.Figures.FindTopLevel(FigureParts.Whole(figure)) ?? figure;
        return steps.FindIndex(step => step.Figures.Contains(whole));
    }

    /// <summary>
    /// The number of the step the figure begins, from 1, for the Figure List; 0 for a
    /// figure that is grouped with the one before it, hidden or no step
    /// </summary>
    public int NumberOf(IFigure figure)
    {
        if (!IsOn)
        {
            return 0;
        }

        int index = steps.FindIndex(step => step.Figures.Count > 0 && step.Figures[0] == figure);
        return index + 1;
    }

    void Partition()
    {
        steps.Clear();
        steps.AddRange(ConstructionSteps.Partition(drawing));
    }

    void Apply()
    {
        hidden.Clear();
        dimmed.Clear();
        for (int i = 0; i < steps.Count; i++)
        {
            var set = i > current ? hidden : i < current ? dimmed : null;
            if (set == null)
            {
                continue;
            }

            foreach (var figure in steps[i].Figures)
            {
                AddWithParts(set, figure);
            }
        }

        RefreshFigures();
    }

    /// <summary>The figure, the shapes it is made of (a slider's knob and track) and its name</summary>
    static void AddWithParts(HashSet<IFigure> set, IFigure figure)
    {
        set.Add(figure);
        if (figure is FigureBase figureBase && figureBase.NameLabel != null)
        {
            set.Add(figureBase.NameLabel);
        }

        if (figure is CompositeFigure composite)
        {
            foreach (var child in composite.Children)
            {
                AddWithParts(set, child);
            }
        }
    }

    void RefreshFigures()
    {
        foreach (var figure in drawing.Figures.OfType<FigureBase>())
        {
            figure.RefreshShown();
        }
    }

    /// <summary>
    /// The current step's figures flash: only they, not what they are built on (flashed
    /// too at first, the parents drew the eye away from the new figure). Nothing is
    /// selected: the halos stayed on while the step was read and took the eye (2026-11-03).
    /// </summary>
    void FlashCurrent()
    {
        var step = steps[current];
        var shown = step.Figures
            .Select(figure => drawing.Figures.FindTopLevel(figure) ?? figure)
            .Where(figure => figure.Visible && !hidden.Contains(figure))
            .Distinct()
            .ToList();
        SelectionFlash.Show(shown, SelectionFlash.StepFade);
    }

    /// <summary>A figure picked in the Figure List takes the stepping to its step (a figure of no step leaves it)</summary>
    public void GoToFigure(IFigure figure)
    {
        int step = StepOf(figure);
        if (step >= 0)
        {
            GoTo(step);
        }
    }
}
