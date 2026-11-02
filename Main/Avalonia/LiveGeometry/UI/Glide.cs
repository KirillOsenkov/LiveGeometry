using System;
using System.Diagnostics;
using Avalonia.Threading;

namespace LiveGeometry;

/// <summary>
/// A short movement by timer: the step is told how far along it is in time, from 0 to 1, at
/// every tick and once more at 1. A move of the view eases out (<see cref="EaseOut"/>); a
/// fade goes linearly (<see cref="SelectionFlash"/>).
/// </summary>
public class Glide
{
    /// <summary>What a move of the view takes</summary>
    public static readonly TimeSpan ViewDuration = TimeSpan.FromMilliseconds(220);

    readonly Action<double> step;
    readonly TimeSpan duration;
    readonly Stopwatch watch = Stopwatch.StartNew();
    DispatcherTimer timer;

    public Glide(TimeSpan duration, Action<double> step)
    {
        this.duration = duration;
        this.step = step;
        timer = new DispatcherTimer(TimeSpan.FromMilliseconds(15), DispatcherPriority.Render, (s, e) => Tick());
        timer.Start();
        Tick();
    }

    /// <summary>Quick at first, settling at the end</summary>
    public static double EaseOut(double t)
    {
        return 1 - System.Math.Pow(1 - t, 3);
    }

    void Tick()
    {
        double t = System.Math.Min(1, watch.Elapsed.TotalMilliseconds / duration.TotalMilliseconds);
        if (t >= 1)
        {
            Finish();
            return;
        }

        step(t);
    }

    /// <summary>To the end at once</summary>
    public void Finish()
    {
        if (timer == null)
        {
            return;
        }

        timer.Stop();
        timer = null;
        step(1);
    }
}
