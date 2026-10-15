using System;

namespace LiveGeometry;

/// <summary>
/// The page's splash screen (the browser's index.html), which stays up until the app has
/// drawn what the page opened at: the gallery with the tiles in view loaded
/// (<see cref="GalleryView"/>), or a drawing laid out (<see cref="MainView"/>). Only the
/// browser head has one and says how to drive it (<see cref="Attach"/>); on the desktop
/// nothing is up and the calls do nothing. Avalonia itself takes its splash down at its
/// first frame, which on the gallery page is seconds before the tiles are in.
/// </summary>
public static class SplashScreen
{
    static Action hide;
    static Action<double> reportProgress;

    /// <summary>The splash is up: the app is not on screen yet</summary>
    public static bool IsUp { get; private set; }

    /// <param name="hideSplash">Takes the splash down</param>
    /// <param name="reportProgress">Moves the bar under the construction: how far the app is, 0 to 1</param>
    public static void Attach(Action hideSplash, Action<double> reportProgress)
    {
        hide = hideSplash;
        SplashScreen.reportProgress = reportProgress;
        IsUp = true;
    }

    public static void ReportProgress(double fraction)
    {
        if (IsUp)
        {
            reportProgress(System.Math.Clamp(fraction, 0, 1));
        }
    }

    public static void Hide()
    {
        if (IsUp)
        {
            IsUp = false;
            hide();
        }
    }
}
