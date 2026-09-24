using System;
using System.Collections.Generic;
using Avalonia.Input;

namespace DynamicGeometry;

/// <summary>
/// Single-letter tool shortcuts - the same letters the original VB6 DG used, for the
/// tools that still exist. Except G: VB6 had it on Distance, here it toggles the grid
/// (MainView.HandlePlainKey), which is reached for far more often.
/// </summary>
public static class BehaviorShortcuts
{
    static readonly Dictionary<Key, Type> tools = new Dictionary<Key, Type>()
    {
        { Key.Q, typeof(Dragger) },
        { Key.P, typeof(FreePointCreator) },
        { Key.S, typeof(SegmentCreator) },
        { Key.Y, typeof(RayCreator) },
        { Key.L, typeof(LineTwoPointsCreator) },
        { Key.N, typeof(ParallelLineCreator) },
        { Key.E, typeof(PerpendicularLineCreator) },
        { Key.B, typeof(AngleBisectorCreator) },
        { Key.C, typeof(CircleCreator) },
        { Key.R, typeof(CircleByRadiusCreator) },
        { Key.A, typeof(CircleArcCreator) },
        { Key.M, typeof(MidpointCreator) },
        { Key.T, typeof(ReflectionCreator) },
        { Key.D, typeof(LocusCreator) },
        { Key.W, typeof(PolygonCreator) },
        { Key.J, typeof(AngleMeasurementCreator) },
        { Key.K, typeof(AreaMeasurementCreator) },
    };

    /// <returns>The type of the tool the key selects, or null</returns>
    public static Type GetTool(Key key)
    {
        tools.TryGetValue(key, out var result);
        return result;
    }

    /// <returns>The shortcut letter of the tool, or null if it has none</returns>
    public static string GetShortcut(Behavior behavior)
    {
        foreach (var pair in tools)
        {
            if (pair.Value == behavior.GetType())
            {
                return pair.Key.ToString();
            }
        }

        return null;
    }
}
