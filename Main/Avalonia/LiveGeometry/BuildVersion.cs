using System.Reflection;

namespace LiveGeometry;

/// <summary>
/// Which build this is. The .NET SDK puts the git commit into the assembly's informational
/// version ("1.0.0+5e59c77..."), so there is nothing to maintain.
/// </summary>
public static class BuildVersion
{
    /// <summary>The whole informational version, e.g. "1.0.0+5e59c77a1b..."</summary>
    public static string Full { get; } =
        typeof(BuildVersion).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion
        ?? "unknown";
}
