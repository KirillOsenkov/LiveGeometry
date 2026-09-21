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

    /// <summary>The first 7 characters of the commit, or the version if there is no commit in it</summary>
    public static string Short
    {
        get
        {
            int plus = Full.IndexOf('+');
            if (plus < 0 || plus + 1 >= Full.Length)
            {
                return Full;
            }

            var commit = Full.Substring(plus + 1);
            return commit.Length > 7 ? commit.Substring(0, 7) : commit;
        }
    }
}
