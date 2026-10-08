using System;
using System.Net;
using System.Threading.Tasks;

namespace LiveGeometry;

/// <summary>
/// A drawing on another web page: the JavaScript player (Main/Player, served as
/// /player/<version>/player.js) and the drawing's own text inside an element it plays. The
/// snippet (Export > Copy embed code) refers to the player on livegeometry.com, pinned to a
/// version folder so that it keeps playing when the format moves on; the page (Save as
/// .html) carries the player itself, one file that works from a disk.
/// </summary>
public static class PlayerEmbed
{
    /// <summary>The player's version folder: a new one when the .lgf format breaks the old player (AGENTS.md, "The JavaScript player")</summary>
    public const string Version = "1";

    public const string ScriptUrl = "https://livegeometry.com/player/" + Version + "/player.js";

    /// <summary>The text of player.js, where the head keeps it: beside the exe on the desktop, fetched from the site in the browser</summary>
    public static Func<Task<string>> OpenPlayerScript { get; set; }

    /// <summary>
    /// The element and the script tag, to paste into any HTML page. The drawing goes in as
    /// is: XML escapes every "&lt;" in text and attributes, so "&lt;/script" can't occur in
    /// it and the tag ends where it should.
    /// </summary>
    public static string CreateSnippet(string drawingText, int height)
    {
        return "<script src=\"" + ScriptUrl + "\" defer></script>\n"
            + "<div class=\"livegeometry\" style=\"width: 100%; height: " + height + "px;\">\n"
            + "<script type=\"text/x-livegeometry\">\n"
            + drawingText.TrimEnd() + "\n"
            + "</script>\n"
            + "</div>\n";
    }

    /// <summary>A page of its own with the player written into it: nothing to fetch, so it plays from a disk too</summary>
    public static string CreatePage(string title, string drawingText, string playerScript)
    {
        // a "</script" inside the player's own text (a string, a comment) would end the tag;
        // the backslash means nothing to JavaScript there
        var script = playerScript.Replace("</script", "<\\/script");
        return "<!DOCTYPE html>\n"
            + "<html lang=\"en\">\n"
            + "<head>\n"
            + "<meta charset=\"utf-8\">\n"
            + "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n"
            + "<title>" + WebUtility.HtmlEncode(title) + "</title>\n"
            + "<style>html, body { margin: 0; height: 100%; } .livegeometry { width: 100%; height: 100%; }</style>\n"
            + "<script>\n" + script.TrimEnd() + "\n</script>\n"
            + "</head>\n"
            + "<body>\n"
            + "<div class=\"livegeometry\">\n"
            + "<script type=\"text/x-livegeometry\">\n"
            + drawingText.TrimEnd() + "\n"
            + "</script>\n"
            + "</div>\n"
            + "</body>\n"
            + "</html>\n";
    }
}
