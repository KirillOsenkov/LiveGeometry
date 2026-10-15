using System;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using DynamicGeometry;

namespace LiveGeometry;

/// <summary>
/// Export > Copy share link: the drawing stored by the site's API under a short id made from
/// its content, and its link, https://livegeometry.com/s/&lt;id&gt;, on the clipboard. The page
/// at the link plays it, with Edit to bring it here (web/index.html, MainView.Handoff.cs); the
/// same drawing again gets the same link. Only a build given the API's address has it
/// (ShareService in LiveGeometry.csproj): other builds never call the service.
/// </summary>
public partial class MainView
{
    /// <summary>The API's address the build was given, or null: no Copy share link then</summary>
    public static readonly string ShareService = ReadShareService();

    const string ShareLinkPrefix = "https://livegeometry.com/s/";

    /// <summary>
    /// The browser head's: starts writing to the clipboard at once, in the click, with text
    /// that comes later (Safari writes only during a click, and the link takes a request);
    /// null text fails it. Null on the desktop, whose clipboard takes the text when it comes.
    /// </summary>
    public static Func<Task<string>, Task> CopyTextWhenReady { get; set; }

    static readonly HttpClient ShareClient = new HttpClient() { Timeout = TimeSpan.FromSeconds(30) };

    // the drawing shared last and its link: the same drawing again needs no request
    string sharedText;
    string sharedLink;

    /// <summary>
    /// Links being made: meanwhile a request that fails and a clipboard that refuses are said
    /// in the status, not reported as errors (<see cref="IsBenign"/>)
    /// </summary>
    static int sharingCount;

    static string ReadShareService()
    {
        var value = typeof(MainView).Assembly
            .GetCustomAttributes<AssemblyMetadataAttribute>()
            .FirstOrDefault(attribute => attribute.Key == "ShareService")?.Value;
        return string.IsNullOrEmpty(value) ? null : value.TrimEnd('/');
    }

    static bool IsSharingFailure(Exception exception)
    {
        return ExceptionReports.IsNetworkFailure(exception);
    }

    async void CopyShareLink()
    {
        sharingCount++;
        try
        {
            var text = DrawingTextToSave();

            // the request and the clipboard write both start here, in the click
            var result = text == sharedText ? Task.FromResult((sharedLink, (string)null)) : MakeShareLink(text);
            var copied = CopyTextWhenReady?.Invoke(LinkOf(result));
            DrawingHost.ShowHint("Making a link...");
            var (link, problem) = await result;
            if (link == null)
            {
                // the clipboard write fails without its text
                await Succeeds(copied);
                DrawingHost.ShowHint(problem);
                return;
            }

            sharedText = text;
            sharedLink = link;
            if (copied == null || await Succeeds(copied))
            {
                if (copied == null)
                {
                    Clipboard.SetText(link);
                }

                DrawingHost.ShowHint("Share link copied: " + link);
                return;
            }

            // a browser without ClipboardItem: its clipboard may still take the text now
            Clipboard.SetText(link);
            DrawingHost.ShowHint("Share link: " + link);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message);
        }
        finally
        {
            sharingCount--;
        }
    }

    static async Task<string> LinkOf(Task<(string Link, string Problem)> result)
    {
        return (await result).Link;
    }

    static async Task<bool> Succeeds(Task task)
    {
        if (task == null)
        {
            return false;
        }

        try
        {
            await task;
            return true;
        }
        catch (Exception)
        {
            return false;
        }
    }

    /// <summary>The link, or why there is none, in words for the status</summary>
    static async Task<(string Link, string Problem)> MakeShareLink(string text)
    {
        HttpResponseMessage response;
        try
        {
            // text/plain: a request the browser sends without asking the server first (no preflight)
            using var content = new StringContent(text, Encoding.UTF8, "text/plain");
            response = await ShareClient.PostAsync(ShareService + "/api/v1/share", content);
        }
        catch (Exception ex) when (IsSharingFailure(ex))
        {
            return (null, "Couldn't reach the server to make a link. Check the connection and try again.");
        }

        using (response)
        {
            switch ((int)response.StatusCode)
            {
                case 413:
                    return (null, "This drawing is too big to share: over 512 KB.");
                case 429:
                    return (null, "Too many links in a minute: try again in a moment.");
            }

            if (!response.IsSuccessStatusCode)
            {
                return (null, "The server couldn't make a link (" + (int)response.StatusCode + ").");
            }

            using var json = JsonDocument.Parse(await response.Content.ReadAsStringAsync());
            return (ShareLinkPrefix + json.RootElement.GetProperty("id").GetString(), null);
        }
    }
}
