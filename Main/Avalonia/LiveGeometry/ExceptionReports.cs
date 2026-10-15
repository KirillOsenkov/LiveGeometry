using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Channels;
using System.Threading.Tasks;

namespace LiveGeometry;

/// <summary>
/// First-chance exceptions sent to the site's API (POST {ShareService}/api/v1/exceptions),
/// which buckets them by stack and keeps them: what the app shows as an error report
/// (<see cref="MainView"/>, CurrentDomain_FirstChanceException, after IsBenign), from the
/// desktop and the browser alike. A report is the exception's type, its message (shortened)
/// and the stack at the throw as method names, without file paths. One the same as any of the
/// last few isn't sent again (an exception thrown on every mouse move goes once), and a
/// session sends a limited number. They queue up, and a background task sends them in batches
/// a few seconds apart. Nothing the reporting throws is reported (<see cref="IsOwn"/>): not
/// while a report is made, not what the sender's code throws, not a network failure while a
/// batch is on its way. Only a build given the API's address sends anything.
/// </summary>
public static class ExceptionReports
{
    /// <summary>How many reports back a new one is compared with</summary>
    const int Remembered = 10;

    const int MaxPerSession = 50;
    const int MaxBatch = 20;
    const int MaxFrames = 100;
    const int MaxMessageLength = 1000;

    /// <summary>How long the sender waits after the first report of a burst, for the rest to join it</summary>
    static readonly TimeSpan GatherDelay = TimeSpan.FromSeconds(5);

    /// <summary>The API's address (MainView.ShareService), set at startup; null sends nothing</summary>
    public static string Address { get; set; }

    record Report(string Type, string Message, List<string> Frames);

    static readonly Channel<Report> Queue = Channel.CreateBounded<Report>(new BoundedChannelOptions(MaxPerSession)
    {
        FullMode = BoundedChannelFullMode.DropWrite,
        SingleReader = true
    });

    static readonly HttpClient Client = new HttpClient() { Timeout = TimeSpan.FromSeconds(30) };
    static readonly Lock Gate = new();
    static readonly string[] Recent = new string[Remembered];
    static int recentNext;
    static int added;
    static Task sender;

    /// <summary>Flows into everything the sender awaits</summary>
    static readonly AsyncLocal<bool> Sending = new();

    /// <summary>Batches on their way: a network failure meanwhile is the sender's, whatever thread it comes on</summary>
    static int inFlight;

    /// <summary>Set while <see cref="Add"/> runs: an exception it throws itself comes back through the handler</summary>
    [ThreadStatic]
    static bool adding;

    /// <summary>Whether the exception is the reporting's own: neither shown nor reported</summary>
    public static bool IsOwn(Exception exception)
    {
        return adding || Sending.Value || Volatile.Read(ref inFlight) > 0 && IsNetworkFailure(exception);
    }

    /// <summary>A request that failed on its way (what the app says in words instead of an error report)</summary>
    public static bool IsNetworkFailure(Exception exception)
    {
        return exception is HttpRequestException
            || exception is TaskCanceledException
            || exception is System.Net.Sockets.SocketException
            || exception is IOException
            || exception is System.Security.Authentication.AuthenticationException
            || exception.GetType().Name == "JSException";
    }

    /// <summary>Called by the first-chance handler, at the throw; never throws</summary>
    public static void Add(Exception exception)
    {
        if (Address == null || adding)
        {
            return;
        }

        adding = true;
        try
        {
            var message = exception.Message ?? "";
            if (message.Length > MaxMessageLength)
            {
                message = message.Substring(0, MaxMessageLength);
            }

            var report = new Report(exception.GetType().FullName ?? exception.GetType().Name, message, CaptureFrames());
            var key = report.Type + "\n" + report.Message + "\n" + string.Join("\n", report.Frames);
            lock (Gate)
            {
                if (Array.IndexOf(Recent, key) >= 0 || added >= MaxPerSession)
                {
                    return;
                }

                Recent[recentNext] = key;
                recentNext = (recentNext + 1) % Remembered;
                added++;
                Queue.Writer.TryWrite(report);
                sender ??= Task.Run(Send);
            }
        }
        catch
        {
            // a report that can't be made is no reason to make another
        }
        finally
        {
            adding = false;
        }
    }

    /// <summary>
    /// The stack at the throw, innermost first, as Namespace.Type.Method: the frames past the
    /// first-chance handler (those up to it are the reporting itself) and the runtime's
    /// exception dispatch under it (System.Runtime.EH.DispatchEx on the desktop, the same for
    /// every exception), from the throw site out. The exception's own trace is no use here: at
    /// first chance it has only the throwing frame, in the browser often nothing.
    /// </summary>
    static List<string> CaptureFrames()
    {
        var frames = new StackTrace(fNeedFileInfo: false).GetFrames();
        int start = Array.FindIndex(frames, frame => frame.GetMethod()?.Name == "CurrentDomain_FirstChanceException") + 1;
        while (start < frames.Length && IsExceptionDispatch(frames[start].GetMethod()?.DeclaringType))
        {
            start++;
        }

        var names = new List<string>();
        for (int i = start; i < frames.Length && names.Count < MaxFrames; i++)
        {
            var method = frames[i].GetMethod();
            if (method != null && method.DeclaringType != typeof(ExceptionReports))
            {
                names.Add((method.DeclaringType?.FullName ?? "?") + "." + method.Name);
            }
        }

        return names;
    }

    static bool IsExceptionDispatch(Type type)
    {
        // DispatchEx is a method of the class System.Runtime.EH
        return type?.FullName == "System.Runtime.EH" || type?.Namespace == "System.Runtime.ExceptionServices";
    }

    static async Task Send()
    {
        Sending.Value = true;
        var reader = Queue.Reader;
        while (await reader.WaitToReadAsync())
        {
            await Task.Delay(GatherDelay);
            var batch = new List<Report>();
            while (batch.Count < MaxBatch && reader.TryRead(out var report))
            {
                batch.Add(report);
            }

            Interlocked.Increment(ref inFlight);
            try
            {
                // text/plain: a request the browser sends without asking the server first (no preflight)
                using var content = new StringContent(ToJson(batch), Encoding.UTF8, "text/plain");
                using var response = await Client.PostAsync(Address + "/api/v1/exceptions", content);

                // sent, refused or limited, a batch isn't sent again
            }
            catch
            {
                // offline or refused: these are lost, which is fine
            }
            finally
            {
                Interlocked.Decrement(ref inFlight);
            }
        }
    }

    /// <summary>{ build, platform, reports: [ { type, message, frames } ] }, written by hand: no reflection for the trimmer to break</summary>
    static string ToJson(List<Report> batch)
    {
        using var stream = new MemoryStream();
        using (var json = new Utf8JsonWriter(stream))
        {
            json.WriteStartObject();
            json.WriteString("build", BuildVersion.Full);
            json.WriteString("platform", OperatingSystem.IsBrowser() ? "browser" : "desktop " + RuntimeInformation.OSDescription);
            json.WriteStartArray("reports");
            foreach (var report in batch)
            {
                json.WriteStartObject();
                json.WriteString("type", report.Type);
                json.WriteString("message", report.Message);
                json.WriteStartArray("frames");
                foreach (var frame in report.Frames)
                {
                    json.WriteStringValue(frame);
                }

                json.WriteEndArray();
                json.WriteEndObject();
            }

            json.WriteEndArray();
            json.WriteEndObject();
        }

        return Encoding.UTF8.GetString(stream.ToArray());
    }
}
