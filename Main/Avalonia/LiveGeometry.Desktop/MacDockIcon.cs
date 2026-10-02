using System;
using System.IO;
using System.Runtime.InteropServices;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Media.Imaging;

namespace LiveGeometry.Desktop;

/// <summary>
/// The app's mark in a Mac's Dock and app switcher. Avalonia ignores a window's icon there
/// (an app bundle's Info.plist would name one), and the app run as a plain executable showed
/// a blank document: the vector mark (<see cref="AppIcon"/>) is drawn into a picture and given
/// to NSApplication as its icon image.
/// </summary>
public static class MacDockIcon
{
    const int PixelSize = 512;

    public static void Apply()
    {
        if (!OperatingSystem.IsMacOS())
        {
            return;
        }

        try
        {
            var png = Render();
            var data = Send(Send(objc_getClass("NSData"), Selector("alloc")), Selector("initWithBytes:length:"), png, (nuint)png.Length);
            var image = Send(Send(objc_getClass("NSImage"), Selector("alloc")), Selector("initWithData:"), data);
            Send(data, Selector("release"));
            if (image == IntPtr.Zero)
            {
                Console.WriteLine("Dock icon: the picture was not taken");
                return;
            }

            var application = Send(objc_getClass("NSApplication"), Selector("sharedApplication"));
            Send(application, Selector("setApplicationIconImage:"), image);
            Send(image, Selector("release"));
        }
        catch (Exception ex)
        {
            // never a reason not to start
            Console.WriteLine("Dock icon: " + ex.Message);
        }
    }

    /// <summary>The mark with a margin, as the icons beside it in the Dock have.</summary>
    static byte[] Render()
    {
        var icon = new Border()
        {
            Width = PixelSize,
            Height = PixelSize,
            Padding = new Thickness(PixelSize / 10),
            Child = AppIcon.Create(PixelSize * 0.8)
        };
        icon.Measure(new Size(PixelSize, PixelSize));
        icon.Arrange(new Rect(0, 0, PixelSize, PixelSize));

        using var bitmap = new RenderTargetBitmap(new PixelSize(PixelSize, PixelSize), new Vector(96, 96));
        bitmap.Render(icon);
        using var stream = new MemoryStream();
        bitmap.Save(stream, new PngBitmapEncoderOptions());
        return stream.ToArray();
    }

    static IntPtr Selector(string name) => sel_registerName(name);

    const string ObjectiveC = "/usr/lib/libobjc.A.dylib";

    [DllImport(ObjectiveC)]
    static extern IntPtr objc_getClass(string name);

    [DllImport(ObjectiveC)]
    static extern IntPtr sel_registerName(string name);

    [DllImport(ObjectiveC, EntryPoint = "objc_msgSend")]
    static extern IntPtr Send(IntPtr receiver, IntPtr selector);

    [DllImport(ObjectiveC, EntryPoint = "objc_msgSend")]
    static extern IntPtr Send(IntPtr receiver, IntPtr selector, IntPtr argument);

    [DllImport(ObjectiveC, EntryPoint = "objc_msgSend")]
    static extern IntPtr Send(IntPtr receiver, IntPtr selector, byte[] bytes, nuint length);
}
