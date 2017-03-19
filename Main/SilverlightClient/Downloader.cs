using System;
using System.Net;

namespace LiveGeometry
{
    public static class Downloader
    {
        public static void DownloadString(string url, Action<string> callback)
        {
            WebClient internet = new WebClient();
            internet.DownloadStringCompleted += new DownloadStringCompletedEventHandler(internet_DownloadStringCompleted);
            internet.DownloadStringAsync(new Uri(url), callback);
        }

        static void internet_DownloadStringCompleted(object sender, DownloadStringCompletedEventArgs e)
        {
            if (e.Cancelled || e.Error != null)
            {
                return;
            }
            var text = e.Result;
            Action<string> callback = e.UserState as Action<string>;
            if (callback != null)
            {
                callback(text);
            }
        }
    }
}
