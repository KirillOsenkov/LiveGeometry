using System;

namespace LiveGeometry
{
    public partial class Page
    {
        private void DownloadDemoFile()
        {
            //CommandSamples.Enabled = false;
            Downloader.DownloadString(
                "http://www.osenkov.com/geometry/demo/Demo.xml",
                xml =>
                {
                    demoText = xml;
                    //CommandSamples.Enabled = true;
                });
        }

        public void DownloadAndDisplayDrawing(string url)
        {
            try
            {
                Downloader.DownloadString(url, LoadDrawing);
            }
            catch (Exception)
            {
                //HandleException(ex);
            }
        }

        string demoText;

        //void Samples()
        //{
        //    drawingHost.LoadDrawing(demoText);
        //}
    }
}
