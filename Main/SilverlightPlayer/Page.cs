using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Browser;
using System.Windows.Controls;
using System.Windows.Media;
using System.Xml.Linq;
using DynamicGeometry;

namespace LiveGeometry
{
    public partial class Page : UserControl
    {
        public Page()
            : this(null)
        {
        }

        private void CreateCanvas()
        {
            Canvas = new Canvas();
            Canvas.Background = new SolidColorBrush(Colors.White);
            Canvas.HorizontalAlignment = HorizontalAlignment.Stretch;
            Canvas.VerticalAlignment = VerticalAlignment.Stretch;

            Canvas.SizeChanged += Canvas_SizeChanged;
        }

        void Canvas_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (CurrentDrawing == null)
            {
                CurrentDrawing = new Drawing(Canvas);
                Canvas.SizeChanged -= Canvas_SizeChanged;
                if (InitParams.ContainsKey("LoadFile"))
                {
                    string url = InitParams["LoadFile"];
                    DownloadAndDisplayDrawing(url);
                }
                else
                {
                    DownloadAndDisplayDrawing("http://guilabs.de/geometry/drawings/fun/car.lgf");
                }
            }
        }

        Drawing mCurrentDrawing;
        public Drawing CurrentDrawing
        {
            get
            {
                return mCurrentDrawing;
            }
            set
            {
                if (mCurrentDrawing != null)
                {
                    mCurrentDrawing.DocumentOpenRequested -= mCurrentDrawing_DocumentOpenRequested;
                    mCurrentDrawing.UnhandledException -= drawingHost_UnhandledException;
                    mCurrentDrawing.Canvas = null;
                }
                mCurrentDrawing = value;
                if (mCurrentDrawing != null)
                {
                    mCurrentDrawing.DocumentOpenRequested += mCurrentDrawing_DocumentOpenRequested;
                    mCurrentDrawing.UnhandledException += drawingHost_UnhandledException;
                }
            }
        }

        void mCurrentDrawing_DocumentOpenRequested(object sender, Drawing.DocumentOpenRequestedEventArgs e)
        {
            LoadDrawing(e.DocumentXml);
        }

        public void LoadDrawing(string drawingXml)
        {
            XElement xml = null;
            try
            {
                xml = XElement.Parse(drawingXml);
            }
            catch (Exception)
            {
                return;
            }
            LoadDrawing(xml);
        }

        public void Clear()
        {
            CurrentDrawing = new Drawing(Canvas);
            CurrentDrawing.Behavior = new Dragger();
        }

        public void LoadDrawing(XElement element)
        {
            try
            {
                Clear();
                CurrentDrawing.AddFromXml(element);
            }
            catch (Exception ex)
            {
                CurrentDrawing.RaiseError(this, ex);
            }
        }

        Canvas Canvas;

        public Page(IDictionary<string, string> initParams)
        {
            CreateCanvas();
            UseLayoutRounding = true;
            this.Content = Canvas;
            InitParams = initParams;
            var settings = Application.Current.Host.Settings;
            if (Application.Current.IsRunningOutOfBrowser)
            {
                settings.EnableAutoZoom = false;
            }

            //settings.EnableGPUAcceleration = true;
            //settings.EnableRedrawRegions = true;
            //settings.EnableCacheVisualization = true;
            //settings.EnableFrameRateCounter = true;

            //DynamicGeometry.Settings.Instance = new IsolatedStorageBasedSettings();
            //PageSettings = new Settings(this);
            //if (initParams.ContainsKey("ShowToolbar"))
            //{
            //    PageSettings.ShowToolbar = bool.Parse(initParams["ShowToolbar"]);
            //}

            //drawingHost.ReadyForInteraction += drawingHost_ReadyForInteraction;
            //drawingHost.UnhandledException += drawingHost_UnhandledException;

            //InitializeToolbar();

            //this.KeyDown += Page_KeyDown;
            DownloadDemoFile();
            //IsolatedStorage.LoadAllTools();
        }

        void drawingHost_UnhandledException(object sender, UnhandledExceptionNotificationEventArgs e)
        {

        }

        //public static LiveGeometrySoapClient liveGeometryWebServices;

        //void drawingHost_ReadyForInteraction(object sender, EventArgs e)
        //{
        //liveGeometryWebServices = new LiveGeometrySoapClient();
        //liveGeometryWebServices.SendErrorReportCompleted += liveGeometryWebServices_SendErrorReportCompleted;
        //}

        //void liveGeometryWebServices_SendErrorReportCompleted(object sender, SendErrorReportCompletedEventArgs e)
        //{
        //    if (e.Result == "OK")
        //    {
        //        drawingHost.ShowHint("Sending the error report completed successfully.");
        //    }
        //    else
        //    {
        //        drawingHost.ShowHint("Sending the error report failed: " + e.Result
        //        + "\n"
        //        + (e.Error == null ? "" : e.Error.ToString()));
        //    }
        //}

        //DrawingHost drawingHost = new DrawingHost();

        public IDictionary<string, string> InitParams { get; set; }

        public const string homePage = "http://livegeometry.codeplex.com";

        void HomepageButton_Click(object sender, RoutedEventArgs e)
        {
            Homepage();
        }

        void Homepage()
        {
            try
            {
                if (Application.Current.IsRunningOutOfBrowser)
                {
                    ProgrammaticHyperlinkButton.Navigate(homePage);
                }
                else
                {
                    HtmlPage.Window.Navigate(new Uri("http://livegeometry.codeplex.com"), "_blank");
                }
            }
            catch (Exception)
            {
            }
        }
    }

    public class ProgrammaticHyperlinkButton : HyperlinkButton
    {
        public ProgrammaticHyperlinkButton(string url)
        {
            base.NavigateUri = new Uri(url);
            TargetName = "_blank";
        }

        public new void Click()
        {
            base.OnClick();
        }

        public static void Navigate(string url)
        {
            new ProgrammaticHyperlinkButton(url).Click();
        }
    }
}
