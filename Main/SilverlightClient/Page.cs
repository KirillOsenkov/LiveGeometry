using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Browser;
using System.Windows.Controls;
using DynamicGeometry;
using LiveGeometry.LiveGeometryWebServices;

namespace LiveGeometry
{
    public partial class Page : UserControl
    {
        private DrawingHost drawingHost;

        private DrawingControl DrawingControl
        {
            get
            {
                return drawingHost.DrawingControl;
            }
        }

        public Page()
            : this(null)
        {
        }

        public Page(IDictionary<string, string> initParams)
        {
            UseLayoutRounding = true;

            MEFHost.Instance.RegisterExtensionAssemblyFromType<Page>();

            // this needs to be initialized first since subsequent calls 
            // rely on having serializable settings already available
            DynamicGeometry.Settings.Instance = new IsolatedStorageBasedSettings();

            drawingHost = new DrawingHost();
            AddBehaviors();
            this.Content = drawingHost;
            
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

            PageSettings = new Settings(this);
            if (initParams.ContainsKey("ShowToolbar"))
            {
                PageSettings.ShowToolbar = bool.Parse(initParams["ShowToolbar"]);
            }

            drawingHost.ReadyForInteraction += drawingHost_ReadyForInteraction;
            drawingHost.UnhandledException += drawingHost_UnhandledException;

            InitializeToolbar();

            this.KeyDown += Page_KeyDown;
            DownloadDemoFile();
            IsolatedStorage.LoadAllTools();
            IsolatedStorage.RegisterToolStorage();
        }

        private void AddBehaviors()
        {
            var behaviors = Behavior.LoadBehaviors(typeof(Dragger).Assembly);
            var dragger = behaviors.First(b => b is Dragger);
            Behavior.Default = dragger;
            foreach (var behavior in behaviors)
            {
                drawingHost.AddToolButton(behavior);
            }
        }

        void drawingHost_UnhandledException(object sender, UnhandledExceptionNotificationEventArgs e)
        {
            HandleException(e.Exception);
        }

        public static LiveGeometrySoapClient liveGeometryWebServices;

        void drawingHost_ReadyForInteraction(object sender, EventArgs e)
        {
            if (InitParams.ContainsKey("LoadFile"))
            {
                string url = InitParams["LoadFile"];
                DownloadAndDisplayDrawing(url);
            }
            liveGeometryWebServices = new LiveGeometrySoapClient();
            liveGeometryWebServices.SendErrorReportCompleted += liveGeometryWebServices_SendErrorReportCompleted;
        }

        void liveGeometryWebServices_SendErrorReportCompleted(object sender, SendErrorReportCompletedEventArgs e)
        {
            if (e.Result == "OK")
            {
                drawingHost.ShowHint("Sending the error report completed successfully.");
            }
            else
            {
                drawingHost.ShowHint("Sending the error report failed: " + e.Result
                + "\n"
                + (e.Error == null ? "" : e.Error.ToString()));
            }
        }

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
