using System;
using System.IO.IsolatedStorage;
using System.Windows;

namespace LiveGeometry
{
    public class LiveGeometryApp : Application
    {
        public LiveGeometryApp()
        {
            this.Startup += this.Application_Startup;
            this.Exit += this.Application_Exit;
            this.UnhandledException += this.Application_UnhandledException;
        }

        private void Application_Startup(object sender, StartupEventArgs e)
        {
            this.RootVisual = new Page(e.InitParams);
        }

        private void Application_Exit(object sender, EventArgs e)
        {
            IsolatedStorageSettings.ApplicationSettings.Save();
        }

        private void Application_UnhandledException(object sender, ApplicationUnhandledExceptionEventArgs e)
        {
        }
    }
}
