using System;
using System.Threading.Tasks;
using System.Windows;

namespace RSS_Image_Interoperability_Tool
{
    public partial class App : Application
    {
        public App()
        {
            this.DispatcherUnhandledException += (s, e) =>
            {
                MessageBox.Show(e.Exception.ToString(), "Unhandled UI Exception");
                e.Handled = true;
            };
            AppDomain.CurrentDomain.UnhandledException += (s, e) =>
                MessageBox.Show(e.ExceptionObject.ToString()!, "Unhandled Domain Exception");
            TaskScheduler.UnobservedTaskException += (s, e) =>
            {
                MessageBox.Show(e.Exception.ToString(), "Unhandled Task Exception");
                e.SetObserved();
            };
        }
    }
}
