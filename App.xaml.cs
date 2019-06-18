using System.Windows;

namespace Thesis
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public static UserSettings Settings;
        public static readonly string AppName = "Thesis";

        public App()
        {
            //Register Syncfusion license
            Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("***REMOVED***");
            Settings = UserSettings.Read();
        }
    }
}
