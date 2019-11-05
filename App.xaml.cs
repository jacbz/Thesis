using System.Windows;
using Thesis.Models;
using Thesis.ViewModels;

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
            // Register Syncfusion license
            Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense(Thesis.Properties.Resources.SyncfusionLicenseKey);
        }
    }
}
