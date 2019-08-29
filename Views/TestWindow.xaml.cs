using System.Windows;

namespace Thesis.Views
{
    /// <summary>
    /// Interaction logic for TestWindow.xaml
    /// </summary>
    public partial class TestWindow : Window
    {
        public TestWindow()
        {
            InitializeComponent();
            spreadsheet.Open("C:\\Users\\jacob\\Desktop\\FormatException.xlsx");
        }
    }
}
