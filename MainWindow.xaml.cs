using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Syncfusion.UI.Xaml.Diagram;
using MahApps.Metro.Controls;
using Microsoft.Win32;
using System.Data;
using Syncfusion.XlsIO;
using Syncfusion.UI.Xaml.Diagram.Layout;
using System.Collections.ObjectModel;
using Syncfusion.UI.Xaml.Diagram.Controls;

namespace Thesis
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : MetroWindow
    {
        private string activeWorksheet;
        public MainWindow()
        {
            InitializeComponent();
            if (!string.IsNullOrEmpty(App.Settings.FilePath))
            {
                FileSelected();
                LoadSpreadsheet();
            }
        }
        public void FileSelected()
        {
            loadFileButton.IsEnabled = true;
            pathLabel.Content = App.Settings.FilePath;
            pathLabel.FontStyle = FontStyles.Normal;
        }

        private void ChooseFileButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel files|*.xls;*.xlsx;*.xlsm";
            if (openFileDialog.ShowDialog() == true)
            {
                string path = openFileDialog.FileName;
                App.Settings.FilePath = path;
                App.Settings.Save();
                FileSelected();
            }
        }

        private void LoadFileButton_Click(object sender, RoutedEventArgs e)
        {
            LoadSpreadsheet();
        }

        private void LoadSpreadsheet()
        {
            spreadsheet.Open(App.Settings.FilePath);
            spreadsheet.Opacity = 100;
        }

        private void Spreadsheet_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "ActiveSheet")
            {
                worksheetLabel.Content = $"Selected {spreadsheet.ActiveSheet.Name}: " +
                    $"{spreadsheet.ActiveSheet.Columns.Length} columns, {spreadsheet.ActiveSheet.Rows.Length} rows ";
                generateButton.IsEnabled = true;
            }
        }
        private void Spreadsheet_WorkbookLoaded(object sender, Syncfusion.UI.Xaml.Spreadsheet.Helpers.WorkbookLoadedEventArgs args)
        {
            // disable editing
            spreadsheet.ActiveGrid.AllowEditing = false;
        }

        private void GenerateButton_Click(object sender, RoutedEventArgs e)
        {
            spreadsheet.ActiveGrid.CellClick += (s, e1) =>
            {
                foreach(var node in ((DiagramCollection<NodeViewModel>)diagram.Nodes))
                {
                    node.IsSelected = ((Vertex)node.Content).CellIndex[0] == e1.RowIndex && ((Vertex)node.Content).CellIndex[1] == e1.ColumnIndex;
                }
            };

            activeWorksheet = spreadsheet.ActiveSheet.Name;
            var allCells = spreadsheet.ActiveSheet.Range[1, 1, spreadsheet.ActiveSheet.Rows.Length, spreadsheet.ActiveSheet.Columns.Length];
            Graph graph = new Graph(allCells);

            //Initialize Nodes and Connectors
            diagram.Nodes = new DiagramCollection<NodeViewModel>();
            diagram.Connectors = new DiagramCollection<CustomConnectorViewModel>();

            var layoutSettings = new DirectedTreeLayout
            {
                Type = LayoutType.Hierarchical,
                Orientation = TreeOrientation.TopToBottom,
                HorizontalSpacing = 80,
                VerticalSpacing = 80,
                Margin = new Thickness()
            };

            diagram.LayoutManager = new LayoutManager { Layout = layoutSettings };

            var settings = new DataSourceSettings();
            settings.ParentId = "Parents";
            settings.Id = "Address";
            settings.DataSource = graph.Vertices;

            diagram.DataSourceSettings = settings;
            diagram.Tool = Tool.ZoomPan | Tool.MultipleSelect;

            diagram.LayoutManager.Layout.UpdateLayout();
            ((IGraphInfo)diagram.Info).ItemTappedEvent += (s, args) =>
            {
                if (args.Item is NodeViewModel)
                {
                    var item = (NodeViewModel)args.Item;
                    if (item.Content is Vertex)
                    {
                        spreadsheet.SetActiveSheet(activeWorksheet);
                        var vertex = (Vertex)item.Content;
                        spreadsheet.ActiveGrid.CurrentCell.MoveCurrentCell(vertex.CellIndex[0], vertex.CellIndex[1]);
                    }
                }
            };
        }

        public class CustomConnectorViewModel : ConnectorViewModel
        {
            public CustomConnectorViewModel()
                : base()
            {
                this.CornerRadius = 10;
            }
        }

    }
}
