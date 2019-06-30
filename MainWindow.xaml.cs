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
using Syncfusion.UI.Xaml.CellGrid;
using System.ComponentModel;
using System.Diagnostics;

namespace Thesis
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : MetroWindow
    {
        private string activeWorksheet;
        private static ObservableCollection<LogItem> log;
        private static ListView logView;

        public MainWindow()
        {
            InitializeComponent();
            log = new ObservableCollection<LogItem>();
            logListView.ItemsSource = log;
            logView = logListView;

            if (!string.IsNullOrEmpty(App.Settings.FilePath))
            {
                FileSelected();
                LoadSpreadsheet();
            }
        }

        public static void Log(LogItemType type, string message)
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                log.Add(new LogItem(type, message));
                logView.SelectedIndex = logView.Items.Count - 1;
                logView.ScrollIntoView(logView.SelectedItem);
            });
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
                Log(LogItemType.Info, "Selected file " + App.Settings.FilePath);
                FileSelected();
            }
        }

        private void LoadFileButton_Click(object sender, RoutedEventArgs e)
        {
            LoadSpreadsheet();
        }

        private void LoadSpreadsheet()
        {
            Log(LogItemType.Info, $"Loading {App.Settings.FilePath}...");
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
            if (((Syncfusion.XlsIO.Implementation.WorkbookImpl)args.Workbook).IsLoaded)
                Log(LogItemType.Success, "Successfully loaded file.");

            // disable editing
            spreadsheet.ActiveGrid.AllowEditing = false;
        }

        private void GenerateButton_Click(object sender, RoutedEventArgs e)
        {
            Log(LogItemType.Info, "Starting generation...");
            diagramLoading.IsActive = true;

            BackgroundWorker worker = new BackgroundWorker();
            worker.DoWork += worker_Generate;
            worker.RunWorkerCompleted += worker_GenerateCompleted;
            worker.RunWorkerAsync();    
        }

        private void worker_Generate(object sender, DoWorkEventArgs e)
        {
            activeWorksheet = spreadsheet.ActiveSheet.Name;
            var allCells = spreadsheet.ActiveSheet.Range[1, 1, spreadsheet.ActiveSheet.Rows.Length, spreadsheet.ActiveSheet.Columns.Length];
            var stopwatch = Stopwatch.StartNew();
            Graph graph = new Graph(allCells);
            Log(LogItemType.Info, $"Generating graph took {stopwatch.ElapsedMilliseconds} ms");

            Dispatcher.Invoke(() =>
            {
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

                stopwatch.Restart();
                Log(LogItemType.Info, "Layouting graph...");
                diagram.LayoutManager.Layout.UpdateLayout();
                Log(LogItemType.Info, $"Layouting graph took {stopwatch.ElapsedMilliseconds} ms");
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
            });                        
        }

        void worker_GenerateCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            spreadsheet.ActiveGrid.CellClick += (s, e1) =>
            {
                foreach (var node in ((DiagramCollection<NodeViewModel>)diagram.Nodes))
                {
                    node.IsSelected = ((Vertex)node.Content).CellIndex[0] == e1.RowIndex && ((Vertex)node.Content).CellIndex[1] == e1.ColumnIndex;
                }
            };
            Log(LogItemType.Success, "Finished graph generation.");
            diagramLoading.IsActive = false;
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
