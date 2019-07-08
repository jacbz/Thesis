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
using System.Drawing;
using Thesis.ViewModels;
using Syncfusion.UI.Xaml.CellGrid.Helpers;

namespace Thesis
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : MetroWindow
    {
        private string activeWorksheet;
        private Generator generator;

        public MainWindow()
        {
            InitializeComponent();
            Logger.Instantiate(logControl.logListView);

            if (!string.IsNullOrEmpty(App.Settings.FilePath))
            {
                LoadSpreadsheet();
            }
        }      

        private void LoadFileButton_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel files|*.xls;*.xlsx;*.xlsm";
            if (openFileDialog.ShowDialog() == true)
            {
                string path = openFileDialog.FileName;
                App.Settings.FilePath = path;
                App.Settings.Save();
                Logger.Log(LogItemType.Info, "Selected file " + App.Settings.FilePath);
                LoadSpreadsheet();
            }
        }

        private void LoadSpreadsheet()
        {
            loadFileButton.IsEnabled = true;
            pathLabel.Content = pathLabel.ToolTip = App.Settings.FilePath;
            pathLabel.FontStyle = FontStyles.Normal;

            Logger.Log(LogItemType.Info, $"Loading {App.Settings.FilePath}");
            spreadsheet.Open(App.Settings.FilePath);
            spreadsheet.Opacity = 100;

            DisableTools();
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
                Logger.Log(LogItemType.Success, "Successfully loaded file.");

            // disable editing
            spreadsheet.ActiveGrid.AllowEditing = false;
        }

        private void GenerateButton_Click(object sender, RoutedEventArgs e)
        {
            Logger.Log(LogItemType.Info, "Starting generation...");
            activeWorksheet = spreadsheet.ActiveSheet.Name;
            diagramLoading.IsActive = true;
            generator = new Generator(this);

            BackgroundWorker worker = new BackgroundWorker();
            worker.DoWork += generator.worker_Generate;
            worker.RunWorkerCompleted += generator.worker_GenerateCompleted;
            worker.RunWorkerAsync();    
        }

        private void DisableTools()
        {
            generator = new Generator(this);
            toolboxTab.IsEnabled = false;
            logTab.IsSelected = true;
            optionsTabControl.IsEnabled = false;
            optionsTabControl.Opacity = 0.5f;
            diagram.Nodes = new NodeCollection();
            diagram.Connectors = new ConnectorCollection();
        }

        public void EnableTools()
        {
            toolboxTab.IsEnabled = true;
            toolboxTab.IsSelected = true;
            optionsTabControl.IsEnabled = true;
            optionsTabControl.Opacity = 1f;
        }

        public void DiagramItemClicked(object sender, DiagramEventArgs e)
        {
            if (generator.IsFinished && e.Item is NodeViewModel)
            {
                var item = (NodeViewModel)e.Item;
                if (item.Content is Vertex)
                {
                    spreadsheet.SetActiveSheet(activeWorksheet);
                    var vertex = (Vertex)item.Content;
                    SpreadsheetSelectVertex(vertex);
                    ListViewSelectVertex(vertex);
                }
            }
        }

        private void SpreadsheetSelectVertex(Vertex vertex)
        {
            spreadsheet.ActiveGrid.CurrentCell.MoveCurrentCell(vertex.CellIndex[0], vertex.CellIndex[1]);
        }

        public void SpreadsheetCellClicked(object sender, GridCellClickEventArgs e)
        {
            if (generator.IsFinished)
            {
                Vertex vertex = generator.Graph.Vertices.FirstOrDefault(v => v.CellIndex[0] == e.RowIndex && v.CellIndex[1] == e.ColumnIndex);

                if (vertex != null)
                {
                    DiagramSelectVertex(vertex);
                    ListViewSelectVertex(vertex);
                }
                else
                {
                    Logger.Log(LogItemType.Error, $"Error selecting cell [{e.RowIndex},{e.ColumnIndex}]");
                }
            }
        }

        private void DiagramSelectVertex(Vertex vertex)
        {
            foreach (var node in ((DiagramCollection<NodeViewModel>)diagram.Nodes))
            {
                node.IsSelected = ((Vertex)node.Content).CellIndex[0] == vertex.CellIndex[0]
                    && ((Vertex)node.Content).CellIndex[1] == vertex.CellIndex[1];
            }
        }

        private void ListViewSelectVertex(Vertex vertex)
        {
            if (vertex.IsOutputField)
            {
                outputFieldsListView.SelectedItem = vertex;
                outputFieldsListView.ScrollIntoView(vertex);
            }
        }

        private void OutputFieldsListView_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (e.AddedItems.Count == 1 && e.AddedItems[0] is Vertex vertex)
            {
                SpreadsheetSelectVertex(vertex);
                DiagramSelectVertex(vertex);
            }
        }

        private void FilterOutputFieldsButton_Click(object sender, RoutedEventArgs e)
        {
            generator.FilterForSelectedOutputFields();
        }

        private void UnselectAllButton_Click(object sender, RoutedEventArgs e)
        {
            generator.UnselectAllOutputFields();
        }
    }
}
