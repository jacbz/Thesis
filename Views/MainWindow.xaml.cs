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
using Syncfusion.XlsIO.Implementation;
using Thesis.Models;

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
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Excel files|*.xls;*.xlsx;*.xlsm"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                string path = openFileDialog.FileName;
                App.Settings.FilePath = path;
                App.Settings.ResetFileSpecificSettings();
                App.Settings.Save();
                Logger.Log(LogItemType.Info, "Selected file " + App.Settings.FilePath);
                LoadSpreadsheet();
            }
        }

        private void LoadSpreadsheet()
        {
            pathLabel.Content = pathLabel.ToolTip = App.Settings.FilePath;
            pathLabel.FontStyle = FontStyles.Normal;

            Logger.Log(LogItemType.Info, $"Loading {App.Settings.FilePath}");
            spreadsheet.Open(App.Settings.FilePath);
            spreadsheet.Opacity = 100;

            generateGraphButton.IsEnabled = selectAllButton.IsEnabled = unselectAllButton.IsEnabled = true;

            DisableGraphOptions();
        }
        
        private void Spreadsheet_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "ActiveSheet")
            {
            }
        }
        private void Spreadsheet_WorkbookLoaded(object sender, Syncfusion.UI.Xaml.Spreadsheet.Helpers.WorkbookLoadedEventArgs args)
        {
            if (((WorkbookImpl)args.Workbook).IsLoaded)
            {
                Logger.Log(LogItemType.Success, "Successfully loaded file.");
                generator = new Generator(this);
                activeWorksheet = spreadsheet.ActiveSheet.Name;
                generator.GenerateGraph();
            }

            // disable editing
            spreadsheet.ActiveGrid.AllowEditing = false;
            spreadsheet.ActiveGrid.CellContextMenuOpening += ActiveGrid_CellContextMenuOpening;
        }

        void ActiveGrid_CellContextMenuOpening(object sender, CellContextMenuOpeningEventArgs e)
        {
            spreadsheet.ActiveGrid.CellContextMenu.Items.Clear();

            Vertex vertex = generator.Graph.Vertices
                .FirstOrDefault(v => v.CellIndex[0] == e.Cell.RowIndex && v.CellIndex[1] == e.Cell.ColumnIndex);

            if (vertex != null && vertex.NodeType == NodeType.OutputField)
            {
                MenuItem includeInGeneration = new MenuItem
                {
                    Header = "Include in generation"
                };
                includeInGeneration.Click += (s, e1) => vertex.Include = true;
                spreadsheet.ActiveGrid.CellContextMenu.Items.Add(includeInGeneration);
            }
        }

        private void GenerateGraphButton_Click(object sender, RoutedEventArgs e)
        {
            generator.LoadDataIntoGraphAndSpreadsheet();
        }

        public void EnableGraphOptions()
        {
            toolboxTab.IsEnabled = true;
            //toolboxTab.IsSelected = true;
            generateClassesButton.IsEnabled = hideConnectionsCheckbox.IsEnabled = true;
        }

        private void DisableGraphOptions()
        {
            generator = new Generator(this);
            toolboxTab.IsEnabled = false;
            logTab.IsSelected = true;
            diagram.Nodes = new NodeCollection();
            diagram.Connectors = new ConnectorCollection();
            generateClassesButton.IsEnabled = hideConnectionsCheckbox.IsEnabled = false;
            DisableClassOptions();
        }

        private void EnableClassOptions()
        {
        }

        private void DisableClassOptions()
        {
        }

        public void DiagramItemClicked(object sender, DiagramEventArgs e)
        {
            DisableDiagramNodeTools();
            if (e.Item is NodeViewModel item && item.Content is Vertex vertex)
            {
                spreadsheet.SetActiveSheet(activeWorksheet);
                SpreadsheetSelectVertex(vertex);
                OutputListViewSelectVertex(vertex);
                InitiateToolbox(vertex);
            }
            else
            {
                InitiateToolbox(null);
            }
        }

        private void DisableDiagramNodeTools()
        {
            DisableDiagramNodeTools(diagram);
            DisableDiagramNodeTools(diagram2);
        }

        private void DisableDiagramNodeTools(SfDiagram diagram)
        {
            // disable remove, rotate buttons etc. on click
            var selecteditem = diagram.SelectedItems as SelectorViewModel;
            (selecteditem.Commands as QuickCommandCollection).Clear();
            selecteditem.SelectorConstraints = selecteditem.SelectorConstraints.Remove(SelectorConstraints.Rotator);
        }

        private void SpreadsheetSelectVertex(Vertex vertex)
        {
            spreadsheet.ActiveGrid.CurrentCell.MoveCurrentCell(vertex.CellIndex[0], vertex.CellIndex[1]);
        }

        public void SpreadsheetCellSelected(object sender, CurrentCellActivatedEventArgs e)
        {
            if (e.ActivationTrigger == ActivationTrigger.Program) return;
            Vertex vertex = generator.Graph.Vertices
                .FirstOrDefault(v => v.CellIndex[0] == e.CurrentRowColumnIndex.RowIndex && v.CellIndex[1] == e.CurrentRowColumnIndex.ColumnIndex);
            if (vertex != null)
            {
                DiagramSelectVertex(vertex);
                OutputListViewSelectVertex(vertex);
                InitiateToolbox(vertex);
            }
            else
            {
                InitiateToolbox(null);
            }
        }

        private void DiagramSelectVertex(Vertex vertex)
        {
            if (diagram.Nodes == null) return;
            foreach (var node in (DiagramCollection<NodeViewModel>)diagram.Nodes)
            {
                if (node.Content is Vertex nodeVertex)
                {
                    if (nodeVertex.CellIndex[0] == vertex.CellIndex[0] && nodeVertex.CellIndex[1] == vertex.CellIndex[1])
                    {
                        node.IsSelected = true;
                        (diagram.Info as IGraphInfo).BringIntoCenter((node.Info as INodeInfo).Bounds);
                        DisableDiagramNodeTools(diagram);
                    }
                    else
                    {
                        node.IsSelected = false;
                    }
                }
            }

            if (diagram2.Groups == null) return;
            foreach (var group in (DiagramCollection<GroupViewModel>)diagram2.Groups)
            {
                foreach (var node in (ObservableCollection<NodeViewModel>)group.Nodes)
                {
                    if (node.Content is Vertex nodeVertex)
                    {
                        if (nodeVertex.CellIndex[0] == vertex.CellIndex[0] && nodeVertex.CellIndex[1] == vertex.CellIndex[1])
                        {
                            node.IsSelected = true;
                            (diagram2.Info as IGraphInfo).BringIntoCenter((node.Info as INodeInfo).Bounds);
                            DisableDiagramNodeTools(diagram2);
                        }
                        else
                        {
                            node.IsSelected = false;
                        }
                    }
                }
            }
        }

        private void OutputListViewSelectVertex(Vertex vertex)
        {
            if (vertex.NodeType == NodeType.OutputField)
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
                InitiateToolbox(vertex);
            }
        }

        private void InitiateToolbox(Vertex vertex)
        {
            if (vertex == null)
            {
                toolboxContent.Visibility = Visibility.Collapsed;
            } else
            {
                toolboxContent.Visibility = Visibility.Visible;
                toolboxTab.IsSelected = true;
                DataContext = vertex;
            }
        }
        
        private void SelectAllButton_Click(object sender, RoutedEventArgs e)
        {
            generator.SelectAllOutputFields();
        }

        private void UnselectAllButton_Click(object sender, RoutedEventArgs e)
        {
            generator.UnselectAllOutputFields();
        }

        private void GenerateClassesButton_Click(object sender, RoutedEventArgs e)
        {
            generator.HideConnections = hideConnectionsCheckbox.IsChecked.Value;
            generator.GenerateClasses();
            EnableClassOptions();
        }
    }
}
