using System;
using System.ComponentModel;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using Syncfusion.UI.Xaml.CellGrid.Helpers;
using Syncfusion.UI.Xaml.Diagram;
using Syncfusion.UI.Xaml.Spreadsheet.Helpers;
using Syncfusion.XlsIO.Implementation;
using Thesis.Models;
using Thesis.Models.VertexTypes;
using Thesis.ViewModels;
using GroupCollection = Syncfusion.UI.Xaml.Diagram.GroupCollection;

namespace Thesis.Views
{
    /// <summary>
    ///     Events for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        private void LoadFileButton_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Excel files|*.xls;*.xlsx;*.xlsm"
            };
            if (openFileDialog.ShowDialog().HasValue)
            {
                var path = openFileDialog.FileName;
                if (string.IsNullOrEmpty(path)) return;
                App.Settings.SelectedFile = path;
                App.Settings.Persist();
                Logger.Log(LogItemType.Info, "Selected file " + App.Settings.SelectedFile);
                LoadSpreadsheet();
            }
        }

        private async void MagicButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            magicButton.IsEnabled = false;
            magicButtonProgressRing.IsActive = true;

            generateGraphTab.IsSelected = true;
            GenerateGraph();
            await FilterGraph();

            generateCodeTab.IsSelected = true;
            await GenerateCode();
            await TestCode();

            magicButton.IsEnabled = true;
            magicButtonProgressRing.IsActive = false;
        }

        private void GenerateGraphButton_Click(object sender, RoutedEventArgs e)
        {
            if (_generator?.Graph != null && App.Settings.SelectedWorksheet == spreadsheet.ActiveSheet.Name)
            {
                var messageBoxResult = MessageBox
                    .Show("This will recreate the entire graph. Any custom labels you entered (except for classes) will be discarded. Are you sure?",
                        App.AppName,
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Question);
                if (messageBoxResult != MessageBoxResult.Yes)
                    return;
            }
            GenerateGraph();
        }

        private void GenerateGraph()
        {
            _generator = new Generator(this);

            if (_generator.GenerateGraph())
            {
                ProvideGraphFilteringOptions();
                EnableCodeGenerationOptions();
            }
            else
            {
                DisableGraphOptions();
            }
        }

        private async void FilterGraphButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            await FilterGraph();
        }

        private async Task FilterGraph()
        {
            filterGraphButton.IsEnabled = false;
            generateGraphProgressRing.IsActive = true;
            diagram.Opacity = 0.3f;

            await Task.Run(() => _generator.FilterGraph());

            generateGraphProgressRing.IsActive = false;
            filterGraphButton.IsEnabled = true;
            diagram.Opacity = 1;

            // scroll to top left
            (diagram.Info as IGraphInfo).BringIntoViewport(new Rect(new Size(0, 0)));
        }

        private async void GenerateCodeButton_Click(object sender, RoutedEventArgs e)
        {
            await GenerateCode();
        }

        private async Task GenerateCode()
        {
            generateCodeButton.IsEnabled = false;
            testCodeButton.IsEnabled = false;
            codeGeneratorProgressRing.IsActive = true;

            await _generator.GenerateCode();

            codeGeneratorProgressRing.IsActive = false;
            testCodeButton.IsEnabled = true;
            generateCodeButton.IsEnabled = true;
        }

        private async void TestCodeButton_Click(object sender, RoutedEventArgs e)
        {
            await TestCode();
        }

        private async Task TestCode()
        {
            testCodeButton.IsEnabled = false;
            logTab.IsSelected = true;
            codeGeneratorProgressRing.IsActive = true;

            await _generator.TestCode();

            codeGeneratorProgressRing.IsActive = false;
            testCodeButton.IsEnabled = true;
        }

        private void ClearLogButton_Click(object sender, RoutedEventArgs e)
        {
            Logger.Clear();
        }

        private void SelectAllButton_Click(object sender, RoutedEventArgs e)
        {
            _generator.SelectAllOutputFields();
        }

        private void UnselectAllButton_Click(object sender, RoutedEventArgs e)
        {
            _generator.UnselectAllOutputFields();
        }

        private void ShowTestFrameworkButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            codeTextBox.ScrollTo(0,0);
            codeTextBox.Text = Properties.Resources.CSharpTestingFramework;
        }
        
        public void SpreadsheetCellSelected(object sender, CurrentCellActivatedEventArgs e)
        {
            if (e.ActivationTrigger == ActivationTrigger.Program || _generator?.Graph == null) return;

            string sheetName = spreadsheet.ActiveSheet.Name;
            var vertex = _generator.Graph.GetVertexByAddress(sheetName, e.CurrentRowColumnIndex.RowIndex, e.CurrentRowColumnIndex.ColumnIndex);

            if (vertex == null || (vertex is CellVertex cellVertex 
                                   && cellVertex.NodeType == NodeType.None 
                                   && cellVertex.CellType == CellType.Unknown))
            {
                InitiateToolbox(null);
            }
            else
            {
                SelectVertexInDiagrams(vertex);
                SelectVertexInOutputListView(vertex);
                SelectVertexInCode(vertex);
                InitiateToolbox(vertex);
            }
        }

        public void DiagramItemClicked(object sender, DiagramEventArgs e)
        {
            DisableDiagramNodeTools();
            if (e.Item is NodeViewModel item)
            {
                if (item.Content is Vertex vertex)
                {
                    SelectVertexInSpreadsheet(vertex);
                    SelectVertexInOutputListView(vertex);
                    InitiateToolbox(vertex);
                }
            }
            else
            {
                InitiateToolbox(null);
            }
        }

        private void OutputFieldsListView_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (outputFieldsListView.Tag == null && e.AddedItems.Count == 1 && e.AddedItems[0] is Vertex vertex)
            {
                SelectVertexInSpreadsheet(vertex);
                SelectVertexInDiagrams(vertex);
                InitiateToolbox(vertex);
            }
        }

        private void CodeTextBox_TextChanged(object sender, EventArgs e)
        {
            _foldingStrategy.UpdateFoldings(_foldingManager, codeTextBox.Document);
        }

        private void CodeTextBoxSelectionChanged(object sender, EventArgs e)
        {
            if (_generator.Code?.VariableNameToVerticesDictionary == null) return;

            var selection = codeTextBox.SelectedText;
            if (_generator.Code.VariableNameToVerticesDictionary.TryGetValue(selection, out var vertexList))
            {
                if (vertexList.Count == 1)
                {
                    SelectVertexInSpreadsheet(vertexList[0]);
                    InitiateToolbox(vertexList[0]);
                }
                else
                {
                    SelectVerticesInSpreadsheet(vertexList);
                    InitiateToolbox(null);
                }
                codeTextBox.Focus();
            }
        }

        private void Spreadsheet_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "ActiveGrid" && spreadsheet.ActiveGrid != null)
            {
                spreadsheet.ActiveGrid.AllowEditing = false;
                spreadsheet.ActiveGrid.FillSeriesController.AllowFillSeries = false;
                spreadsheet.ActiveGrid.CellContextMenuOpening += ActiveGrid_CellContextMenuOpening;
                spreadsheet.ActiveGrid.CurrentCellActivated += SpreadsheetCellSelected;
            }
            else if (e.PropertyName == "ActiveSheet" && spreadsheet.ActiveSheet != null)
            {
                if (_generator != null && _generator.ActiveWorksheet == spreadsheet.ActiveSheet.Name)
                    ProvideGraphFilteringOptions();
                else
                    ProvideGraphGenerationOptions();
            }
        }

        private void Spreadsheet_WorkbookLoaded(object sender, WorkbookLoadedEventArgs args)
        {
            if (((WorkbookImpl)args.Workbook).IsLoaded)
            {
                Logger.Log(LogItemType.Success, "Loaded file.");

                if (!string.IsNullOrEmpty(App.Settings.SelectedWorksheet)
                    && spreadsheet.Workbook.Worksheets.Any(w => w.Name == App.Settings.SelectedWorksheet))
                {
                    Logger.Log(LogItemType.Info, $"Loading last selected worksheet {App.Settings.SelectedWorksheet}");
                    spreadsheet.SetActiveSheet(App.Settings.SelectedWorksheet);
                }

                generateGraphButton.IsEnabled = magicButton.IsEnabled = true;
                generateGraphTab.IsSelected = true;
            }
        }

        private void ActiveGrid_CellContextMenuOpening(object sender, CellContextMenuOpeningEventArgs e)
        {
            spreadsheet.ActiveGrid.CellContextMenu.Items.Clear();

            string sheetName = spreadsheet.ActiveSheet.Name;
            var vertex = _generator.Graph.GetVertexByAddress(sheetName, e.Cell.RowIndex, e.Cell.ColumnIndex);

            if (vertex is CellVertex cellVertex && cellVertex.NodeType == NodeType.OutputField)
            {
                var includeInGeneration = new MenuItem
                {
                    Header = "Include in generation"
                };
                includeInGeneration.Click += (s, e1) => cellVertex.Include = true;
                spreadsheet.ActiveGrid.CellContextMenu.Items.Add(includeInGeneration);
            }
        }

        public void DiagramAnnotationChanged(object sender, ChangeEventArgs<object, AnnotationChangedEventArgs> args)
        {
            // disable annotation editing
            args.Cancel = true;
        }
    }
}
