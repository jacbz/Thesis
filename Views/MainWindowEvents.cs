using System;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using Microsoft.Win32;
using Syncfusion.UI.Xaml.CellGrid.Helpers;
using Syncfusion.UI.Xaml.Diagram;
using Syncfusion.UI.Xaml.Spreadsheet.Helpers;
using Syncfusion.XlsIO.Implementation;
using Thesis.Models;
using Thesis.Models.VertexTypes;
using Thesis.ViewModels;

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
                App.Settings.FilePath = path;
                App.Settings.ResetWorkbookSpecificSettings();
                App.Settings.Persist();
                Logger.Log(LogItemType.Info, "Selected file " + App.Settings.FilePath);
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

            generateClassesTab.IsSelected = true;
            await GenerateClasses();

            generateCodeTab.IsSelected = true;
            await GenerateCode();
            await TestCode();

            magicButton.IsEnabled = true;
            magicButtonProgressRing.IsActive = false;
        }

        private void GenerateGraphButton_Click(object sender, RoutedEventArgs e)
        {
            GenerateGraph();
        }

        private void GenerateGraph()
        {
            _generator = new Generator(this);

            if (_generator.GenerateGraph())
            {
                ProvideGraphFilteringOptions();
                EnableClassGenerationOptions();
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

        private async void GenerateClassesButton_Click(object sender, RoutedEventArgs e)
        {
            await GenerateClasses();
        }
        
        private async Task GenerateClasses()
        {
            // unselect all - otherwise sometimes NullReferenceException is triggered due to a bug in SfDiagram group layouting
            if (diagram2.Groups != null)
                foreach (var group in (GroupCollection)diagram2.Groups)
                {
                    foreach (var node in (NodeCollection)group.Nodes)
                        if (node.Content is Vertex)
                            node.IsSelected = false;
                    group.IsSelected = false;
                }

            _generator.HideConnections = hideConnectionsCheckbox.IsChecked.Value;

            generateClassesButton.IsEnabled = false;
            generateClassesProgressRing.IsActive = true;
            diagram2.Opacity = 0.3f;

            await Task.Run(() => _generator.GenerateClasses());

            generateClassesProgressRing.IsActive = false;
            generateClassesButton.IsEnabled = true;
            diagram2.Opacity = 1;

            EnableCodeGenerationOptions();

            // scroll to top left
            (diagram2.Info as IGraphInfo).BringIntoViewport(new Rect(new Size(0, 0)));
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
            codeTextBox.Text = Properties.Resources.CSharpTestingFramework;
        }
        
        public void SpreadsheetCellSelected(object sender, CurrentCellActivatedEventArgs e)
        {
            if (e.ActivationTrigger == ActivationTrigger.Program || _generator.Graph == null) return;

            string sheetName = spreadsheet.ActiveSheet.Name;
            var vertex = _generator.Graph.GetVertexByAddress(sheetName, e.CurrentRowColumnIndex.RowIndex, e.CurrentRowColumnIndex.ColumnIndex);
            if (vertex != null)
            {
                SelectVertexInDiagrams(vertex);
                SelectVertexInOutputListView(vertex);
                InitiateToolbox(vertex);
            }
            else
            {
                InitiateToolbox(null);
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
                else if (item.Content is Class @class)
                {
                    SelectClassVerticesInSpreadsheet(@class);
                    InitiateToolbox(@class);
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
            if (_generator.Code?.VariableNameToVertexDictionary == null) return;

            var selection = codeTextBox.SelectedText;

            // check if user selected both class and variable name (e.g. Global.A1)
            var match = _generator.Code.VariableNameToVertexDictionary
                .FirstOrDefault(kvp => selection == kvp.Key.className + "." + kvp.Key.variableName);
            // check if user selected only variable name (e.g. A1)
            if (match.Value == null)
                match = _generator.Code.VariableNameToVertexDictionary
                    .FirstOrDefault(kvp => selection == kvp.Key.variableName);

            if (match.Value != null)
            {
                var vertex = match.Value;
                SelectVertexInSpreadsheet(vertex);
                InitiateToolbox(vertex);
                // restore focus to textbox so user can copy selected text
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
