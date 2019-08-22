using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Animation;
using Microsoft.Win32;
using Syncfusion.UI.Xaml.CellGrid.Helpers;
using Syncfusion.UI.Xaml.Diagram;
using Syncfusion.UI.Xaml.Spreadsheet.Helpers;
using Syncfusion.XlsIO.Implementation;
using Thesis.Models;
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

        private void Spreadsheet_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "ActiveGrid" && spreadsheet.ActiveGrid != null)
            {
                spreadsheet.ActiveGrid.AllowEditing = false;
                spreadsheet.ActiveGrid.FillSeriesController.AllowFillSeries = false;
                spreadsheet.ActiveGrid.CellContextMenuOpening += ActiveGrid_CellContextMenuOpening;
                spreadsheet.ActiveGrid.CurrentCellActivated += SpreadsheetCellSelected;
            }
        }

        private void Spreadsheet_WorkbookLoaded(object sender, WorkbookLoadedEventArgs args)
        {
            if (((WorkbookImpl)args.Workbook).IsLoaded)
            {
                Logger.Log(LogItemType.Success, "Successfully loaded file.");

                if (!string.IsNullOrEmpty(App.Settings.SelectedWorksheet)
                    && spreadsheet.Workbook.Worksheets.Any(w => w.Name == App.Settings.SelectedWorksheet))
                {
                    Logger.Log(LogItemType.Info, $"Loading last selected worksheet {App.Settings.SelectedWorksheet}");
                    spreadsheet.SetActiveSheet(App.Settings.SelectedWorksheet);
                }

                _generator = new Generator(this);
                _generator.GenerateGraph();
                Logger.Log(LogItemType.Success, "Successfully generated graph.");

                this.generateGraphButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                this.generateClassesButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                this.generateCodeButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                generateCodeTab.IsSelected = true;
            }
        }

        private void ActiveGrid_CellContextMenuOpening(object sender, CellContextMenuOpeningEventArgs e)
        {
            spreadsheet.ActiveGrid.CellContextMenu.Items.Clear();

            var vertex = _generator.GetVertexByAddress(e.Cell.RowIndex, e.Cell.ColumnIndex);

            if (vertex != null && vertex.NodeType == NodeType.OutputField)
            {
                var includeInGeneration = new MenuItem
                {
                    Header = "Include in generation"
                };
                includeInGeneration.Click += (s, e1) => vertex.Include = true;
                spreadsheet.ActiveGrid.CellContextMenu.Items.Add(includeInGeneration);
            }
        }

        private void ClearLogButton_Click(object sender, RoutedEventArgs e)
        {
            Logger.Clear();
        }

        private void GenerateGraphButton_Click(object sender, RoutedEventArgs e)
        {
            _generator.FilterAndDisplayGraphIntoUi();

            // scroll to top left
            (diagram.Info as IGraphInfo).BringIntoViewport(new Rect(new Size(0, 0)));
        }

        private void SelectAllButton_Click(object sender, RoutedEventArgs e)
        {
            _generator.SelectAllOutputFields();
        }

        private void UnselectAllButton_Click(object sender, RoutedEventArgs e)
        {
            _generator.UnselectAllOutputFields();
        }

        private void GenerateClassesButton_Click(object sender, RoutedEventArgs e)
        {
            // unselect all - otherwise sometimes NullReferenceException is triggered due to a bug in SfDiagram group layouting
            if (diagram2.Groups != null)
                foreach (var group in (DiagramCollection<GroupViewModel>)diagram2.Groups)
                foreach (var node in (ObservableCollection<NodeViewModel>)group.Nodes)
                    if (node.Content is Vertex)
                    {
                        node.IsSelected = false;
                    }

            _generator.HideConnections = hideConnectionsCheckbox.IsChecked.Value;
            _generator.GenerateClasses();
            EnableCodeGenerationOptions();

            // scroll to top left
            (diagram2.Info as IGraphInfo).BringIntoViewport(new Rect(new Size(0, 0)));
        }

        private async void GenerateCodeButton_Click(object sender, RoutedEventArgs e)
        {
            codeGeneratorProgressRing.IsActive = true;

            await _generator.GenerateCode();

            codeGeneratorProgressRing.IsActive = false;
            testCodeButton.IsEnabled = true;
        }

        private async void TestCodeButton_Click(object sender, RoutedEventArgs e)
        {
            logTab.IsSelected = true;
            codeGeneratorProgressRing.IsActive = true;

            await _generator.TestCode();

            codeGeneratorProgressRing.IsActive = false;
        }

        public void DiagramAnnotationChanged(object sender, ChangeEventArgs<object, AnnotationChangedEventArgs> args)
        {
            // disable annotation editing
            args.Cancel = true;
        }

        public void SpreadsheetCellSelected(object sender, CurrentCellActivatedEventArgs e)
        {
            if (e.ActivationTrigger == ActivationTrigger.Program) return;
            var vertex = _generator.GetVertexByAddress(e.CurrentRowColumnIndex.RowIndex, e.CurrentRowColumnIndex.ColumnIndex);
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
            if (e.Item is NodeViewModel item && item.Content is Vertex vertex)
            {
                SelectVertexInSpreadsheet(vertex);
                SelectVertexInOutputListView(vertex);
                InitiateToolbox(vertex);
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

        private void CodeTextBox_TextChanged(object sender, System.EventArgs e)
        {
            _foldingStrategy.UpdateFoldings(_foldingManager, codeTextBox.Document);
        }

        private void CodeTextBoxSelectionChanged(object sender, EventArgs e)
        {
            if (_generator.CodeGenerator?.VariableNameToVertexDictionary == null) return;

            var selection = codeTextBox.SelectedText;
            if (_generator.CodeGenerator.VariableNameToVertexDictionary.TryGetValue(selection, out var vertex))
            {
                SelectVertexInSpreadsheet(vertex);
                InitiateToolbox(vertex);
            }
        }
    }
}
