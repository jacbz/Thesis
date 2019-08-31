using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using Syncfusion.UI.Xaml.CellGrid.Helpers;
using Syncfusion.UI.Xaml.Diagram;
using Thesis.Models;
using Thesis.Models.CodeGeneration;
using Thesis.Models.VertexTypes;
using Thesis.Views;
using CSharpGenerator = Thesis.Models.CodeGeneration.CSharp.CSharpGenerator;

namespace Thesis.ViewModels
{
    public class Generator
    {
        private readonly MainWindow _window;

        public Graph Graph { get; set; }
        public ClassCollection ClassCollection { get; set; }
        public Code Code { get; set; }

        public bool HideConnections { get; set; }
        public ObservableCollection<CellVertex> OutputVertices { get; set; }
        public string ActiveWorksheet { get; set; }

        public Generator(MainWindow mainWindow)
        {
            _window = mainWindow;
        }

        // return value: success
        public bool GenerateGraph()
        {
            _window.generateGraphButton.IsEnabled = false;

            ActiveWorksheet = _window.spreadsheet.ActiveSheet.Name;
            
            App.Settings.SelectedWorksheet = ActiveWorksheet;
            App.Settings.Persist();

            var logItem = Logger.Log(LogItemType.Info, "Generate graph from spreadsheet cells to determine output fields...", true);

            var (rowCount, columnCount) = _window.GetSheetDimensions();
            if (rowCount == 0 || columnCount == 0)
            {
                Logger.Log(LogItemType.Error, "Selected worksheet is empty!");
                return false;
            }

            var allCells = _window.spreadsheet.ActiveSheet.Range[1, 1, rowCount, columnCount];

            Graph = Graph.FromSpreadsheet(
                ActiveWorksheet, 
                allCells, 
                _window.GetRangeFromCurrentWorksheet,
                _window.spreadsheet.Workbook.Names);

            Logger.Log(LogItemType.Success, "Graph generation successful.");
            ClassCollection = null;
            Code = null;

            _window.generateGraphButton.IsEnabled = true;
            logItem.AppendElapsedTime();

            OutputVertices = new ObservableCollection<CellVertex>(Graph.GetOutputFields());
            LoadPersistedOutputFields();
            _window.outputFieldsListView.ItemsSource = OutputVertices;

            LayoutGraph();

            return true;
        }

        // return value: success
        private void LoadPersistedOutputFields()
        {
            // load selected output fields from settings
            var worksheetData = App.Settings.CurrentWorksheetSettings;

            if (worksheetData?.SelectedOutputFields != null 
                && worksheetData.SelectedOutputFields.Count > 0 
                && OutputVertices.Count > 0 
                && !OutputVertices.Where(v => v.Include).Select(v => v.StringAddress).ToList()
                    .SequenceEqual(worksheetData.SelectedOutputFields))
            {
                Logger.Log(LogItemType.Info, "Loading selected output fields from persisted user settings...");
                UnselectAllOutputFields();
                foreach (var v in OutputVertices)
                {
                    if (worksheetData.SelectedOutputFields.Contains(v.StringAddress))
                    {
                        v.Include = true;
                    }
                }

                var firstIncludedOutputVertex = OutputVertices.FirstOrDefault(v => v.Include);
                if (firstIncludedOutputVertex != null)
                    _window.outputFieldsListView.ScrollIntoView(firstIncludedOutputVertex);
            }
        }

        /// <summary>
        /// Filters by selected output vertices, displays graph, and colors spreadsheet cells
        /// </summary>
        public void FilterGraph()
        {
            var includedVertices = OutputVertices.Where(v => v.Include).ToList();
            var includedVertexAddresses = includedVertices.Select(v => v.StringAddress).ToList();
            App.Settings.CurrentWorksheetSettings.SelectedOutputFields =
                includedVertexAddresses.Count != OutputVertices.Count 
                    ? includedVertexAddresses 
                    : null;
            App.Settings.Persist();

            Logger.Log(LogItemType.Info,
                $"Selected output fields: {string.Join(", ", includedVertexAddresses)}");

            Graph.PerformTransitiveFilter(includedVertices);

            Application.Current.Dispatcher.Invoke(LayoutGraph, DispatcherPriority.Background);

            Logger.Log(LogItemType.Success, "Graph generation complete.");
        }

        private void LayoutGraph()
        {
            var logItem = Logger.Log(LogItemType.Info, "Layouting graph...", true);

            _window.diagram.Nodes = new NodeCollection();
            _window.diagram.Connectors = new ConnectorCollection();

            // only display non-external vertices
            foreach (var vertex in Graph.Vertices.Where(v => !v.IsExternal))
            {
                ((NodeCollection)_window.diagram.Nodes).Add(
                    vertex is CellVertex cellVertex 
                        ? cellVertex.FormatCellVertex(Graph) 
                        : ((RangeVertex)vertex).FormatRangeVertexLarge(Graph));
            }

            foreach (var vertex in Graph.Vertices.GetCellVertices())
            foreach (var child in vertex.Children)
                ((ConnectorCollection)_window.diagram.Connectors).Add(vertex.FormatEdge(child, false));

            logItem.AppendElapsedTime();

            _window.ResetAndColorAllCells(Graph.AllVertices);
        }

        public void LayoutClasses()
        {
            var logItem = Logger.Log(LogItemType.Info, "Layouting classes...", true);

            _window.diagram2.Nodes = new NodeCollection();
            _window.diagram2.Connectors = new ConnectorCollection();
            try
            {
                _window.diagram2.Groups = new GroupCollection();
            }
            catch (System.Exception)
            {
                // error in Syncfusion SfDiagram, rarely triggered, but unknown cause
                Logger.Log(LogItemType.Error, "Error in diagram control!");
                return;
            }

            double nextPos = 0;
            foreach (var generatedClass in ClassCollection.Classes)
            {
                var (group, nextPosX) = generatedClass.FormatClass(nextPos);
                nextPos = nextPosX;
                ((GroupCollection)_window.diagram2.Groups).Add(group);
            }
            
            foreach (var vertex in Graph.Vertices.GetCellVertices())
            foreach (var child in vertex.Children)
            {
                if (HideConnections && vertex.Class != child.Class) continue;
                ((ConnectorCollection)_window.diagram2.Connectors).Add(vertex.FormatEdge(child));
            }

            logItem.AppendElapsedTime();
        }

        public void SelectAllOutputFields()
        {
            OutputVertices.ForEach(v => v.Include = true);
        }

        public void UnselectAllOutputFields()
        {
            OutputVertices.ForEach(v => v.Include = false);
        }

        public void GenerateClasses()
        {
            var logItem = Logger.Log(LogItemType.Info, "Generate classes for selected output fields...", true);

            var customClassNames = ClassCollection == null
                ? App.Settings.CurrentWorksheetSettings.CustomClassNames
                : ClassCollection.GetCustomClassNames();

            ClassCollection = customClassNames != null 
                ? ClassCollection.FromGraph(Graph, customClassNames)
                : ClassCollection.FromGraph(Graph);

            logItem.AppendElapsedTime();

            Application.Current.Dispatcher.Invoke(() =>
            {
                LayoutClasses();

                _window.ResetSpreadsheetColors();
                foreach (var generatedClass in ClassCollection.Classes)
                {
                    _window.ColorSpreadsheetCells(generatedClass.Vertices.Where(v => !v.IsExternal),
                        (vertex, style) =>
                        {
                            _window.StyleCellByColor(generatedClass.Color, style);
                        },
                        _window.StyleBorderByNodeType);
                }

                _window.ColorSpreadsheetExternalCells(Graph.ExternalVertices);
                _window.spreadsheet.ActiveGrid.InvalidateCells();
            }, DispatcherPriority.Background);

            Logger.Log(LogItemType.Success, $"Generated {ClassCollection.Classes.Count} classes.");
        }

        public async Task GenerateCode()
        {
            var logItem = Logger.Log(LogItemType.Info, $"Generating code...", true);

            _window.codeTextBox.Text = "";

            var addressToVertexDictionary = Graph.Vertices
                .GetCellVertices()
                .ToDictionary(v => (v.WorksheetName, v.StringAddress), v => v);

            // implement different languages here
            CodeGenerator codeGenerator;
            switch (_window.languageComboBox.SelectedIndex)
            {
                default:
                    codeGenerator = new CSharpGenerator(
                        ClassCollection,
                        addressToVertexDictionary,
                        Graph.RangeDictionary,
                        Graph.NameDictionary);
                    break;
            }

            Code = await Code.GenerateWith(codeGenerator);
            logItem.AppendElapsedTime();

            _window.codeTextBox.Text = Code.SourceCode;

            Logger.Log(LogItemType.Success, $"Successfully generated code.");
        }

        public async Task TestCode()
        {
            Logger.Log(LogItemType.Info, $"Testing code...");

            if (Code != null)
            {
                var report = await Task.Run(() => Code.GenerateTestReportAsync());

                _window.codeTextBox.Text = report.Code;
                Logger.Log(report.NullCount == 0
                           && report.TypeMismatchCount == 0
                           && report.ValueMismatchCount == 0 
                           && report.ErrorCount == 0
                        ? LogItemType.Success
                        : report.PassCount == 0
                            ? LogItemType.Error
                            : LogItemType.Warning,
                    report.ToString());
            }
            else
            {
                Logger.Log(LogItemType.Error, $"Code generator was not initialized.");
            }
        }
    }
}