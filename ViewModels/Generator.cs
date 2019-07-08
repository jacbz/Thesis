using Syncfusion.UI.Xaml.CellGrid.Helpers;
using Syncfusion.UI.Xaml.Diagram;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Threading;

namespace Thesis.ViewModels
{
    class Generator
    {
        private MainWindow window;
        public Graph Graph { get; set; }
        public bool IsFinished { get; set; }

        public ObservableCollection<Vertex> OutputVertices { get; set; }

        public Generator(MainWindow mainWindow)
        {
            this.window = mainWindow;
            this.IsFinished = false;
        }

        public void worker_Generate(object sender, DoWorkEventArgs e)
        {
            var allCells = window.spreadsheet.ActiveSheet.Range[1, 1, window.spreadsheet.ActiveSheet.Rows.Length,
                window.spreadsheet.ActiveSheet.Columns.Length];
            var stopwatch = Stopwatch.StartNew();
            Graph = new Graph(allCells);
            Logger.Log(LogItemType.Info, $"Generating graph took {stopwatch.ElapsedMilliseconds} ms");
        }

        public void worker_GenerateCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            LayoutGraph();

            window.diagram.Tool = Tool.ZoomPan | Tool.MultipleSelect;
            ((IGraphInfo)window.diagram.Info).ItemTappedEvent += (s, e1) => window.DiagramItemClicked(s, e1);
            window.spreadsheet.ActiveGrid.CellClick += (s, e1) => window.SpreadsheetCellClicked(s, e1);

            Logger.Log(LogItemType.Info, "Finished graph generation.");
            window.diagramLoading.IsActive = false;

            ColorSpreadsheetCells();

            FormatOptionsTab();

            Logger.Log(LogItemType.Success, "Done.");
            IsFinished = true;

            window.EnableTools();
        }

        private void LayoutGraph()
        {
            Logger.Log(LogItemType.Info, "Layouting graph...");
            var stopwatch = Stopwatch.StartNew();

            window.diagram.Nodes = new NodeCollection();
            window.diagram.Connectors = new ConnectorCollection();
            foreach (Vertex vertex in Graph.Vertices)
            {
                (window.diagram.Nodes as NodeCollection).Add(vertex.FormatNode(Graph));
            }
            foreach (Vertex vertex in Graph.Vertices)
            {
                foreach (Vertex child in vertex.Children)
                {
                    (window.diagram.Connectors as ConnectorCollection).Add(vertex.FormatConnector(child));
                }
            }

            Logger.Log(LogItemType.Info, $"Layouting graph took {stopwatch.ElapsedMilliseconds} ms");
        }

        private void ColorSpreadsheetCells()
        {
            Logger.Log(LogItemType.Info, "Coloring cells...");
            var allCells = window.spreadsheet.ActiveSheet.Range[1, 1, window.spreadsheet.ActiveSheet.Rows.Length, window.spreadsheet.ActiveSheet.Columns.Length];
            allCells.CellStyle.ColorIndex = ExcelKnownColors.White;

            foreach (Vertex v in Graph.Vertices)
            {
                IRange range = window.spreadsheet.ActiveSheet.Range[v.Address];
                range.CellStyle.Color = ColorTranslator.FromHtml(v.GetColor());
                range.CellStyle.Font.Color = ExcelKnownColors.White;
                window.spreadsheet.ActiveGrid.InvalidateCell(range.Row, range.Column);
            }

            window.spreadsheet.ActiveGrid.InvalidateCells();
        }

        private void FormatOptionsTab()
        {
            OutputVertices = new ObservableCollection<Vertex>(Graph.GetOutputFields());
            window.outputFieldsListView.ItemsSource = OutputVertices;
        }

        public void FilterForSelectedOutputFields()
        {
            var includedVertices = OutputVertices.Where(v => v.Include).ToList();
            Logger.Log(LogItemType.Info, $"Filtering for selected output fields {string.Join(", ", includedVertices.Select(v => v.Address)) }");
            Graph.TransitiveFilter(includedVertices);
            LayoutGraph();
            ColorSpreadsheetCells();
            Logger.Log(LogItemType.Success, "Done.");
        }

        public void UnselectAllOutputFields()
        {
            foreach(Vertex vertex in OutputVertices)
            {
                vertex.Include = false;
            }
        }
    }
}
