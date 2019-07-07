using Syncfusion.UI.Xaml.CellGrid.Helpers;
using Syncfusion.UI.Xaml.Diagram;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
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
        private Graph graph;
        public bool isFinished = false;

        public Generator(MainWindow mainWindow)
        {
            this.window = mainWindow;
        }

        public void worker_Generate(object sender, DoWorkEventArgs e)
        {
            var allCells = window.spreadsheet.ActiveSheet.Range[1, 1, window.spreadsheet.ActiveSheet.Rows.Length,
                window.spreadsheet.ActiveSheet.Columns.Length];
            var stopwatch = Stopwatch.StartNew();
            graph = new Graph(allCells);
            Logger.Log(LogItemType.Info, $"Generating graph took {stopwatch.ElapsedMilliseconds} ms");
        }

        public void worker_GenerateCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Logger.Log(LogItemType.Info, "Layouting graph...");
            var stopwatch = Stopwatch.StartNew();
            window.diagram.Nodes = new NodeCollection();
            window.diagram.Connectors = new ConnectorCollection();

            foreach (Vertex vertex in graph.Vertices)
            {
                (window.diagram.Nodes as NodeCollection).Add(vertex.FormatNode());
            }

            foreach (Edge edge in graph.Edges)
            {
                (window.diagram.Connectors as ConnectorCollection).Add(edge.FormatConnector());
            }

            Logger.Log(LogItemType.Info, $"Layouting graph took {stopwatch.ElapsedMilliseconds} ms");

            window.diagram.Tool = Tool.ZoomPan | Tool.MultipleSelect;
            ((IGraphInfo)window.diagram.Info).ItemTappedEvent += (s, e1) => window.DiagramItemClicked(s, e1);
            window.spreadsheet.ActiveGrid.CellClick += (s, e1) => window.SpreadsheetCellClicked(s, e1);

            Logger.Log(LogItemType.Info, "Finished graph generation.");
            window.diagramLoading.IsActive = false;

            Logger.Log(LogItemType.Info, "Coloring cells...");
            var allCells = window.spreadsheet.ActiveSheet.Range[1, 1, window.spreadsheet.ActiveSheet.Rows.Length, window.spreadsheet.ActiveSheet.Columns.Length];
            allCells.CellStyle.ColorIndex = ExcelKnownColors.White;

            foreach (Vertex v in graph.Vertices)
            {
                if (v.Address == "root") continue;
                IRange range = window.spreadsheet.ActiveSheet.Range[v.Address];
                range.CellStyle.Color = ColorTranslator.FromHtml(v.GetColor());
                range.CellStyle.Font.Color = ExcelKnownColors.White;
                window.spreadsheet.ActiveGrid.InvalidateCell(range.Row, range.Column);
            }

            window.spreadsheet.ActiveGrid.InvalidateCells();

            Logger.Log(LogItemType.Success, "Done.");
            isFinished = true;

            window.EnableTools();
        }
    }
}
