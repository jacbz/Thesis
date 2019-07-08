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
using Thesis.Models;

namespace Thesis.ViewModels
{
    class Generator
    {
        private MainWindow window;
        public Graph Graph { get; set; }
        public bool IsFinishedGeneratingGraph { get; set; }
        public bool HideConnections { get; set; }
        public List<GeneratedClass> GeneratedClasses { get; set; }
        public ObservableCollection<Vertex> OutputVertices { get; set; }


        public Generator(MainWindow mainWindow)
        {
            this.window = mainWindow;
            this.IsFinishedGeneratingGraph = false;
        }

        public void worker_Generate(object sender, DoWorkEventArgs e)
        {
            var logItem = Logger.Log(LogItemType.Info, "Generating graph...");
            var allCells = window.spreadsheet.ActiveSheet.Range[1, 1, window.spreadsheet.ActiveSheet.Rows.Length,
                window.spreadsheet.ActiveSheet.Columns.Length];
            var stopwatch = Stopwatch.StartNew();
            Graph = new Graph(allCells);
            logItem.AppendTime(stopwatch.ElapsedMilliseconds);
        }

        public void worker_GenerateCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            LayoutGraph();

            window.diagram.Tool = Tool.ZoomPan | Tool.MultipleSelect;
            ((IGraphInfo)window.diagram.Info).ItemTappedEvent += (s, e1) => window.DiagramItemClicked(s, e1);
            window.spreadsheet.ActiveGrid.CellClick += (s, e1) => window.SpreadsheetCellClicked(s, e1);

            ColorSpreadsheetCells();

            FormatOptionsTab();

            Logger.Log(LogItemType.Success, "Done.");
            IsFinishedGeneratingGraph = true;

            window.EnableGraphOptions();

            // load selected output fields from settings
            if (App.Settings.SelectedOutputFields != null &&
                !OutputVertices.Where(v => v.Include).Select(v => v.Address).ToList().SequenceEqual(App.Settings.SelectedOutputFields))
            {
                Logger.Log(LogItemType.Info, "Loading selected output fields from user settings");
                UnselectAllOutputFields();
                foreach (Vertex v in OutputVertices)
                {
                    if (App.Settings.SelectedOutputFields.Contains(v.Address))
                        v.Include = true;
                }
                FilterForSelectedOutputFields();
            }
        }

        private void LayoutGraph()
        {
            var logItem = Logger.Log(LogItemType.Info, "Layouting graph...");
            var stopwatch = Stopwatch.StartNew();

            window.diagram.Nodes = new NodeCollection();
            window.diagram.Connectors = new ConnectorCollection();

            if (window.diagram.Groups != null)
            {
                (window.diagram.Groups as GroupCollection).Clear();
            }

            foreach (Vertex vertex in Graph.Vertices)
            {
                (window.diagram.Nodes as NodeCollection).Add(vertex.FormatVertex(Graph));
            }
            foreach (Vertex vertex in Graph.Vertices)
            {
                foreach (Vertex child in vertex.Children)
                {
                    (window.diagram.Connectors as ConnectorCollection).Add(vertex.FormatEdge(child));
                }
            }

            logItem.AppendTime(stopwatch.ElapsedMilliseconds);
        }

        public void LayoutClasses()
        {
            var logItem = Logger.Log(LogItemType.Info, "Layouting classes...");
            var stopwatch = Stopwatch.StartNew();

            if (window.diagram.Groups != null)
            {
                (window.diagram.Groups as GroupCollection).Clear();
            }

            window.diagram.Nodes = new NodeCollection();
            window.diagram.Connectors = new ConnectorCollection();
            window.diagram.Groups = new GroupCollection();

            double nextPos = 0;
            foreach(var generatedClass in GeneratedClasses)
            {
                var formattedClass = generatedClass.FormatClass(nextPos);
                nextPos = formattedClass.nextPosX;
                (window.diagram.Groups as GroupCollection).Add(formattedClass.group);
            }
            foreach (Vertex vertex in Graph.Vertices)
            {
                foreach (Vertex child in vertex.Children)
                {
                    if (HideConnections && vertex.Class != child.Class) continue;
                    (window.diagram.Connectors as ConnectorCollection).Add(vertex.FormatEdge(child));
                }
            }

            logItem.AppendTime(stopwatch.ElapsedMilliseconds);
        }

        private void ColorSpreadsheetCells()
        {
            ResetSpreadsheetColors();
            ColorSpreadsheetCellsInner(Graph.Vertices, v => Color.White, v => ColorTranslator.FromHtml(v.GetColor()));
            window.spreadsheet.ActiveGrid.InvalidateCells();
        }

        private void ColorSpreadsheetCellsInner(List<Vertex> vertices, Func<Vertex, Color> cellColorFunc, Func<Vertex, Color> borderColorFunc)
        {
            foreach (Vertex v in vertices)
            {
                IRange range = window.spreadsheet.ActiveSheet.Range[v.Address];
                range.CellStyle.Color = cellColorFunc(v);
                range.CellStyle.Font.RGBColor = cellColorFunc(v).GetTextColor();
                range.Borders.ColorRGB = borderColorFunc(v);
                range.Borders.LineStyle = ExcelLineStyle.Thick;

                window.spreadsheet.ActiveGrid.InvalidateCell(range.Row, range.Column);
            }
        }

        private void ResetSpreadsheetColors()
        {
            Logger.Log(LogItemType.Info, "Coloring cells...");

            var allCells = window.spreadsheet.ActiveSheet.Range[1, 1, window.spreadsheet.ActiveSheet.Rows.Length, window.spreadsheet.ActiveSheet.Columns.Length];
            allCells.CellStyle.Color = Color.Transparent;
            allCells.Borders.LineStyle = ExcelLineStyle.None;
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

            App.Settings.SelectedOutputFields = includedVertices.Select(v => v.Address).ToList();
            App.Settings.Save();
        }

        public void SelectAllOutputFields()
        {
            foreach (Vertex vertex in OutputVertices)
            {
                vertex.Include = true;
            }
        }

        public void UnselectAllOutputFields()
        {
            foreach (Vertex vertex in OutputVertices)
            {
                vertex.Include = false;
            }
        }

        public void GenerateClasses()
        {
            Logger.Log(LogItemType.Info, "Generate classes for selected output fields...");
            GeneratedClasses = Graph.GenerateClasses();

            ResetSpreadsheetColors();
            for (int i = 0; i < GeneratedClasses.Count; i++)
            {
                GeneratedClass generatedClass = GeneratedClasses[i];
                ColorSpreadsheetCellsInner(generatedClass.Vertices, v => generatedClass.Color, v => ColorTranslator.FromHtml(v.GetColor()));
            }
            window.spreadsheet.ActiveGrid.InvalidateCells();

            LayoutClasses();

            Logger.Log(LogItemType.Success, $"Generated {GeneratedClasses.Count} classes.");
        }
    }
}
