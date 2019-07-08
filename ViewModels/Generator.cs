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
        public bool HideConnections { get; set; }
        public List<GeneratedClass> GeneratedClasses { get; set; }
        public ObservableCollection<Vertex> OutputVertices { get; set; }


        public Generator(MainWindow mainWindow)
        {
            this.window = mainWindow;
        }

        public void GenerateGraph()
        {
            var logItem = Logger.Log(LogItemType.Info, "Generating graph...");
            var allCells = window.spreadsheet.ActiveSheet.Range[1, 1, window.spreadsheet.ActiveSheet.Rows.Length,
                window.spreadsheet.ActiveSheet.Columns.Length];
            var stopwatch = Stopwatch.StartNew();
            Graph = new Graph(allCells);
            logItem.AppendTime(stopwatch.ElapsedMilliseconds);

            Graph.GenerateLabels(window.spreadsheet.ActiveSheet);

            OutputVertices = new ObservableCollection<Vertex>(Graph.GetOutputFields());
            window.outputFieldsListView.ItemsSource = OutputVertices;

            PrepareUI();
        }

        private void PrepareUI()
        {
            window.diagram.Tool = Tool.ZoomPan | Tool.MultipleSelect;
            ((IGraphInfo)window.diagram.Info).ItemTappedEvent += (s, e1) => window.DiagramItemClicked(s, e1);
            window.diagram2.Tool = Tool.ZoomPan | Tool.MultipleSelect;
            ((IGraphInfo)window.diagram2.Info).ItemTappedEvent += (s, e1) => window.DiagramItemClicked(s, e1);
            window.spreadsheet.ActiveGrid.CurrentCellActivated += (s, e1) => window.SpreadsheetCellSelected(s, e1);

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
                window.outputFieldsListView.ScrollIntoView(OutputVertices.Where(v => v.Include).First());
                LoadDataIntoGraphAndSpreadsheet();
            }
        }

        public void LoadDataIntoGraphAndSpreadsheet()
        {
            window.EnableGraphOptions();

            var includedVertices = OutputVertices.Where(v => v.Include).ToList();
            Logger.Log(LogItemType.Info, $"Including selected output fields {string.Join(", ", includedVertices.Select(v => v.Address)) }");
            Graph.TransitiveFilter(includedVertices);
            LayoutGraph();
            ColorSpreadsheetCells();
            Logger.Log(LogItemType.Success, "Done.");

            App.Settings.SelectedOutputFields = includedVertices.Select(v => v.Address).ToList();
            App.Settings.Save();
        }

        private void LayoutGraph()
        {
            var logItem = Logger.Log(LogItemType.Info, "Layouting graph...");
            var stopwatch = Stopwatch.StartNew();

            window.diagram.Nodes = new NodeCollection();
            window.diagram.Connectors = new ConnectorCollection();

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

            if (window.diagram2.Groups is GroupCollection gc && gc.Count > 0)
                gc.Clear();

            window.diagram2.Nodes = new NodeCollection();
            window.diagram2.Connectors = new ConnectorCollection();
            window.diagram2.Groups = new GroupCollection();

            double nextPos = 0;
            foreach(var generatedClass in GeneratedClasses)
            {
                var (group, nextPosX) = generatedClass.FormatClass(nextPos);
                nextPos = nextPosX;
                (window.diagram2.Groups as GroupCollection).Add(group);
            }
            foreach (Vertex vertex in Graph.Vertices)
            {
                foreach (Vertex child in vertex.Children)
                {
                    if (HideConnections && vertex.Class != child.Class) continue;
                    (window.diagram2.Connectors as ConnectorCollection).Add(vertex.FormatEdge(child));
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
