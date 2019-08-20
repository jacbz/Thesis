using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Windows;
using System.Windows.Navigation;
using Syncfusion.UI.Xaml.CellGrid.Helpers;
using Syncfusion.UI.Xaml.Diagram;
using Syncfusion.XlsIO;
using Thesis.Models;
using Thesis.Models.CodeGenerators;
using Thesis.Views;

namespace Thesis.ViewModels
{
    public class Generator
    {
        private readonly MainWindow _window;

        public Graph Graph { get; set; }
        public bool HideConnections { get; set; }
        public List<GeneratedClass> GeneratedClasses { get; set; }
        public ObservableCollection<Vertex> OutputVertices { get; set; }
        public string ActiveWorksheet { get; set; }

        public Generator(MainWindow mainWindow)
        {
            _window = mainWindow;
        }

        public void GenerateGraph()
        {
            ActiveWorksheet = _window.spreadsheet.ActiveSheet.Name;
            App.Settings.SelectedWorksheet = ActiveWorksheet;
            App.Settings.Save();

            var logItem = Logger.Log(LogItemType.Info, "Generating graph...");
            var allCells = _window.spreadsheet.ActiveSheet.Range[1, 1, _window.spreadsheet.ActiveSheet.Rows.Length,
                _window.spreadsheet.ActiveSheet.Columns.Length];
            var stopwatch = Stopwatch.StartNew();
            Graph = new Graph(allCells);
            logItem.AppendTime(stopwatch.ElapsedMilliseconds);

            OutputVertices = new ObservableCollection<Vertex>(Graph.GetOutputFields());
            _window.outputFieldsListView.ItemsSource = OutputVertices;

            PrepareUi();
        }

        private void PrepareUi()
        {
            _window.diagram.Tool = Tool.ZoomPan | Tool.MultipleSelect;
            ((IGraphInfo)_window.diagram.Info).AnnotationChanged += _window.DiagramAnnotationChanged;
            ((IGraphInfo) _window.diagram.Info).ItemTappedEvent += (s, e1) => _window.DiagramItemClicked(s, e1);
            _window.diagram2.Tool = Tool.ZoomPan | Tool.MultipleSelect;
            ((IGraphInfo)_window.diagram2.Info).AnnotationChanged += _window.DiagramAnnotationChanged;
            ((IGraphInfo) _window.diagram2.Info).ItemTappedEvent += (s, e1) => _window.DiagramItemClicked(s, e1);
            _window.spreadsheet.ActiveGrid.CurrentCellActivated += (s, e1) => _window.SpreadsheetCellSelected(s, e1);

            // load selected output fields from settings
            if (App.Settings.SelectedOutputFields != null &&
                App.Settings.SelectedOutputFields.Count > 0 &&
                OutputVertices.Count > 0 && 
                !OutputVertices.Where(v => v.Include).Select(v => v.StringAddress).ToList()
                    .SequenceEqual(App.Settings.SelectedOutputFields))
            {
                Logger.Log(LogItemType.Info, "Loading selected output fields from user settings");
                UnselectAllOutputFields();
                foreach (var v in OutputVertices)
                {
                    if (App.Settings.SelectedOutputFields.Contains(v.StringAddress))
                    {
                        v.Include = true;
                    }
                }

                _window.outputFieldsListView.ScrollIntoView(OutputVertices.First(v => v.Include));
                LoadDataIntoGraphAndSpreadsheet();
            }
        }

        public void LoadDataIntoGraphAndSpreadsheet()
        {
            if (_window.spreadsheet.ActiveSheet.Name != ActiveWorksheet)
            {
                App.Settings.ResetWorkbookSpecificSettings();
                GenerateGraph();
            }

            _window.EnableGraphOptions();

            var includedVertices = OutputVertices.Where(v => v.Include).ToList();
            Logger.Log(LogItemType.Info,
                $"Including selected output fields {string.Join(", ", includedVertices.Select(v => v.StringAddress))}");
            Graph.TransitiveFilter(includedVertices);
            LayoutGraph();
            ColorSpreadsheetCells();
            Logger.Log(LogItemType.Success, "Done.");

            App.Settings.SelectedOutputFields = includedVertices.Select(v => v.StringAddress).ToList();
            App.Settings.Save();
        }

        private void LayoutGraph()
        {
            var logItem = Logger.Log(LogItemType.Info, "Layouting graph...");
            var stopwatch = Stopwatch.StartNew();

            _window.diagram.Nodes = new NodeCollection();
            _window.diagram.Connectors = new ConnectorCollection();

            foreach (var vertex in Graph.Vertices)
                ((NodeCollection) _window.diagram.Nodes).Add(vertex.FormatVertex(Graph));
            foreach (var vertex in Graph.Vertices)
            foreach (var child in vertex.Children)
                ((ConnectorCollection) _window.diagram.Connectors).Add(vertex.FormatEdge(child));

            logItem.AppendTime(stopwatch.ElapsedMilliseconds);
        }

        public void LayoutClasses()
        {
            var logItem = Logger.Log(LogItemType.Info, "Layouting classes...");
            var stopwatch = Stopwatch.StartNew();

            _window.diagram2.Nodes = new NodeCollection();
            _window.diagram2.Connectors = new ConnectorCollection();
            _window.diagram2.Groups = new GroupCollection();

            double nextPos = 0;
            foreach (var generatedClass in GeneratedClasses)
            {
                var (group, nextPosX) = generatedClass.FormatClass(nextPos);
                nextPos = nextPosX;
                (_window.diagram2.Groups as GroupCollection).Add(group);
            }

            foreach (var vertex in Graph.Vertices)
            foreach (var child in vertex.Children)
            {
                if (HideConnections && vertex.Class != child.Class) continue;
                ((ConnectorCollection) _window.diagram2.Connectors).Add(vertex.FormatEdge(child));
            }

            logItem.AppendTime(stopwatch.ElapsedMilliseconds);
        }

        private void ColorSpreadsheetCells()
        {
            ResetSpreadsheetColors();
            //ColorSpreadsheetCellsInner(Graph.Vertices, v => Color.White);
            ColorSpreadsheetCellsInner(Graph.AllVertices, v =>
            {
                switch (v.Label.Type)
                {
                    case LabelType.Attribute:
                        return ((System.Windows.Media.Color)Application.Current.Resources["AttributeColor"]).ToDColor();
                    case LabelType.Data:
                        return ((System.Windows.Media.Color)Application.Current.Resources["DataColor"]).ToDColor();
                    case LabelType.Header:
                        return ((System.Windows.Media.Color)Application.Current.Resources["HeaderColor"]).ToDColor();
                    default:
                        return Color.Transparent;
                }
            });
            _window.spreadsheet.ActiveGrid.InvalidateCells();
        }

        private void ColorSpreadsheetCellsInner(List<Vertex> vertices, Func<Vertex, Color> cellColorFunc)
        {
            foreach (var v in vertices)
            {
                var range = _window.spreadsheet.ActiveSheet.Range[v.StringAddress];
                range.CellStyle.Color = cellColorFunc(v);
                range.CellStyle.Font.RGBColor = cellColorFunc(v).GetTextColor();
                if (!Graph.Vertices.Contains(v))
                {
                    range.Borders.LineStyle = ExcelLineStyle.None;
                }
                else
                {
                    range.Borders.ColorRGB = v.GetNodeTypeColor();
                    range.Borders.LineStyle = ExcelLineStyle.Thick;
                }
            }
            _window.spreadsheet.ActiveGrid.InvalidateCells();
        }

        private void ResetSpreadsheetColors()
        {
            Logger.Log(LogItemType.Info, "Coloring cells...");

            var allCells = _window.spreadsheet.ActiveSheet.Range[1, 1, _window.spreadsheet.ActiveSheet.Rows.Length,
                _window.spreadsheet.ActiveSheet.Columns.Length];
            allCells.CellStyle.Color = Color.Transparent;
            allCells.CellStyle.Font.RGBColor = Color.Black;
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
            foreach (var generatedClass in GeneratedClasses)
            {
                ColorSpreadsheetCellsInner(generatedClass.Vertices, v => generatedClass.Color);
            }

            _window.spreadsheet.ActiveGrid.InvalidateCells();

            LayoutClasses();

            Logger.Log(LogItemType.Success, $"Generated {GeneratedClasses.Count} classes.");
        }

        public void GenerateCode()
        {
            var logItem = Logger.Log(LogItemType.Info, $"Generating code...");

            _window.codeTextBox.Text = "";
            var stopwatch = Stopwatch.StartNew();

            var addressToVertexDictionary = Graph.Vertices.ToDictionary(v => v.StringAddress, v => v);
            CodeGenerator codeGenerator;
            // implement different languages here
            switch (_window.languageComboBox.SelectedIndex)
            {
                default:
                    codeGenerator = new CSharpGenerator(GeneratedClasses, addressToVertexDictionary);
                    break;
            }
            string code = codeGenerator.GetCode();
            logItem.AppendTime(stopwatch.ElapsedMilliseconds);

            _window.codeTextBox.Text = code;
            Logger.Log(LogItemType.Success, $"Successfully generated code.");
        }
    }
}