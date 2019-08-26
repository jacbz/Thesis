﻿using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
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
        public ClassCollection ClassCollection { get; set; }
        public Code Code { get; set; }

        public bool HideConnections { get; set; }
        public ObservableCollection<Vertex> OutputVertices { get; set; }
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
            var allCells = _window.spreadsheet.ActiveSheet.Range[1, 1, rowCount, columnCount];

            Graph = Graph.FromSpreadsheet(ActiveWorksheet, allCells, _window.GetCellFromWorksheet, _window.spreadsheet.Workbook.Names);
            Logger.Log(LogItemType.Success, "Graph generation successful.");

            _window.generateGraphButton.IsEnabled = true;
            logItem.AppendElapsedTime();

            OutputVertices = new ObservableCollection<Vertex>(Graph.GetOutputFields());
            _window.outputFieldsListView.ItemsSource = OutputVertices;

            LoadPersistedOutputFields();

            return true;
        }

        // return value: success
        private bool LoadPersistedOutputFields()
        {
            // load selected output fields from settings
            if (App.Settings.SelectedOutputFields != null &&
                App.Settings.SelectedOutputFields.Count > 0 &&
                OutputVertices.Count > 0 && 
                !OutputVertices.Where(v => v.Include).Select(v => v.StringAddress).ToList()
                    .SequenceEqual(App.Settings.SelectedOutputFields))
            {
                Logger.Log(LogItemType.Info, "Loading selected output fields from persisted user settings...");
                UnselectAllOutputFields();
                foreach (var v in OutputVertices)
                {
                    if (App.Settings.SelectedOutputFields.Contains(v.StringAddress))
                    {
                        v.Include = true;
                    }
                }

                var firstIncludedOutputVertex = OutputVertices.FirstOrDefault(v => v.Include);
                if (firstIncludedOutputVertex != null)
                    _window.outputFieldsListView.ScrollIntoView(firstIncludedOutputVertex);
                return true;
            }

            return false;
        }

        /// <summary>
        /// Filters by selected output vertices, displays graph, and colors spreadsheet cells
        /// </summary>
        public void FilterAndDisplayGraphIntoUi()
        {
            if (_window.spreadsheet.ActiveSheet.Name != ActiveWorksheet)
            {
                App.Settings.ResetWorkbookSpecificSettings();
                GenerateGraph();
            }

            _window.EnableClassGenerationOptions();

            var includedVertices = OutputVertices.Where(v => v.Include).ToList();
            var includedVertixStrings = includedVertices.Select(v => v.StringAddress).ToList();
            App.Settings.SelectedOutputFields = includedVertixStrings;
            App.Settings.Persist();

            Logger.Log(LogItemType.Info,
                $"Selected output fields: {string.Join(", ", includedVertixStrings)}");

            Graph.PerformTransitiveFilter(includedVertices);

            LayoutGraph();

            _window.ResetAndColorAllCells(Graph.AllVertices);

            Logger.Log(LogItemType.Success, "Graph generation complete.");
        }

        private void LayoutGraph()
        {
            var logItem = Logger.Log(LogItemType.Info, "Layouting graph...", true);

            _window.diagram.Nodes = new NodeCollection();
            _window.diagram.Connectors = new ConnectorCollection();

            foreach (var vertex in Graph.Vertices)
                ((NodeCollection) _window.diagram.Nodes).Add(vertex.FormatVertex(Graph));
            foreach (var vertex in Graph.Vertices)
            foreach (var child in vertex.Children)
                ((ConnectorCollection) _window.diagram.Connectors).Add(vertex.FormatEdge(child, true));

            logItem.AppendElapsedTime();
        }

        public void LayoutClasses()
        {
            var logItem = Logger.Log(LogItemType.Info, "Layouting classes...", true);

            _window.diagram2.Nodes = new NodeCollection();
            _window.diagram2.Connectors = new ConnectorCollection();
            _window.diagram2.Groups = new GroupCollection();

            double nextPos = 0;
            foreach (var generatedClass in ClassCollection.Classes)
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

        public Vertex GetVertexByAddress(int row, int col)
        {
            return Graph.AllVertices
                .FirstOrDefault(v =>
                    v.Address.row == row &&
                    v.Address.col == col);
        }

        public void GenerateClasses()
        {
            var logItem = Logger.Log(LogItemType.Info, "Generate classes for selected output fields...", true);

            ClassCollection = ClassCollection.FromGraph(Graph);

            logItem.AppendElapsedTime();
            _window.ResetSpreadsheetColors();
            foreach (var generatedClass in ClassCollection.Classes)
            {
                _window.ColorSpreadsheetCells(generatedClass.Vertices.Where(v => v.NodeType != NodeType.External), 
                    (vertex, style) => { _window.StyleCellByColor(generatedClass.Color, style); }, 
                    _window.StyleBorderByNodeType);
            }

            _window.ColorSpreadsheetExternalCells(Graph.ExternalVertices);
            _window.spreadsheet.ActiveGrid.InvalidateCells();

            LayoutClasses();

            Logger.Log(LogItemType.Success, $"Generated {ClassCollection.Classes.Count} classes.");
        }

        public async Task GenerateCode()
        {
            var logItem = Logger.Log(LogItemType.Info, $"Generating code...", true);

            _window.codeTextBox.Text = "";

            var addressToVertexDictionary = Graph.Vertices
                .Where(v => !string.IsNullOrEmpty(v.StringAddress))
                .ToDictionary(v => v.StringAddress, v => v);
            // implement different languages here
            CodeGenerator codeGenerator;
            switch (_window.languageComboBox.SelectedIndex)
            {
                default:
                    codeGenerator = new CSharpGenerator(ClassCollection, addressToVertexDictionary, Graph.NamedRangeDictionary);
                    break;
            }

            Code = await Code.GenerateFrom(codeGenerator);
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
                           && report.SkippedCount == 0
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