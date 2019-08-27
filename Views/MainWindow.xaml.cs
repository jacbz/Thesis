using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Xml;
using ICSharpCode.AvalonEdit.Folding;
using ICSharpCode.AvalonEdit.Highlighting;
using ICSharpCode.AvalonEdit.Highlighting.Xshd;
using Syncfusion.UI.Xaml.Diagram;
using Syncfusion.XlsIO;
using Thesis.Models.VertexTypes;
using Thesis.ViewModels;

namespace Thesis.Views
{
    /// <summary>
    ///     Basic UI logic for MainWindow.xaml
    /// 
    ///     Events are found in MainWindowEvents.cs
    ///     Advanced UI manipulation found in MainWindowActions.cs
    /// </summary>
    public partial class MainWindow
    {
        private Generator _generator;
        private FoldingManager _foldingManager;
        private BraceFoldingStrategy _foldingStrategy;

        public MainWindow()
        {
            InitializeComponent();
            Logger.Instantiate(logControl.logListView, SelectLogTab);
            if (!string.IsNullOrEmpty(App.Settings.FilePath)) LoadSpreadsheet();
            SetUpUi();
        }

        private void SetUpUi()
        {
            // enable folding in code text box
            _foldingManager = FoldingManager.Install(codeTextBox.TextArea);
            _foldingStrategy = new BraceFoldingStrategy();
            codeTextBox.TextArea.SelectionChanged += CodeTextBoxSelectionChanged;

            // enable syntax highlighting
            using (Stream stream = new MemoryStream(Encoding.UTF8.GetBytes(Properties.Resources.CSharpSyntaxHighlighting ?? "")))
                using (XmlTextReader reader = new XmlTextReader(stream))
                    codeTextBox.SyntaxHighlighting = HighlightingLoader.Load(reader, HighlightingManager.Instance);

            // disable pasting in spreadsheet
            spreadsheet.HistoryManager.Enabled = false;
            spreadsheet.CopyPaste.Pasting += (sender, e) => e.Cancel = true;

            diagram.Tool = Tool.ZoomPan | Tool.MultipleSelect;
            ((IGraphInfo)diagram.Info).AnnotationChanged += DiagramAnnotationChanged;
            ((IGraphInfo)diagram.Info).ItemTappedEvent += DiagramItemClicked;
            diagram2.Tool = Tool.ZoomPan | Tool.MultipleSelect;
            ((IGraphInfo)diagram2.Info).AnnotationChanged += DiagramAnnotationChanged;
            ((IGraphInfo)diagram2.Info).ItemTappedEvent += DiagramItemClicked;

            // improve diagram loading performance by virtualization
            diagram.Constraints = diagram.Constraints | GraphConstraints.Virtualize;
            diagram2.Constraints = diagram2.Constraints | GraphConstraints.Virtualize;
        }

        public (int rowCount, int columnCount) GetSheetDimensions()
        {
            if (spreadsheet.ActiveSheet == null) return (0, 0);
            var rowCount = spreadsheet.ActiveSheet.UsedRange.LastRow;
            var columnCount = spreadsheet.ActiveSheet.UsedRange.LastColumn;
            return (rowCount, columnCount);
        }

        public IRange GetCellFromWorksheet(string sheetName, string address)
        {
            return spreadsheet.GridCollection.TryGetValue(sheetName, out var grid) 
                ? grid.Worksheet.Range[address] 
                : null;
        }

        public void EnableGraphGenerationOptions()
        {
            generateGraphButton.IsEnabled = selectAllButton.IsEnabled = unselectAllButton.IsEnabled = true;
        }
        
        public void DisableGraphGenerationOptions()
        {
            generateGraphButton.IsEnabled = selectAllButton.IsEnabled = unselectAllButton.IsEnabled = false;
        }

        public void EnableClassGenerationOptions()
        {
            toolboxTab.IsEnabled = true;
            generateClassesButton.IsEnabled = hideConnectionsCheckbox.IsEnabled = true;
        }

        private void DisableClassGenerationOptions()
        {
            _generator = new Generator(this);
            toolboxTab.IsEnabled = false;
            logTab.IsSelected = true;
            diagram.Nodes = new NodeCollection();
            diagram.Connectors = new ConnectorCollection();
            generateClassesButton.IsEnabled = hideConnectionsCheckbox.IsEnabled = false;
            DisableCodeGenerationOptions();
        }

        private void EnableCodeGenerationOptions()
        {
            generateCodeButton.IsEnabled = true;
        }

        private void DisableCodeGenerationOptions()
        {
            generateCodeButton.IsEnabled = false;
            testCodeButton.IsEnabled = false;
        }

        private void DisableDiagramNodeTools()
        {
            DisableDiagramNodeTools(diagram);
            DisableDiagramNodeTools(diagram2);
        }

        private void DisableDiagramNodeTools(SfDiagram diagram)
        {
            // disable remove, rotate buttons etc. on click
            var selectedItem = diagram.SelectedItems as SelectorViewModel;
            (selectedItem.Commands as QuickCommandCollection).Clear();
            selectedItem.SelectorConstraints = selectedItem.SelectorConstraints.Remove(SelectorConstraints.Rotator);
        }

        private void InitiateToolbox(Vertex vertex)
        {
            if (vertex == null)
            {
                toolboxContent.Opacity = 0.3f;
                toolboxContent.IsEnabled = false;
            }
            else
            {
                toolboxContent.Opacity = 1f;
                toolboxContent.IsEnabled = true;
                toolboxTab.IsSelected = true;
                DataContext = vertex;
            }
        }

        private void SelectLogTab()
        {
            logTab.IsSelected = true;
        }
    }
}