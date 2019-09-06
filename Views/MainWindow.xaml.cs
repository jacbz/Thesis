using System;
using System.IO;
using System.Text;
using System.Windows.Input;
using System.Windows.Threading;
using System.Xml;
using ICSharpCode.AvalonEdit.Folding;
using ICSharpCode.AvalonEdit.Highlighting;
using ICSharpCode.AvalonEdit.Highlighting.Xshd;
using Syncfusion.UI.Xaml.Diagram;
using Syncfusion.XlsIO;
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

            App.Settings = UserSettings.Read();
            if (!string.IsNullOrEmpty(App.Settings.SelectedFile)) LoadSpreadsheet();
            SetUpUi();
        }

        private void SetUpUi()
        {
            Formatter.InitFormatter();

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

            // improve diagram loading performance by virtualization
            diagram.Constraints |= GraphConstraints.Virtualize;
        }

        public (int rowCount, int columnCount) GetSheetDimensions()
        {
            if (spreadsheet.ActiveSheet == null) return (0, 0);
            var rowCount = spreadsheet.ActiveSheet.UsedRange.LastRow;
            var columnCount = spreadsheet.ActiveSheet.UsedRange.LastColumn;
            return (rowCount, columnCount);
        }

        public IRange GetRangeFromCurrentWorksheet(string range)
        {
            try
            {
                // will throw error for multi-ranges (e.g. A1:B2:C3:D4) and entire rows/columns (e.g. A:C, 1:3)
                // has been reported to Syncfusion, solution in the future
                return spreadsheet.ActiveSheet.Range[range];
            }
            catch (Exception)
            {
                Logger.Log(LogItemType.Error, $"Could not get range {range}. Entire rows/column currently not implemented.");
                return null;
            }
        }

        public void ProvideGraphGenerationOptions()
        {
            generateGraphButton.IsEnabled = labelGenerationOptionsGrid.IsEnabled = labelGenerationOptionsLabel.IsEnabled = true;
            filterGraphButton.IsEnabled = outputFieldsButtons.IsEnabled = outputFieldsListView.IsEnabled = false;
        }
        
        public void ProvideGraphFilteringOptions()
        {
            filterGraphButton.IsEnabled = outputFieldsButtons.IsEnabled = outputFieldsListView.IsEnabled = true;
        }

        public void DisableGraphOptions()
        {
            generateGraphButton.IsEnabled = filterGraphButton.IsEnabled = outputFieldsButtons.IsEnabled = outputFieldsListView.IsEnabled = false;
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
            // disable remove, rotate buttons etc. on click
            var selectedItem = diagram.SelectedItems as SelectorViewModel;
            (selectedItem.Commands as QuickCommandCollection).Clear();
            selectedItem.SelectorConstraints = selectedItem.SelectorConstraints
                .Remove(SelectorConstraints.Pivot)
                .Remove(SelectorConstraints.Rotator)
                .Remove(SelectorConstraints.Resizer);
        }

        private void InitiateToolbox(dynamic obj)
        {
            if (obj != null)
            {
                toolboxContent.Opacity = 1f;
                toolboxContent.IsEnabled = true;
                toolboxTab.IsSelected = true;
                DataContext = obj;

                // select and focus name text box
                Dispatcher.BeginInvoke(DispatcherPriority.Input,
                    new Action(delegate
                    {
                        nameTextBox.Focus();
                        Keyboard.Focus(nameTextBox);
                        nameTextBox.SelectAll();
                    }));
            }
            else
            {
                toolboxContent.Opacity = 0.3f;
                toolboxContent.IsEnabled = false;
            }
        }

        private void SelectLogTab()
        {
            logTab.IsSelected = true;
        }
    }
}