using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using ICSharpCode.AvalonEdit.Folding;
using Syncfusion.UI.Xaml.Diagram;
using Thesis.Models;
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

            // disable pasting in spreadsheet
            spreadsheet.HistoryManager.Enabled = false;
            spreadsheet.CopyPaste.Pasting += (sender, e) => e.Cancel = true;

            diagram.Tool = Tool.ZoomPan | Tool.MultipleSelect;
            ((IGraphInfo)diagram.Info).AnnotationChanged += DiagramAnnotationChanged;
            ((IGraphInfo)diagram.Info).ItemTappedEvent += DiagramItemClicked;
            diagram2.Tool = Tool.ZoomPan | Tool.MultipleSelect;
            ((IGraphInfo)diagram2.Info).AnnotationChanged += DiagramAnnotationChanged;
            ((IGraphInfo)diagram2.Info).ItemTappedEvent += DiagramItemClicked;
        }

        public void EnableGraphOptions()
        {
            toolboxTab.IsEnabled = true;
            generateClassesButton.IsEnabled = hideConnectionsCheckbox.IsEnabled = true;
        }

        private void DisableGraphOptions()
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