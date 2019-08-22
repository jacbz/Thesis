using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using Syncfusion.UI.Xaml.Diagram;
using Syncfusion.XlsIO;
using Thesis.Models;
using Thesis.ViewModels;

namespace Thesis.Views
{
    /// <summary>
    ///     Manipulates UI elements for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        private void LoadSpreadsheet()
        {
            pathLabel.Content = pathLabel.ToolTip = App.Settings.FilePath;
            pathLabel.FontStyle = FontStyles.Normal;

            Logger.Log(LogItemType.Info, $"Loading {App.Settings.FilePath}");
            try
            {
                spreadsheet.Open(App.Settings.FilePath);
            }
            catch (IOException e)
            {
                Logger.Log(LogItemType.Error, "Could not open file: " + e.Message);
            }

            spreadsheet.Opacity = 100;
            generateGraphButton.IsEnabled = selectAllButton.IsEnabled = unselectAllButton.IsEnabled = true;
            DisableGraphOptions();
        }

        private void SelectVertexInSpreadsheet(Vertex vertex)
        {
            spreadsheet.SetActiveSheet(_generator.ActiveWorksheet);
            spreadsheet.ActiveGrid.CurrentCell.MoveCurrentCell(vertex.Address.row, vertex.Address.col);

            // highlight selected vertex yellow for one second
            var cell = spreadsheet.ActiveSheet.Range[vertex.StringAddress];
            var originalBgColor = cell.CellStyle.Color;
            if (spreadsheet.Tag != null) return;
            spreadsheet.Tag = true;
            cell.CellStyle.Color = Color.Yellow;
            spreadsheet.ActiveGrid.InvalidateCell(vertex.Address.row, vertex.Address.col);
            Task.Factory.StartNew(() => Thread.Sleep(1000))
                .ContinueWith((t) =>
                {
                    cell.CellStyle.Color = originalBgColor;
                    spreadsheet.ActiveGrid.InvalidateCell(vertex.Address.row, vertex.Address.col);
                    spreadsheet.Tag = null;
                }, TaskScheduler.FromCurrentSynchronizationContext());
        }

        private void SelectVertexInDiagrams(Vertex vertex)
        {
            if (diagram.Nodes == null) return;
            foreach (var node in (DiagramCollection<NodeViewModel>)diagram.Nodes)
                if (node.Content is Vertex nodeVertex)
                {
                    if (nodeVertex.Address.row == vertex.Address.row &&
                        nodeVertex.Address.col == vertex.Address.col)
                    {
                        node.IsSelected = true;
                        (diagram.Info as IGraphInfo).BringIntoCenter((node.Info as INodeInfo).Bounds);
                        DisableDiagramNodeTools(diagram);
                    }
                    else
                    {
                        node.IsSelected = false;
                    }
                }

            if (diagram2.Groups == null) return;
            foreach (var group in (DiagramCollection<GroupViewModel>)diagram2.Groups)
            foreach (var node in (ObservableCollection<NodeViewModel>)group.Nodes)
                if (node.Content is Vertex nodeVertex)
                {
                    if (nodeVertex.Address.row == vertex.Address.row &&
                        nodeVertex.Address.col == vertex.Address.col)
                    {
                        node.IsSelected = true;
                        (diagram2.Info as IGraphInfo).BringIntoCenter((node.Info as INodeInfo).Bounds);
                        DisableDiagramNodeTools(diagram2);
                    }
                    else
                    {
                        node.IsSelected = false;
                    }
                }
        }

        private void SelectVertexInOutputListView(Vertex vertex)
        {
            if (vertex.NodeType == NodeType.OutputField)
            {
                outputFieldsListView.Tag = true; // avoiding triggering OutputFieldsListView_SelectionChanged
                outputFieldsListView.ScrollIntoView(vertex);
                outputFieldsListView.SelectedItem = vertex;
                outputFieldsListView.Tag = null;
            }
        }

        /// <summary>
        /// Reset all spreadsheet cell colors, then color cell backgrounds using label type, and borders using ColorSpreadsheetCells
        /// </summary>
        /// <param name="allVertices"></param>
        /// <param name="filteredVertices"></param>
        public void ResetAndColorAllCells(List<Vertex> allVertices, List<Vertex> filteredVertices)
        {
            ResetSpreadsheetColors();
            ColorSpreadsheetCells(allVertices, StyleCellByLabelType, StyleBorderByNodeType);
            spreadsheet.ActiveGrid.InvalidateCells();
        }

        public void StyleCellByLabelType(Vertex vertex, IStyle cellStyle)
        {
            Color color;
            switch (vertex.Label.Type)
            {
                case LabelType.Attribute:
                    color = ((System.Windows.Media.Color)Application.Current.Resources["AttributeColor"]).ToDColor();
                    break;
                case LabelType.Data:
                    color = ((System.Windows.Media.Color)Application.Current.Resources["DataColor"]).ToDColor();
                    break;
                case LabelType.Header:
                    color = ((System.Windows.Media.Color)Application.Current.Resources["HeaderColor"]).ToDColor();
                    break;
                default:
                    color = Color.Transparent;
                    break;
            }

            StyleCellByColor(color, cellStyle);
        }

        public void StyleCellByColor(Color color, IStyle cellStyle)
        {
            cellStyle.Color = color;
            cellStyle.Font.RGBColor = color.GetTextColor();
        }

        public void StyleBorderByNodeType(Vertex vertex, IBorders borderStyle)
        {
            borderStyle.ColorRGB = vertex.GetNodeTypeColor();
            borderStyle.LineStyle = vertex.NodeType == NodeType.None ? ExcelLineStyle.None : ExcelLineStyle.Thick;
        }

        /// <summary>
        /// Color spreadsheet cells by given vertices and styling functions.
        /// </summary>
        /// <param name="vertices">Cells to style</param>
        /// <param name="styleCell">Cell styling function</param>
        /// <param name="styleBorder">Border styling function</param>
        public void ColorSpreadsheetCells(List<Vertex> vertices, Action<Vertex, IStyle> styleCell, Action<Vertex, IBorders> styleBorder)
        {
            foreach (var vertex in vertices)
            {
                var range = spreadsheet.ActiveSheet.Range[vertex.StringAddress];

                styleCell(vertex, range.CellStyle);
                styleBorder(vertex, range.Borders);
            }
            spreadsheet.ActiveGrid.InvalidateCells();
        }

        public void ResetSpreadsheetColors()
        {
            Logger.Log(LogItemType.Info, "Coloring cells...");

            var allCells = spreadsheet.ActiveSheet.Range[1, 1, spreadsheet.ActiveSheet.Rows.Length,
                spreadsheet.ActiveSheet.Columns.Length];
            allCells.CellStyle.Color = Color.Transparent;
            allCells.CellStyle.Font.RGBColor = Color.Black;
            allCells.Borders.LineStyle = ExcelLineStyle.None;
        }
    }
}
