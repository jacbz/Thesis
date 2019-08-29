using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using Syncfusion.UI.Xaml.CellGrid.Helpers;
using Syncfusion.UI.Xaml.Diagram;
using Syncfusion.XlsIO;
using Thesis.Models;
using Thesis.Models.VertexTypes;
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
            DisableGraphGenerationOptions();
            
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
            DisableClassGenerationOptions();
        }

        private void SelectVertexInSpreadsheet(Vertex vertex)
        {
            spreadsheet.SetActiveSheet(string.IsNullOrEmpty(vertex.ExternalWorksheetName) 
                ? _generator.ActiveWorksheet
                : vertex.ExternalWorksheetName);

            int row, col;
            if (vertex is CellVertex cell)
            {
                row = cell.Address.row;
                col = cell.Address.col;
            }
            else if (vertex is RangeVertex rangeVertex)
            {
                (row, col) = rangeVertex.StartAddress;
            }
            else return;

            spreadsheet.ActiveGrid.CurrentCell.MoveCurrentCell(row, col);

            // highlight selected vertex yellow for one second
            var spreadsheetCell = spreadsheet.ActiveSheet.Range[row, col];
            var originalBgColor = spreadsheetCell.CellStyle.Color;
            if (spreadsheet.Tag != null) return;
            spreadsheet.Tag = true;
            spreadsheetCell.CellStyle.Color = Color.Yellow;
            spreadsheet.ActiveGrid.InvalidateCell(row, col);
            Task.Factory.StartNew(() => Thread.Sleep(1000))
                .ContinueWith((t) =>
                {
                    spreadsheetCell.CellStyle.Color = originalBgColor;
                    spreadsheet.ActiveGrid.InvalidateCell(row, col);
                    spreadsheet.Tag = null;
                }, TaskScheduler.FromCurrentSynchronizationContext());
        }

        private void SelectVertexInDiagrams(Vertex vertex)
        {
            if (diagram.Nodes == null || !(vertex is CellVertex cell)) return;

            foreach (var node in (NodeCollection)diagram.Nodes)
                if (node.Content is CellVertex nodeCell)
                {
                    if (nodeCell.Address.row == cell.Address.row &&
                        nodeCell.Address.col == cell.Address.col)
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
            foreach (var group in (GroupCollection)diagram2.Groups)
            foreach (var node in (NodeCollection)group.Nodes)
                if (node.Content is CellVertex nodeCell)
                {
                    if (nodeCell.Address.row == cell.Address.row &&
                        nodeCell.Address.col == cell.Address.col)
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
            if (vertex is CellVertex cell && cell.NodeType == NodeType.OutputField)
            {
                outputFieldsListView.Tag = true; // avoiding triggering OutputFieldsListView_SelectionChanged
                outputFieldsListView.ScrollIntoView(cell);
                outputFieldsListView.SelectedItem = cell;
                outputFieldsListView.Tag = null;
            }
        }

        /// <summary>
        /// Reset all spreadsheet cell colors, then color cell backgrounds using label type, and borders using ColorSpreadsheetCells
        /// </summary>
        /// <param name="allVertices"></param>
        /// <param name="filteredVertices"></param>
        public void ResetAndColorAllCells(List<Vertex> allVertices)
        {
            ResetSpreadsheetColors();
            ColorSpreadsheetCells(allVertices.GetCellVertices(), StyleCellByLabelType, StyleBorderByNodeType);
            spreadsheet.ActiveGrid.InvalidateCells();
        }

        public void StyleCellByLabelType(CellVertex cellVertex, IStyle cellStyle)
        {
            Color color;
            switch (cellVertex.Label.Type)
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

        public void StyleBorderByNodeType(CellVertex cellVertex, IBorders borderStyle)
        {
            borderStyle.ColorRGB = cellVertex.GetNodeTypeColor();
            borderStyle.LineStyle = cellVertex.NodeType == NodeType.None ? ExcelLineStyle.None : ExcelLineStyle.Thick;
        }

        /// <summary>
        /// Color cells by given vertices and styling functions.
        /// </summary>
        /// <param name="vertices">Cells to style</param>
        /// <param name="styleCell">Cell styling function</param>
        /// <param name="styleBorder">Border styling function</param>
        public void ColorSpreadsheetCells(IEnumerable<Vertex> vertices, Action<CellVertex, IStyle> styleCell, Action<CellVertex, IBorders> styleBorder)
        {
            foreach (var vertex in vertices)
            {
                if (vertex is CellVertex cellVertex)
                {
                    var range = spreadsheet.ActiveSheet.Range[cellVertex.StringAddress];
                    styleCell(cellVertex, range.CellStyle);
                    styleBorder(cellVertex, range.Borders);
                }
                else if (vertex is RangeVertex rangeVertex)
                {
                    ColorSpreadsheetCells(rangeVertex.CellsInRange.Select(c => new CellVertex(c)), styleCell, styleBorder);
                }
            }
        }

        public void ColorSpreadsheetExternalCells(List<Vertex> externalVertices)
        {
            foreach(var externalVertex in externalVertices)
            {
                if (!spreadsheet.GridCollection.TryGetValue(externalVertex.ExternalWorksheetName, out var grid))
                    continue;

                if (externalVertex is CellVertex cellVertex)
                {
                    grid.Worksheet.Range[cellVertex.StringAddress].CellStyle.Color = Class.ExternalColor;
                    grid.InvalidateCell(cellVertex.Address.row, cellVertex.Address.col);
                }
                else if (externalVertex is RangeVertex rangeVertex)
                {
                    foreach (var cell in rangeVertex.CellsInRange)
                    {
                        cell.CellStyle.Color = Class.ExternalColor;
                        grid.InvalidateCell(cell.Row, cell.Column);
                    }
                }
            }
            // we don't reset the colors because that will take too long for all sheets
            // colors will thus remain even if they are not external vertices anymore
        }

        public void ResetSpreadsheetColors()
        {
            Logger.Log(LogItemType.Info, "Coloring cells...");

            var (rowCount, columnCount) = GetSheetDimensions();
            var allCells = spreadsheet.ActiveSheet.Range[1, 1, rowCount, columnCount];
            allCells.CellStyle.Color = Color.Transparent;
            allCells.CellStyle.Font.RGBColor = Color.Black;
        }
    }
}
