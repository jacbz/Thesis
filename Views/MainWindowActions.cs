// Thesis - An Excel to code converter
// Copyright (C) 2019 Jacob Zhang
// 
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
// 
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
// 
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <https://www.gnu.org/licenses/>.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using Syncfusion.UI.Xaml.Diagram;
using Syncfusion.UI.Xaml.Spreadsheet.Helpers;
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
            DisableGraphOptions();
            
            pathLabel.Content = pathLabel.ToolTip = App.Settings.SelectedFile;
            pathLabel.FontStyle = FontStyles.Normal;

            Logger.Log(LogItemType.Info, $"Loading {App.Settings.SelectedFile}");
            try
            {
                spreadsheet.Open(App.Settings.SelectedFile);
            }
            catch (IOException e)
            {
                Logger.Log(LogItemType.Error, "Could not open file: " + e.Message);
                generateGraphButton.IsEnabled = magicButton.IsEnabled = false;
            }

            spreadsheet.Opacity = 100;
            DisableClassGenerationOptions();
        }

        private void SaveCustomClassNames()
        {
            var classNames = _generator?.ClassCollection?.GetCustomClassNames();
            if (classNames == null) return;

            foreach (var kvp in _generator.ClassCollection.GetCustomClassNames())
            {
                if (App.Settings.CurrentWorksheetSettings.CustomClassNames.ContainsKey(kvp.Key))
                    App.Settings.CurrentWorksheetSettings.CustomClassNames[kvp.Key] = kvp.Value;
                else
                    App.Settings.CurrentWorksheetSettings.CustomClassNames.Add(kvp.Key, kvp.Value);
            }
            App.Settings.Persist();
        }

        private void SelectVertexInSpreadsheet(Vertex vertex)
        {
            spreadsheet.SetActiveSheet(string.IsNullOrEmpty(vertex.ExternalWorksheetName) 
                ? _generator.ActiveWorksheet
                : vertex.ExternalWorksheetName);

            if (vertex is CellVertex cell)
            {
                FlashAndSelectSpreadsheetCells(spreadsheet.ActiveSheet.Range[cell.StringAddress]);
                // scroll into view
                spreadsheet.ActiveGrid.CurrentCell.MoveCurrentCell(cell.Address.row, cell.Address.col);
            }
            else if (vertex is RangeVertex rangeVertex)
            {
                FlashAndSelectSpreadsheetCells(rangeVertex.CellsInRange);
                spreadsheet.ActiveGrid.CurrentCell
                    .MoveCurrentCell(rangeVertex.StartAddress.row, rangeVertex.StartAddress.column);
                spreadsheet.ActiveGrid.CurrentCell
                    .MoveCurrentCell(rangeVertex.EndAddress.row, rangeVertex.EndAddress.column);
            }
        }

        private void SelectClassVerticesInSpreadsheet(Class @class)
        {
            // get the most common worksheet (current worksheet is null)
            var mostCommonWorksheetList = @class.Vertices
                .GroupBy(v => v.ExternalWorksheetName)
                .OrderByDescending(gp => gp.Count())
                .Take(1)
                .Select(gp => gp.Key)
                .ToList();
            if (mostCommonWorksheetList.Count == 0) return;
            var mostCommonWorksheet = mostCommonWorksheetList[0];

            // navigate to the most common worksheet
            spreadsheet.SetActiveSheet(mostCommonWorksheet ?? _generator.ActiveWorksheet);

            var ranges = @class.Vertices
                .Where(v => v.WorksheetName == mostCommonWorksheet)
                .Select(vertex => spreadsheet.ActiveSheet.Range[vertex.StringAddress])
                .ToArray();

            FlashAndSelectSpreadsheetCells(ranges);

            spreadsheet.ActiveGrid.CurrentCell
                .MoveCurrentCell(ranges.Min(r => r.Row), ranges.Min(r => r.Column));
            spreadsheet.ActiveGrid.CurrentCell
                .MoveCurrentCell(ranges.Max(r => r.LastRow), ranges.Max(r => r.LastColumn));
        }

        private void FlashAndSelectSpreadsheetCells(params IRange[] ranges)
        {
            // highlight selected cell(s) yellow for one second
            var originalBgColors = ranges
                .SelectMany(c => c.Cells)
                .ToDictionary(c => c, c => c.CellStyle.Color);

            if (spreadsheet.Tag != null) return;
            spreadsheet.Tag = true;
            foreach (var cell in originalBgColors)
                cell.Key.CellStyle.Color = Color.Yellow;

            spreadsheet.ActiveGrid.InvalidateCells();
            Task.Factory.StartNew(() => Thread.Sleep(1000))
                .ContinueWith(t =>
                {
                    foreach (var cell in originalBgColors)
                        cell.Key.CellStyle.Color = cell.Value;

                    spreadsheet.ActiveGrid.InvalidateCells();
                    spreadsheet.Tag = null;

                    spreadsheet.ActiveGrid.SelectionController.ClearSelection();
                    foreach (var range in ranges)
                        spreadsheet.ActiveGrid.SelectionController.AddSelection(range.ConvertExcelRangeToGridRange());

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

        private void SelectVertexInCode(Vertex vertex)
        {
            if (_generator.Code == null || vertex.Class == null || generateCodeTab.IsSelected == false) return;

            var code = codeTextBox.Text;

            var indexOfClassNameInCode = code.IndexOf($"class {vertex.Class.Name}", StringComparison.InvariantCulture);
            if (indexOfClassNameInCode == -1) return;

            var indexOfVariableInCode = code.IndexOf(vertex.Name, indexOfClassNameInCode, StringComparison.InvariantCulture);
            if (indexOfVariableInCode == -1) return;

            var line = codeTextBox.Document.GetLineByOffset(indexOfVariableInCode);
            codeTextBox.ScrollTo(line.LineNumber, line.Length);
            codeTextBox.Select(indexOfVariableInCode, vertex.Name.Length);
        }

        public void StyleBackground(Color color, IStyle cellStyle)
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
        /// <param name="styleBorder">Border styling function</param>
        /// <param name="styleCell">Cell styling function</param>
        public void ColorSpreadsheetCells(IEnumerable<Vertex> vertices, Action<CellVertex, IBorders> styleBorder,
            Action<CellVertex, IStyle> styleCell = null)
        {
            if (spreadsheet.ActiveSheet.Name != _generator.ActiveWorksheet)
                spreadsheet.SetActiveSheet(_generator.ActiveWorksheet);

            var vertexList = vertices.ToList();
            foreach (var rangeVertex in vertexList.OfType<RangeVertex>())
            {
                ColorSpreadsheetCells(rangeVertex.CellsInRange.Select(c => new CellVertex(c)), styleBorder, styleCell);
            }

            foreach (var cellVertex in vertexList.OfType<CellVertex>())
            {
                var cellRange = spreadsheet.ActiveSheet.Range[cellVertex.StringAddress];
                styleCell?.Invoke(cellVertex, cellRange.CellStyle);
                styleBorder?.Invoke(cellVertex, cellRange.Borders);
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
