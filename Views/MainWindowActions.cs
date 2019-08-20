using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
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
            spreadsheet.ActiveGrid.CurrentCell.MoveCurrentCell(vertex.Address.row, vertex.Address.col);
        }

        private void SelectVertexInOutputListView(Vertex vertex)
        {
            if (vertex.NodeType == NodeType.OutputField)
            {
                outputFieldsListView.ScrollIntoView(vertex);
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
            ColorSpreadsheetCells(allVertices, filteredVertices, vertex =>
            {
                switch (vertex.Label.Type)
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
            spreadsheet.ActiveGrid.InvalidateCells();
        }

        /// <summary>
        /// Color spreadsheet cells, given their vertices.
        /// </summary>
        /// <param name="allVertices">For all cells, the background color will be determined using cellColorFunc</param>
        /// <param name="filteredVertices">Only these will get a border (correspondig to the node type)</param>
        /// <param name="cellColorFunc">Determines cell background color for all cells.</param>
        public void ColorSpreadsheetCells(List<Vertex> allVertices, List<Vertex> filteredVertices, Func<Vertex, Color> cellColorFunc)
        {
            foreach (var vertex in allVertices)
            {
                var range = spreadsheet.ActiveSheet.Range[vertex.StringAddress];
                range.CellStyle.Color = cellColorFunc(vertex);
                range.CellStyle.Font.RGBColor = cellColorFunc(vertex).GetTextColor();
                if (!filteredVertices.Contains(vertex))
                {
                    range.Borders.LineStyle = ExcelLineStyle.None;
                }
                else
                {
                    range.Borders.ColorRGB = vertex.GetNodeTypeColor();
                    range.Borders.LineStyle = ExcelLineStyle.Thick;
                }
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
