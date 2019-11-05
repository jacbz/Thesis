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

using Irony.Parsing;
using Syncfusion.XlsIO;
using Thesis.ViewModels;

namespace Thesis.Models.VertexTypes
{
    /// <summary>
    /// A vertex which represents a spreadsheet cell.
    /// </summary>
    public class CellVertex : Vertex
    {
        public (int row, int col) Address { get; set; }
        public (string worksheet, string address) GlobalAddress => IsExternal
            ? (ExternalWorksheetName, StringAddress)
            : (null, StringAddress);
        public dynamic Value { get; set; }
        public string DisplayValue { get; set; }
        public string Formula { get; set; }
        public NameGenerator.Region Region { get; set; }
        public ParseTreeNode ParseTree { get; set; }

        public CellType CellType { get; set; }

        public Classification Classification =>
            Parents.Count > 0
                ? Children.Count == 0
                    ? string.IsNullOrEmpty(Value.ToString())
                      ? Classification.InputField
                      : Classification.Constant
                    : Classification.Formula
                : Children.Count > 0 || ParseTree != null
                    ? Classification.OutputField
                    : Classification.None;

        public CellVertex(IRange cell)
        {
            CellType = GetCellType(cell);
            Value = GetCellValue(cell);
            DisplayValue = cell.DisplayText;
            Formula = cell.HasFormula ? FormatFormula(cell.Formula) : null;
            Address = (cell.Row, cell.Column);
            StringAddress = cell.AddressLocal;

            if (cell.HasFormula)
                ParseTree = XLParser.ExcelFormulaParser.Parse(Formula);
        }

        public CellVertex(IRange cell, string name) : this(cell)
        {
            Name = new Name(name, StringAddress);
        }

        public CellType GetCellType(IRange cell)
        {
            if (cell.HasBoolean || cell.HasFormulaBoolValue) return CellType.Bool;
            if (cell.HasNumber || cell.HasFormulaNumberValue) return CellType.Number;
            if (cell.HasDateTime || cell.HasFormulaDateTime) return CellType.Date;
            if (cell.HasFormulaErrorValue) return CellType.Error;
            if (cell.HasString || cell.HasFormulaStringValue ||
                !string.IsNullOrEmpty(cell.DisplayText)) return CellType.Text;
            return CellType.Unknown;
        }

        public object GetCellValue(IRange cell)
        {
            if (cell.HasBoolean)
                return cell.Boolean;
            if (cell.HasFormulaBoolValue)
                return cell.FormulaBoolValue;
            if (cell.HasNumber)
                return cell.Number;
            if (cell.HasFormulaNumberValue)
                return cell.FormulaNumberValue;
            if (cell.HasDateTime)
                return cell.DateTime;
            if (cell.HasFormulaDateTime)
                return cell.FormulaDateTime;
            if (cell.HasFormulaErrorValue)
                return cell.FormulaErrorValue;
            return cell.DisplayText;
        }

        private string FormatFormula(string formula)
        {
            if (Formatter.CurrentCultureInfo.NumberFormat.NumberDecimalSeparator == ",")
                formula = formula.Replace(",", ".").Replace(";", ",");
            return formula.Replace("$", "");
        }

        public override string ToString()
        {
            return $"{StringAddress},{CellType.ToString()}: {DisplayValue}";
        }
    }

    public enum CellType
    {
        Bool,
        Number,
        Date,
        Text,
        Error,
        Unknown
    }

    public enum Classification
    {
        InputField,
        Constant,
        Formula,
        OutputField,
        None
    }
}
