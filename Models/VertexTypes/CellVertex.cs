using System.Globalization;
using System.Threading;
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
        public Label Label { get; set; }
        public ParseTreeNode ParseTree { get; set; }

        public CellType CellType { get; set; }

        public NodeType NodeType =>
            Parents.Count > 0
                ? Children.Count == 0
                    ? NodeType.Constant
                    : NodeType.Formula
                : Children.Count > 0 || ParseTree != null
                    ? NodeType.OutputField
                    : NodeType.None;

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

    public enum NodeType
    {
        Formula,
        OutputField,
        Constant,
        None
    }
}
