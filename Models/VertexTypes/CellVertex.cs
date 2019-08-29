using Irony.Parsing;
using Syncfusion.XlsIO;

namespace Thesis.Models.VertexTypes
{
    /// <summary>
    /// A vertex which represents a spreadsheet cell.
    /// </summary>
    public partial class CellVertex : Vertex
    {
        public (int row, int col) Address { get; set; }
        public string StringAddress { get; set; }
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
                : Children.Count > 0
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
            return cell.DisplayText;
        }

        private string FormatFormula(string formula)
        {
            // for German formulas
            return formula.Replace(",", ".").Replace(";", ",").Replace("$", "");
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
