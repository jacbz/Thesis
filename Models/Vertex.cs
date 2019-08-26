using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using Irony.Parsing;
using Syncfusion.UI.Xaml.Diagram;
using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation;

namespace Thesis.Models
{
    public class Vertex : INotifyPropertyChanged
    {
        public dynamic Value { get; set; }
        public string DisplayValue { get; set; }
        public string Formula { get; set; }
        public Label Label { get; set; }
        public string VariableName { get; set; }
        public string ExternalWorksheetName { get; set; }
        public (int row, int col) Address { get; set; }
        public string StringAddress { get; set; }
        public HashSet<Vertex> Parents { get; set; }
        public HashSet<Vertex> Children { get; set; }
        public ParseTreeNode ParseTree { get; set; }
        public bool IsSpreadsheetCell => !string.IsNullOrEmpty(StringAddress);

        private bool _include;
        public bool Include
        {
            get => _include;
            set
            {
                _include = value;
                OnPropertyChanged();
            }
        }

        public CellType CellType { get; set; }
        public NodeType NodeType =>
            string.IsNullOrEmpty(ExternalWorksheetName)
            ? Parents.Count > 0
                ? Children.Count == 0
                    ? NodeType.Constant
                    : NodeType.Formula
                : Children.Count > 0
                    ? NodeType.OutputField
                    : NodeType.None
            : NodeType.External;

        public NodeViewModel Node { get; set; }
        public Class Class { get; set; }

        // used in XAML binding
        public string CellTypeString => CellType.ToString();

        public event PropertyChangedEventHandler PropertyChanged;

        public Vertex()
        {
            Parents = new HashSet<Vertex>();
            Children = new HashSet<Vertex>();
            Include = true;
        }

        public Vertex(IRange cell) : this()
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

        public static Vertex CreateNamedRangeVertex(string namedRangeName, string namedRangeAddress)
        {
            var formula = "=" + namedRangeAddress;
            return new Vertex
            {
                CellType = CellType.Range,
                Value = null,
                DisplayValue = "",
                Formula = formula,
                Address = (0, 0),
                StringAddress = null,
                ParseTree = XLParser.ExcelFormulaParser.Parse(formula),
                VariableName = namedRangeName
            };
        }

        public void MarkAsExternal(string worksheetName, string variableName)
        {
            ExternalWorksheetName = worksheetName;
            VariableName = GenerateExternalVariableName(worksheetName, variableName);
        }

        public static string GenerateExternalVariableName(string worksheetName, string variableName)
        {
            return worksheetName.MakeNameVariableConform() + "_" + variableName;
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

        public HashSet<Vertex> GetReachableVertices(bool ignoreExternal = true)
        {
            var vertices = new HashSet<Vertex> { this };
            // exclude external vertices
            foreach (var v in Children)
            {
                if (ignoreExternal && v.NodeType == NodeType.External) continue;
                vertices.UnionWith(v.GetReachableVertices(ignoreExternal));
            }

            return vertices;
        }

        public override string ToString()
        {
            return $"{StringAddress},{CellType.ToString()}: {DisplayValue}";
        }

        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public enum CellType
    {
        Bool,
        Number,
        Date,
        Text,
        Range,
        Unknown
    }

    public enum NodeType
    {
        Formula,
        OutputField,
        Constant,
        External,
        None
    }
}