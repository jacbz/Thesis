using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using Irony.Parsing;
using Syncfusion.UI.Xaml.Diagram;
using Syncfusion.XlsIO;

namespace Thesis.Models
{
    public class Vertex : INotifyPropertyChanged
    {
        public dynamic Value { get; set; }
        public string DisplayValue { get; set; }
        public string Formula { get; set; }
        public Label Label { get; set; }
        public string VariableName { get; set; }
        public (int row, int col) Address { get; set; }
        public string StringAddress { get; }
        public HashSet<Vertex> Parents { get; set; }
        public HashSet<Vertex> Children { get; set; }
        public ParseTreeNode ParseTree { get; set; }

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
            Parents.Count > 0
                ? Children.Count == 0
                    ? NodeType.Constant
                    : NodeType.Formula
                : Children.Count > 0
                    ? NodeType.OutputField
                    : NodeType.None;

        public NodeViewModel Node { get; set; }
        public GeneratedClass Class { get; set; }

        // used in XAML binding
        public string CellTypeString => CellType.ToString();

        public event PropertyChangedEventHandler PropertyChanged;

        public Vertex(IRange cell)
        {
            CellType = GetCellType(cell);
            Value = GetCellValue(cell);
            DisplayValue = cell.DisplayText;
            Formula = cell.HasFormula ? FormatFormula(cell.Formula) : null;
            Address = (cell.Row, cell.Column);
            StringAddress = cell.AddressLocal;
            Parents = new HashSet<Vertex>();
            Children = new HashSet<Vertex>();
            Include = true;

            if (cell.HasFormula)
                ParseTree = XLParser.ExcelFormulaParser.Parse(Formula);
        }

        public CellType GetCellType(IRange cell)
        {
            if (cell.HasBoolean || cell.HasFormulaBoolValue) return CellType.Bool;
            if (cell.HasNumber || cell.HasFormulaNumberValue) return CellType.Number;
            if (cell.HasDateTime || cell.HasFormulaDateTime) return CellType.Date;
            if (cell.HasString || cell.HasFormulaStringValue) return CellType.Text;
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

        public HashSet<Vertex> GetReachableVertices()
        {
            var vertices = new HashSet<Vertex> {this};
            foreach (var v in Children) vertices.UnionWith(v.GetReachableVertices());
            return vertices;
        }

        public override string ToString()
        {
            return $"{StringAddress},{CellType.ToString()}: {DisplayValue}";
        }

        public override int GetHashCode()
        {
            return StringAddress.GetHashCode();
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