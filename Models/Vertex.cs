using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using Irony.Parsing;
using Syncfusion.UI.Xaml.Diagram;
using Syncfusion.XlsIO;
using Thesis.Models;
using Thesis.Models.CodeGenerators;

namespace Thesis
{
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

    public class Vertex : INotifyPropertyChanged
    {
        private bool _include;

        public object Value { get; set; }
        public string Formula { get; set; }
        public Label Label { get; set; }
        public CellType Type { get; set; }
        public (int row, int col) Address { get; set; }
        public string StringAddress { get; }
        public HashSet<Vertex> Parents { get; set; }
        public HashSet<Vertex> Children { get; set; }
        public ParseTreeNode ParseTree { get; set; }

        public bool Include
        {
            get => _include;
            set
            {
                _include = value;
                OnPropertyChanged();
            }
        }

        public NodeType NodeType => ParseTree != null 
            ? (Parents.Count == 0 ? NodeType.OutputField : NodeType.Formula) 
            : (string)Value != "" 
                ? NodeType.Constant 
                : NodeType.None;

        public NodeViewModel Node { get; set; }
        public GeneratedClass Class { get; set; }
        // ReSharper disable once UnusedMember.Global
        public string CellTypeString => Type.ToString();
        public string NameInCode => Label.VariableName;

        public event PropertyChangedEventHandler PropertyChanged;

        public Vertex(IRange cell)
        {
            Type = GetCellType(cell);
            Value = cell.DisplayText;
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
            if (cell.HasDateTime || cell.HasFormulaBoolValue) return CellType.Date;
            if (cell.HasString || cell.HasFormulaStringValue) return CellType.Text;
            return CellType.Unknown;
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
            return $"{StringAddress},{Type.ToString()}: {Value}";
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
}