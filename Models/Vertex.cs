using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using Syncfusion.UI.Xaml.Diagram;
using Syncfusion.XlsIO;
using Thesis.Models;

namespace Thesis
{
    public enum CellType
    {
        Formula,
        Bool,
        Number,
        Date,
        Text
    }

    public enum NodeType
    {
        Formula,
        OutputField,
        Constant
    }

    public class Vertex : INotifyPropertyChanged
    {
        private bool _include;

        public object Value { get; set; }
        public string Formula { get; set; }
        public string Label { get; set; }
        public CellType Type { get; set; }
        public int[] CellIndex { get; set; }
        public string Address { get; }
        public HashSet<Vertex> Parents { get; set; }
        public HashSet<Vertex> Children { get; set; }

        public bool Include
        {
            get => _include;
            set
            {
                _include = value;
                OnPropertyChanged();
            }
        }

        public NodeType NodeType => Type == CellType.Formula
            ? Parents.Count == 0 ? NodeType.OutputField : NodeType.Formula
            : NodeType.Constant;

        public NodeViewModel Node { get; set; }
        public GeneratedClass Class { get; set; }
        public bool HasLabel => !string.IsNullOrEmpty(Label);

        public event PropertyChangedEventHandler PropertyChanged;

        public Vertex(IRange cell)
        {
            Type = GetCellType(cell);
            Value = Type == CellType.Formula ? FormatFormula(cell.Value) : cell.Value;
            Formula = cell.Formula != null ? FormatFormula(cell.Formula) : null;
            CellIndex = new[] { cell.Row, cell.Column };
            Address = cell.AddressLocal;
            Parents = new HashSet<Vertex>();
            Children = new HashSet<Vertex>();
            Include = true;
        }

        private string FormatFormula(string formula)
        {
            return formula.Replace(",", ".").Replace(";", ",").Replace("$", "");
        }

        public static CellType GetCellType(IRange cell)
        {
            if (cell.HasFormula)
                return CellType.Formula;
            if (cell.HasBoolean) return CellType.Bool;
            if (cell.HasNumber) return CellType.Number;
            if (cell.HasDateTime) return CellType.Date;
            return CellType.Text;
        }

        public HashSet<Vertex> GetReachableVertices()
        {
            var vertices = new HashSet<Vertex> {this};
            foreach (var v in Children) vertices.UnionWith(v.GetReachableVertices());
            return vertices;
        }

        public override string ToString()
        {
            return $"{Address},{Type.ToString()}: {Value}";
        }

        public override int GetHashCode()
        {
            return Address.GetHashCode();
        }

        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}