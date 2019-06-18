using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;

namespace Thesis
{
    enum CellType
    {
        Formula,
        Bool,
        Number,
        Date,
        Text
    }

    class Vertex
    {
        public object Value { get; set; }
        public string Formula { get; set; }
        public CellType Type { get; set; }
        public int[] CellIndex { get; set; }
        public string Address { get; set; }
        public HashSet<string> Parents { get; set; }
        public HashSet<string> Children { get; set; }

        // Layouting constants
        public string Color
        {
            get
            {
                switch (Type)
                {
                    case CellType.Formula:
                        return "#8e44ad";
                    case CellType.Number:
                        return "#2980b9";
                    case CellType.Bool:
                        return "#27ae60";
                    case CellType.Date:
                        return "#f39c12";
                    default:
                        return "#2c3e50";
                }
            }
        }

        public Vertex(string s)
        {
            Address = s;
            Type = CellType.Text;
            CellIndex = new int[2] { 1, 1 };
            Parents = new HashSet<string>();
            Children = new HashSet<string>();
        }

        public Vertex(IRange cell)
        {
            Value = cell.Value;
            Formula = cell.Formula;
            Type = GetCellType(cell);
            CellIndex = new int[2] { cell.Row, cell.Column };
            Address = cell.AddressLocal;            
            Parents = new HashSet<string>();
            Children = new HashSet<string>();
        }

        public CellType GetCellType(IRange cell)
        {
            if (cell.HasFormula) return CellType.Formula;
            if (cell.HasBoolean) return CellType.Bool;
            if (cell.HasNumber) return CellType.Number;
            if (cell.HasDateTime) return CellType.Date;
            return CellType.Text;
        }

        public override string ToString()
        {
            return $"{Address},{Type.ToString()}: {Value}";
        }
    }
}
