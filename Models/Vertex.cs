﻿using Syncfusion.UI.Xaml.Diagram;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
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
        public object Value { get; set; }
        public string Formula { get; set; }
        public string Label { get; set; }
        public CellType Type { get; set; }
        public int[] CellIndex { get; set; }
        public string Address { get; set; }
        public HashSet<Vertex> Parents { get; set; }
        public HashSet<Vertex> Children { get; set; }
        private bool include;
        public bool Include
        {
            get => include;
            set
            {
                include = value;
                OnPropertyChanged("Include");
            }
        }
        public NodeType NodeType
        {
            get => Type == CellType.Formula ? (Parents.Count == 0 ? NodeType.OutputField : NodeType.Formula) : NodeType.Constant;
        }
        public NodeViewModel Node { get; set; }
        public GeneratedClass Class { get; set; }
        public bool HasLabel { get => !string.IsNullOrEmpty(Label); }

        public Vertex(IRange cell)
        {           
            Value = cell.Value;
            Formula = cell.Formula;
            Type = GetCellType(cell);
            CellIndex = new int[2] { cell.Row, cell.Column };
            Address = cell.AddressLocal;            
            Parents = new HashSet<Vertex>();
            Children = new HashSet<Vertex>();
            Include = true;
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
            var vertices = new HashSet<Vertex>() { this };
            foreach(Vertex v in Children)
            {
                vertices.UnionWith(v.GetReachableVertices());
            }
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

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
