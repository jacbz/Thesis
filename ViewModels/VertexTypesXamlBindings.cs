﻿using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using Thesis.Models.VertexTypes;

namespace Thesis.Models.VertexTypes
{
    partial class Vertex
    {
        public bool IsRangeVertex => this is RangeVertex;
        public bool IsCellVertex => this is CellVertex;

        public XamlVertexType LayoutVertexType => IsRangeVertex
            ? XamlVertexType.Range
            : (XamlVertexType) Enum.Parse(typeof(XamlVertexType), ((CellVertex) this).NodeType.ToString());

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    partial class CellVertex
    {
        public string CellTypeString => CellType.ToString();
    }

    public enum XamlVertexType
    {
        Range,
        Formula,
        OutputField,
        Constant,
        None
    }
}
