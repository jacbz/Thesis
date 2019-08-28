﻿using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using Irony.Parsing;
using Syncfusion.UI.Xaml.Diagram;
using Syncfusion.XlsIO;

namespace Thesis.Models.VertexTypes
{
    public abstract class Vertex : INotifyPropertyChanged
    {
        public string StringAddress { get; set; }

        public (string worksheet, string address) GlobalAddress => IsExternal
            ? (ExternalWorksheetName, StringAddress)
            : (null, StringAddress);
        public HashSet<Vertex> Parents { get; set; }
        public HashSet<Vertex> Children { get; set; }
        public string VariableName { get; set; } // for ranges, this will be either a name (named range), or address (e.g. A1:C3)

        public bool IsExternal => !string.IsNullOrEmpty(ExternalWorksheetName);
        public string ExternalWorksheetName { get; set; }
        public string WorksheetName => IsExternal ? ExternalWorksheetName : null;

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

        public NodeViewModel Node { get; set; }
        public Class Class { get; set; }


        public event PropertyChangedEventHandler PropertyChanged;

        public Vertex()
        {
            Parents = new HashSet<Vertex>();
            Children = new HashSet<Vertex>();
            Include = true;
        }

        public void MarkAsExternal(string worksheetName)
        {
            ExternalWorksheetName = worksheetName;
        }

        public void MarkAsExternal(string worksheetName, string variableName)
        {
            ExternalWorksheetName = worksheetName;
            VariableName = GenerateExternalVariableName(worksheetName, variableName);
        }

        public static string GenerateExternalVariableName(string worksheetName, string variableName)
        {
            return worksheetName.MakeNameVariableConform() + "_" + variableName.MakeNameVariableConform();
        }

        public HashSet<Vertex> GetReachableVertices(bool ignoreExternal = true)
        {
            var vertices = new HashSet<Vertex> { this };
            // exclude external vertices
            foreach (var v in Children)
            {
                if (ignoreExternal && v.IsExternal) continue;
                vertices.UnionWith(v.GetReachableVertices(ignoreExternal));
            }

            return vertices;
        }

        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}