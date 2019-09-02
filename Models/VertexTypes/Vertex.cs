﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using Syncfusion.UI.Xaml.Diagram;
using Thesis.ViewModels;

namespace Thesis.Models.VertexTypes
{
    public abstract class Vertex : INotifyPropertyChanged
    {
        public HashSet<Vertex> Parents { get; set; }
        public HashSet<Vertex> Children { get; set; }
        private string _name;
        public string Name { get => _name; set { _name = value; OnPropertyChanged(); } }
        public string StringAddress { get; set; }
        public bool IsExternal => !string.IsNullOrEmpty(ExternalWorksheetName);
        public string ExternalWorksheetName { get; set; }
        public string WorksheetName => IsExternal ? ExternalWorksheetName : null;
        private bool _include;
        public bool Include { get => _include; set { _include = value; OnPropertyChanged(); } }

        public NodeViewModel Node { get; set; }
        public Class Class { get; set; }

        protected Vertex()
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
            Name = GenerateExternalVariableName(worksheetName, variableName);
        }

        public static string GenerateExternalVariableName(string worksheetName, string variableName)
        {
            return worksheetName.MakeNameVariableConform() + "_" + variableName.MakeNameVariableConform();
        }

        public HashSet<Vertex> GetReachableVertices(bool ignoreExternal = true)
        {
            var vertices = new HashSet<Vertex> { this };

            // exclude external vertices
            foreach (var vertex in Children)
            {
                if (ignoreExternal && vertex.IsExternal) continue;
                vertices.UnionWith(vertex.GetReachableVertices(ignoreExternal));
            }

            return vertices;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}