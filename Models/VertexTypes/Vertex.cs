// Thesis - An Excel to code converter
// Copyright (C) 2019 Jacob Zhang
// 
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
// 
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
// 
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <https://www.gnu.org/licenses/>.

using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using Syncfusion.UI.Xaml.Diagram;

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