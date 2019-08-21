using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using Thesis.Models.CodeGenerators;
using Thesis.ViewModels;

namespace Thesis.Models
{
    public class GeneratedClass
    {
        private static Color _sharedColor = ColorTranslator.FromHtml("#CFD8DC");
        public GeneratedClass(string name, Vertex outputVertex, List<Vertex> vertices, Random rnd)
        {
            Name = name;
            OutputVertex = outputVertex;
            Vertices = vertices;
            Vertices.ForEach(v =>
            {
                v.VariableName = string.IsNullOrWhiteSpace(v.VariableName) ? "_" + v.StringAddress : v.VariableName;
                v.Class = this;
            });
            Color = outputVertex == null
                ? _sharedColor
                : Color.FromArgb(rnd.Next(180, 256), rnd.Next(180, 256), rnd.Next(180, 256));
        }

        public string Name { get; set; }
        public bool IsSharedClass => OutputVertex == null;
        public Vertex OutputVertex { get; set; }
        public List<Vertex> Vertices { get; set; }
        public Color Color { get; set; }

        // Implementation of Kahn's algorithm
        public void TopologicalSort()
        {
            var sortedVertices = new List<Vertex>();
            var verticesWithNoParents = Vertices.Where(v => v.Parents.Count == 0).ToHashSet();
            var edges = new HashSet<(Vertex from, Vertex to)>();

            foreach (var (vertex, child) in Vertices.SelectMany(vertex =>
                vertex.Children.Select(child => (vertex, child)))) edges.Add((vertex, child));

            while (verticesWithNoParents.Count > 0)
            {
                var n = verticesWithNoParents.OrderBy(v => v.Children.Count).First();
                verticesWithNoParents.Remove(n);
                // do not re-add deleted vertices
                if (Vertices.Contains(n)) sortedVertices.Add(n);
                foreach (var m in n.Children)
                {
                    edges.Remove((n, m));
                    if (edges.Count(x => x.to == m) == 0) verticesWithNoParents.Add(m);
                }
            }

            if (edges.Count != 0)
                Logger.Log(LogItemType.Error, "Error during topological sort. Graph has at least one cycle?");
            Vertices = sortedVertices;
        }
    }
}