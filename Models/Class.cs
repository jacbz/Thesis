using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using Thesis.Models.VertexTypes;
using Thesis.ViewModels;

namespace Thesis.Models
{
    public class Class
    {
        public static readonly Color GlobalColor = ColorTranslator.FromHtml("#CFD8DC");
        public static readonly Color ExternalColor = ColorTranslator.FromHtml("#c4f7ed");

        public string Name { get; set; }
        public string DefaultName { get; }
        public bool IsStaticClass => OutputVertex == null;
        public Vertex OutputVertex { get; set; }
        public List<Vertex> Vertices { get; set; }
        public Color Color { get; set; }

        public Class(string name, string defaultName, Vertex outputVertex, List<Vertex> vertices, Color? color = null)
        {
            Name = name;
            DefaultName = defaultName;
            OutputVertex = outputVertex;
            Vertices = vertices;
            foreach (var vertex in Vertices)
            {
                vertex.Name = vertex is CellVertex cellVertex 
                    ? string.IsNullOrWhiteSpace(vertex.Name)
                        ? "_" + cellVertex.StringAddress
                        : vertex.Name.MakeNameVariableConform()
                    : vertex.Name;
                vertex.Class = this;
            }

            Color = color.GetValueOrDefault();
        }

        // Implementation of Kahn's algorithm
        public void TopologicalSort()
        {
            var sortedVertices = new List<Vertex>();
            // only consider parents from the same class
            var verticesWithNoParents = Vertices.Where(v => v.Parents.Count(p => p.Class == this) == 0).ToHashSet();

            var edges = new HashSet<(Vertex from, Vertex to)>();
            foreach (var (vertex, child) in Vertices
                .SelectMany(vertex => vertex.Children.Select(child => (vertex, child))))
                edges.Add((vertex, child));

            while (verticesWithNoParents.Count > 0)
            {
                // get first vertex that has the lowest number of children
                var currentVertex = verticesWithNoParents.OrderBy(v => v.Children.Count).First();

                verticesWithNoParents.Remove(currentVertex);
                // do not re-add deleted vertices
                if (Vertices.Contains(currentVertex)) sortedVertices.Add(currentVertex);

                foreach (var childOfCurrentVertex in currentVertex.Children)
                {
                    edges.Remove((currentVertex, childOfCurrentVertex));
                    if (edges.Count(x => x.to == childOfCurrentVertex) == 0)
                        verticesWithNoParents.Add(childOfCurrentVertex);
                }
            }

            if (edges.Count != 0)
                Logger.Log(LogItemType.Error, "Error during topological sort. Graph has at least one cycle?");

            // reverse order
            sortedVertices.Reverse();
            Vertices = sortedVertices;
        }
    }
}