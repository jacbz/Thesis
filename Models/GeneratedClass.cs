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
        public static readonly Color StaticColor = ColorTranslator.FromHtml("#CFD8DC");
        public static readonly Color ExternalColor = ColorTranslator.FromHtml("#AFEEEE");

        public GeneratedClass(string name, Vertex outputVertex, List<Vertex> vertices, Random rnd = null)
        {
            Name = name;
            OutputVertex = outputVertex;
            Vertices = vertices;
            Vertices.ForEach(v =>
            {
                v.VariableName = string.IsNullOrWhiteSpace(v.VariableName)
                    ? "_" + v.StringAddress 
                    : v.VariableName.MakeNameVariableConform();
                v.Class = this;
            });
            Color = name == "External" ? ExternalColor
                : rnd == null ? StaticColor
                : Color.FromArgb(rnd.Next(180, 256), rnd.Next(180, 256), rnd.Next(180, 256));
        }

        public string Name { get; set; }
        public bool IsStaticClass => OutputVertex == null;
        public Vertex OutputVertex { get; set; }
        public List<Vertex> Vertices { get; set; }
        public Color Color { get; set; }

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