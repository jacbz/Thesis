using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Thesis.ViewModels;

namespace Thesis.Models
{
    public class GeneratedClass
    {
        public string Name { get; set; }
        public Vertex OutputVertex { get; set; }
        public List<Vertex> Vertices { get; set; }
        public Color Color { get; set; }

        public GeneratedClass(string name, Vertex outputVertex, List<Vertex> vertices, Random rnd)
        {
            Name = name;
            OutputVertex = outputVertex;
            Vertices = vertices;
            Vertices.ForEach(v => v.Class = this);
            Color = Color.FromArgb(rnd.Next(164, 256), rnd.Next(164, 256), rnd.Next(164, 256));
        }

        // Implementation of Kahn's algorithm
        public void TopologicalSort()
        {
            var sortedVertices = new List<Vertex>();
            var verticesWithNoParents = Vertices.Where(v => v.Parents.Count == 0).ToHashSet();
            var edges = new HashSet<(Vertex, Vertex)>();
            foreach (var (vertex, child) in Vertices.SelectMany(vertex => vertex.Children.Select(child => (vertex, child))))
            {
                edges.Add((vertex, child));
            }

            while (verticesWithNoParents.Count > 0)
            {
                Vertex n = verticesWithNoParents.OrderBy(v => v.Children.Count).First();
                verticesWithNoParents.Remove(n);
                // do not re-add deleted vertices
                if (Vertices.Contains(n)) sortedVertices.Add(n);
                foreach (Vertex m in n.Children)
                {
                    edges.Remove((n, m));
                    if (edges.Count(x => x.Item2 == m) == 0)
                    {
                        verticesWithNoParents.Add(m);
                    }
                }
            }

            if (edges.Count != 0)
            {
                Logger.Log(LogItemType.Error, "Error during topological sort. Graph has at least one cycle?");
            }
            Vertices = sortedVertices;
        }

    }
}
