using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Thesis.ViewModels;

namespace Thesis.Models
{
    public class ClassCollection
    {
        public List<Class> Classes { get; }

        public ClassCollection()
        {
            Classes = new List<Class>();
        }

        public static ClassCollection FromGraph(Graph graph)
        {
            ClassCollection classCollection = new ClassCollection();

            var vertexToOutputFieldVertices = graph.Vertices.ToDictionary(v => v, v => new HashSet<Vertex>());
            var rnd = new Random();

            foreach (var vertex in graph.GetOutputFields())
            {
                var reachableVertices = vertex.GetReachableVertices();
                foreach (var v in reachableVertices) vertexToOutputFieldVertices[v].Add(vertex);
                var newClass = new Class($"Class{vertex.StringAddress}", vertex, reachableVertices.ToList(), rnd);
                classCollection.Classes.Add(newClass);
            }

            var logItem = Logger.Log(LogItemType.Info, "Applying topological sort...", true);
            foreach (var generatedClass in classCollection.Classes)
            {
                generatedClass.Vertices.RemoveAll(v => vertexToOutputFieldVertices[v].Count > 1);
                generatedClass.TopologicalSort();
            }
            logItem.AppendElapsedTime();

            // vertices used by more than one class
            var sharedVertices = vertexToOutputFieldVertices
                .Where(kvp => kvp.Value.Count > 1)
                .Select(kvp => kvp.Key)
                .ToList();

            if (graph.ExternalVertices.Count > 0)
            {
                var externalClass = new Class("External", null, graph.ExternalVertices);
                externalClass.TopologicalSort();
                classCollection.Classes.Add(externalClass);
            }

            if (sharedVertices.Count > 0)
            {
                var globalClass = new Class("Global", null, sharedVertices);
                globalClass.TopologicalSort();
                classCollection.Classes.Add(globalClass);
            }

            if (graph.Vertices.Count + graph.ExternalVertices.Count != 
                classCollection.Classes.Sum(l => l.Vertices.Count))
                Logger.Log(LogItemType.Error, "Error creating classes; number of vertices does not match number of vertices in classes");

            return classCollection;
        }
    }
}
