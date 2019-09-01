using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Media;
using Thesis.Models.VertexTypes;
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

        public static ClassCollection FromGraph(Graph graph, Dictionary<string, string> customClassNames = null)
        {
            ClassCollection classCollection = new ClassCollection();

            var vertexToOutputFieldVertices = graph.Vertices.ToDictionary(v => v, v => new HashSet<CellVertex>());
            
            // output fields without children: formulas such as =TODAY()
            var outputFieldsWithoutChildren = graph.GetOutputFields().Where(v => v.Children.Count == 0).ToList();

            foreach (var vertex in graph.GetOutputFields().Except(outputFieldsWithoutChildren))
            {
                var reachableVertices = vertex.GetReachableVertices();
                foreach (var v in reachableVertices)
                    vertexToOutputFieldVertices[v].Add(vertex);

                var defaultName = $"Class{vertex.StringAddress}";
                var className = defaultName;
                if (customClassNames != null && customClassNames.TryGetValue(className, out var customName))
                    className = customName;
                var newClass = new Class(className, defaultName, vertex, reachableVertices.ToList());
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

            if (outputFieldsWithoutChildren.Count > 0)
            {
                var defaultName = "OutputFieldsWithoutChildren";
                var className = defaultName;
                if (customClassNames != null && customClassNames.TryGetValue(className, out var customName))
                    className = customName;
                classCollection.Classes.Add(new Class(className, defaultName,null, new List<Vertex>(outputFieldsWithoutChildren)));
            }

            if (graph.ExternalVertices.Count > 0)
            {
                var defaultName = "External";
                var className = defaultName;
                if (customClassNames != null && customClassNames.TryGetValue(className, out var customName))
                    className = customName;
                var externalClass = new Class(className, defaultName, null, graph.ExternalVertices, Class.ExternalColor);
                externalClass.TopologicalSort();
                classCollection.Classes.Add(externalClass);
            }

            if (sharedVertices.Count > 0)
            {
                var defaultName = "Global";
                var className = defaultName;
                if (customClassNames != null && customClassNames.TryGetValue(className, out var customName))
                    className = customName;
                var globalClass = new Class(className, defaultName, null, sharedVertices, Class.GlobalColor);
                globalClass.TopologicalSort();
                classCollection.Classes.Add(globalClass);
            }

            // generate colors
            var classesWithoutColors = classCollection.Classes.Where(c => c.Color == default).ToArray();
            var colors = Formatter.GenerateNDistinctColors(classesWithoutColors.Length);
            for (int i = 0; i < classesWithoutColors.Length; i++)
                classesWithoutColors[i].Color = colors[i];

            if (graph.Vertices.Count + graph.ExternalVertices.Count != 
                classCollection.Classes.Sum(l => l.Vertices.Count))
                Logger.Log(LogItemType.Error, "Error creating classes; number of vertices does not match number of vertices in classes");

            return classCollection;
        }

        public Dictionary<string, string> GetCustomClassNames()
        {
            return Classes
                .Where(c => c.DefaultName != c.Name)
                .ToDictionary(c => c.DefaultName, c => c.Name);
        }
    }
}
