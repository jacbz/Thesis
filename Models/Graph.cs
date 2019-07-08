using Irony.Parsing;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Thesis.Models;
using Thesis.ViewModels;

namespace Thesis
{
    public class Graph
    {
        public List<Vertex> Vertices { get; set; }

        // Preserve copy for filtering purposes
        public List<Vertex> AllVertices { get; set; }

        // For layouting purposes
        public List<int> PopulatedRows { get; set; }
        public List<int> PopulatedColumns { get; set; }

        public Graph(IRange cells)
        {
            var VerticesDict = new Dictionary<string, Vertex>();
            Vertices = new List<Vertex>();

            foreach(var cell in cells.Cells)
            {
                var vertex = new Vertex(cell);

                VerticesDict.Add(vertex.Address, vertex);
                Vertices.Add(vertex);
            }
            Logger.Log(LogItemType.Info, $"Adding {Vertices.Count} vertices...");

            var allAddresses = Vertices.Select(x => x.Address);

            foreach (var vertex in Vertices)
            {
                if (vertex.Type != CellType.Formula) continue;
                var formula = vertex.Formula.Replace(",", ".").Replace(";", ",").Replace("$", "");

                // TODO: create "external" vertices (SheetNameQuotedToken etc)
                //

                try
                {
                    var parseTree = XLParser.ExcelFormulaParser.Parse(formula);
                    foreach(var cell in GetListOfReferencedCells(parseTree))
                    {
                        vertex.Children.Add(VerticesDict[cell]);
                        VerticesDict[cell].Parents.Add(vertex);
                    }
                }
                catch (Exception ex)
                {
                    Logger.Log(LogItemType.Error, $"Error processing formula in {vertex.Address} ({formula}): {ex.Message}");
                    continue;
                }

            }
            TransitiveFilter(GetOutputFields());
            Logger.Log(LogItemType.Info, $"Filtered for reachable vertices from output fields. {Vertices.Count} remaining");

            AllVertices = Vertices.ToList();
        }

        // recursively gets list of referenced cells from parse tree
        private IEnumerable<string> GetListOfReferencedCells(ParseTreeNode parseTree)
        {
            if (parseTree.ChildNodes.Count == 0 && parseTree.Term.Name == "CellToken")
            {
                yield return parseTree.Token.Text;
            }
            foreach(var child in parseTree.ChildNodes)
            {
                foreach(var cell in GetListOfReferencedCells(child))
                {
                    yield return cell;
                }
            }
        }

        public List<Vertex> GetOutputFields()
        {
            return Vertices.Where(v => v.IsOutputField).ToList();
        }

        public void Reset()
        {
            Vertices = AllVertices.ToList();
        }

        // Remove all vertices that are not transitively reachable from any vertex in the given list
        public void TransitiveFilter(List<Vertex> vertices)
        {
            Vertices = vertices.SelectMany(v => v.GetReachableVertices()).Distinct().ToList();

            PopulatedRows = Vertices.Select(v => v.CellIndex[0]).Distinct().ToList();
            PopulatedRows.Sort();
            PopulatedColumns = Vertices.Select(v => v.CellIndex[1]).Distinct().ToList();
            PopulatedColumns.Sort();
        }

        public List<GeneratedClass> GenerateClasses()
        {
            var classesList = new List<GeneratedClass>();

            var vertexToOutputFieldVertices = Vertices.ToDictionary(v => v, v => new HashSet<Vertex>());

            var rnd = new Random();

            foreach (var vertex in GetOutputFields())
            {
                var reachableVertices = vertex.GetReachableVertices();
                foreach(var v in reachableVertices)
                {
                    vertexToOutputFieldVertices[v].Add(vertex);
                }
                var newClass = new GeneratedClass($"Class{vertex.Address}", vertex, reachableVertices.ToList(), rnd);
                classesList.Add(newClass);
            }

            Logger.Log(LogItemType.Info, "Applying topological sort...");
            foreach (var generatedClass in classesList)
            {
                generatedClass.Vertices.RemoveAll(v => vertexToOutputFieldVertices[v].Count > 1);
                generatedClass.TopologicalSort();
            }

            var sharedVertices = vertexToOutputFieldVertices
                .Where(kvp => kvp.Value.Count > 1)
                .Select(kvp => kvp.Key)
                .ToList();

            if (sharedVertices.Count > 0)
                classesList.Add(new GeneratedClass("Shared", null, sharedVertices, rnd));

            if (Vertices.Count != classesList.Sum(l => l.Vertices.Count))
                Logger.Log(LogItemType.Error, "Error creating classes; length mismatch");

            return classesList;
        }
    }
}
