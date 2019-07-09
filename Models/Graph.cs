using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Irony.Parsing;
using Syncfusion.XlsIO;
using Thesis.Models;
using Thesis.Models.CodeGenerators;
using Thesis.ViewModels;
using XLParser;

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
            var verticesDict = new Dictionary<string, Vertex>();
            Vertices = new List<Vertex>();

            foreach (var cell in cells.Cells)
            {
                var vertex = new Vertex(cell);

                verticesDict.Add(vertex.Address, vertex);
                Vertices.Add(vertex);
            }

            Logger.Log(LogItemType.Info, $"Considering {Vertices.Count} vertices...");

            foreach (var vertex in Vertices)
            {
                if (vertex.Type != CellType.Formula) continue;

                try
                {
                    var (referencedCells, externalReferencedCells) =
                        GetListOfReferencedCells(vertex.Address, vertex.ParseTree);
                    foreach (var cell in referencedCells)
                    {
                        vertex.Children.Add(verticesDict[cell]);
                        verticesDict[cell].Parents.Add(vertex);
                    }
                    // TODO: so something with external cells

                }
                catch (Exception ex)
                {
                    Logger.Log(LogItemType.Error,
                        $"Error processing formula in {vertex.Address} ({vertex.Formula}): {ex.Message}");
                }
            }

            TransitiveFilter(GetOutputFields());
            Logger.Log(LogItemType.Info,
                $"Filtered for reachable vertices from output fields. {Vertices.Count} remaining");

            AllVertices = Vertices.ToList();
        }

        // recursively gets list of referenced cells from parse tree using DFS
        private (List<string>, List<(string, string)>) GetListOfReferencedCells(string address, ParseTreeNode parseTree)
        {
            var referencedCells = new List<string>();
            var externalReferencedCells = new List<(string, string)>();
            var stack = new Stack<ParseTreeNode>();
            var visited = new HashSet<ParseTreeNode>();
            stack.Push(parseTree);

            // maps a Reference node to an external sheet
            var external = new List<(ParseTreeNode, string)>();

            while (stack.Count > 0)
            {
                var node = stack.Pop();
                if (visited.Contains(node)) continue;

                visited.Add(node);

                switch (node.Term.Name)
                {
                    case "Cell":
                        var externalMatch = external.FirstOrDefault(e => e.Item1 == node);
                        if (!externalMatch.Equals(default))
                        {
                            Logger.Log(LogItemType.Warning, $"Skipping external reference in {address}: {externalMatch.Item2}");
                            externalReferencedCells.Add((node.FindTokenAndGetText(), externalMatch.Item2));
                        }
                        else
                        {
                            referencedCells.Add(node.FindTokenAndGetText());
                        }
                        break;
                    case "Prefix":
                        // SheetNameToken vs ParsedSheetNameToken
                        string refName = node.ChildNodes.Count == 2 ? node.ChildNodes[1].FindTokenAndGetText() : node.FindTokenAndGetText();

                        // reference must on as far above as possible, to parse e.g. 'Sheet'!A1:A8
                        ParseTreeNode reference = node.Parent(parseTree);
                        var traverse = node;
                        while (traverse != parseTree)
                        {
                            traverse = traverse.Parent(parseTree);
                            if (traverse.Term.Name == "Reference")
                                reference = traverse;
                        }
                        MarkChildNodesAsExternal(external, reference, refName);
                        break;
                }

                if (node.Term.Name == "UDFunctionCall")
                {
                    Logger.Log(LogItemType.Warning,
                        $"Skipping user defined function {node.FindTokenAndGetText()} in {address}");
                }
                else
                {
                    for (var i = node.ChildNodes.Count - 1; i >= 0; i--)
                    {
                        stack.Push(node.ChildNodes[i]);
                    }
                }
            }

            return (referencedCells, externalReferencedCells);
        }

        private void MarkChildNodesAsExternal(List<(ParseTreeNode, string)> externalList, ParseTreeNode node, string refName)
        {
            foreach (var child in node.ChildNodes)
            {
                externalList.Add((child, refName));
                MarkChildNodesAsExternal(externalList, child, refName);
            }
        }

        public List<Vertex> GetOutputFields()
        {
            return Vertices.Where(v => v.NodeType == NodeType.OutputField).ToList();
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

        public void GenerateLabels(IWorksheet worksheet)
        {
            foreach (var vertex in Vertices)
            {
                // go to the left until
                var row = vertex.CellIndex[0];
                var column = vertex.CellIndex[1] - 1;
                while (column > 0 && !vertex.HasLabel)
                {
                    var cell = worksheet.Range[row, column];
                    if (Vertices.Any(v => v.Address == cell.Address)) continue;
                    if (Vertex.GetCellType(cell) == CellType.Text && !string.IsNullOrWhiteSpace(cell.DisplayText))
                        vertex.Label = CodeGenerator.ToCamelCase(cell.DisplayText);
                    column--;
                }
            }
        }

        public List<GeneratedClass> GenerateClasses()
        {
            var classesList = new List<GeneratedClass>();
            var vertexToOutputFieldVertices = Vertices.ToDictionary(v => v, v => new HashSet<Vertex>());
            var rnd = new Random();

            foreach (var vertex in GetOutputFields())
            {
                var reachableVertices = vertex.GetReachableVertices();
                foreach (var v in reachableVertices) vertexToOutputFieldVertices[v].Add(vertex);
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