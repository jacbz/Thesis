﻿using System;
using System.Collections.Generic;
using System.Linq;
using Irony.Parsing;
using Syncfusion.XlsIO;
using Thesis.ViewModels;
using XLParser;

namespace Thesis.Models
{
    public class Graph
    {
        public List<Vertex> Vertices { get; set; }
        public List<Vertex> ExternalVertices { get; set; }

        // Preserve copy for filtering purposes
        public List<Vertex> AllVertices { get; set; }
        public List<Vertex> AllExternalVertices { get; set; }

        // For layouting purposes
        public List<int> PopulatedRows { get; set; }
        public List<int> PopulatedColumns { get; set; }

        public Graph(IRange cells, Func<string, string, IRange> getExternalCellFunc)
        {
            var verticesDict = new Dictionary<string, Vertex>();
            var externalVerticesDict = new Dictionary<(string sheetName, string address), Vertex>();

            Vertices = new List<Vertex>();
            ExternalVertices = new List<Vertex>();

            foreach (var cell in cells.Cells)
            {
                var vertex = new Vertex(cell);

                verticesDict.Add(vertex.StringAddress, vertex);
                Vertices.Add(vertex);
            }

            Logger.Log(LogItemType.Info, $"Considering {Vertices.Count} vertices...");

            foreach (var vertex in Vertices)
            {
                if (vertex.ParseTree == null) continue;
                try
                {
                    var (referencedCells, externalReferencedCells) = GetListOfReferencedCells(vertex.StringAddress, vertex.ParseTree);
                    foreach (var cell in referencedCells)
                    {
                        vertex.Children.Add(verticesDict[cell]);
                        verticesDict[cell].Parents.Add(vertex);
                    }
                    // process external cells
                    foreach(var (sheetName, address) in externalReferencedCells)
                    {
                        if (externalVerticesDict.TryGetValue((sheetName, address), out Vertex externalVertex))
                        {
                            vertex.Children.Add(externalVertex);
                            externalVertex.Parents.Add(vertex);
                        }
                        else
                        {
                            var externalCellIRange = getExternalCellFunc(sheetName, address);
                            if (externalCellIRange == null)
                            {
                                Logger.Log(LogItemType.Warning, $"Could not get cell {sheetName}.{address}");
                            }
                            else
                            {
                                externalVertex = new Vertex(externalCellIRange);
                                externalVertex.ExternalWorksheetName = sheetName;
                                externalVertex.VariableName = sheetName.ToPascalCase() + "_" + address;

                                vertex.Children.Add(externalVertex);
                                externalVertex.Parents.Add(vertex);
                                ExternalVertices.Add(externalVertex);
                                externalVerticesDict.Add((sheetName, address), externalVertex);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.Log(LogItemType.Error,
                        $"Error processing formula in {vertex.StringAddress} ({vertex.Formula}): {ex.GetType().Name} ({ex.Message})");
                }
            }

            if (ExternalVertices.Count > 0)
                Logger.Log(LogItemType.Info, $"Discovered {ExternalVertices.Count} external cells.");

            GenerateLabels();

            AllVertices = Vertices.ToList();
            AllExternalVertices = ExternalVertices.ToList();

            PerformTransitiveFilter(GetOutputFields());

            Logger.Log(LogItemType.Info,
                $"Filtered for reachable vertices from output fields. {Vertices.Count} remaining");

        }

        private void GenerateLabels()
        {
            var logItem = Logger.Log(LogItemType.Info, "Generating labels...", true);
            // Create labels
            Dictionary<(int row, int col), Label> labelDictionary = new Dictionary<(int row, int col), Label>();
            foreach (var vertex in Vertices)
            {
                Label label = new Label(vertex);
                if (vertex.CellType == CellType.Unknown && vertex.NodeType == NodeType.None)
                {
                    label.Type = LabelType.None;
                }
                else if (vertex.Children.Count == 0 && vertex.Parents.Count == 0 && vertex.CellType == CellType.Text)
                {
                    if (labelDictionary.TryGetValue((vertex.Address.row - 1, vertex.Address.col), out Label labelAbove)
                        && (labelAbove.Type == LabelType.Attribute || labelAbove.Type == LabelType.Header))
                    {
                        label.Type = LabelType.Attribute;
                        label.Text = vertex.DisplayValue;
                        labelAbove.Type = LabelType.Attribute;
                    }
                    else
                    {
                        label.Type = LabelType.Header;
                        label.Text = vertex.DisplayValue;
                    }
                }
                else
                {
                    label.Type = LabelType.Data;
                }

                vertex.Label = label;
                labelDictionary.Add((vertex.Address.row, vertex.Address.col), label);
            }

            // assign attributes and headers for each data type
            foreach (var vertex in Vertices)
            {
                if (vertex.Label.Type != LabelType.Data) continue;

                (int row, int col) currentPos = vertex.Address;

                // add attributes
                bool foundAttribute = false;
                int distanceToAttribute = 0;
                // a list that stores how far all attributes are to the vertex. e.g. attribute in 2,3, vertex in 8: [5,6]
                List<int> distancesToAttribute = new List<int>();
                while (currentPos.col-- > 1)
                {
                    var currentLabel = labelDictionary[currentPos];
                    if (foundAttribute && currentLabel.Type != LabelType.Attribute) break;

                    distanceToAttribute++;
                    if (currentLabel.Type == LabelType.Attribute)
                    {
                        foundAttribute = true;
                        vertex.Label.Attributes.Add(currentLabel);
                        distancesToAttribute.Add(distanceToAttribute);
                    }
                }

                // add headers
                currentPos = vertex.Address;
                if (!foundAttribute)
                {
                    // no attributes, use first header on the top
                    while (currentPos.row-- > 1)
                    {
                        var currentLabel = labelDictionary[currentPos];
                        if (currentLabel.Type == LabelType.Header)
                        {
                            vertex.Label.Headers.Add(currentLabel);
                            break;
                        }
                    }
                }
                else
                {
                    // keep adding headers, until there is no attribute to the left or left bottom with the exact distance
                    bool foundHeader = false;
                    while (currentPos.row-- > 1)
                    {
                        var currentLabel = labelDictionary[currentPos];

                        bool anyAttributeDistanceMatch = false;
                        foreach (int dist in distancesToAttribute)
                        {
                            if (labelDictionary[(currentPos.row, currentPos.col - dist)].Type == LabelType.Attribute ||
                                labelDictionary[(currentPos.row + 1, currentPos.col - dist)].Type == LabelType.Attribute)
                                anyAttributeDistanceMatch = true;
                        }

                        if (!anyAttributeDistanceMatch)
                            break;
                        if (foundHeader && currentLabel.Type != LabelType.Header)
                            break;

                        if (currentLabel.Type == LabelType.Header)
                        {
                            foundHeader = true;
                            vertex.Label.Headers.Add(currentLabel);
                        }
                    }
                }

                vertex.Label.GenerateVariableName();
            }
            logItem.AppendElapsedTime();
        }

        // recursively gets list of referenced cells from parse tree using DFS
        private (List<string>, List<(string, string)>) GetListOfReferencedCells(string address, ParseTreeNode parseTree)
        {
            var referencedCells = new List<string>();
            var externalReferencedCells = new List<(string sheetName, string address)>();
            var stack = new Stack<ParseTreeNode>();
            var visited = new HashSet<ParseTreeNode>();
            stack.Push(parseTree);

            // maps a Reference node to an external sheet
            var referenceNodeToExternalSheet = new HashSet<(ParseTreeNode node, string sheetName)>();

            while (stack.Count > 0)
            {
                var node = stack.Pop();
                if (visited.Contains(node)) continue;

                visited.Add(node);

                switch (node.Term.Name)
                {
                    case "ReferenceFunctionCall":
                        // ranges
                        if (node.ChildNodes.Count == 3 && node.GetFunction() == ":")
                        {
                            var leftChild = node.ChildNodes[0].ChildNodes[0];
                            var rightChild = node.ChildNodes[2].ChildNodes[0];
                            if (leftChild.Term.Name == "Cell" && rightChild.Term.Name == "Cell")
                            {
                                var leftAddress = leftChild.FindTokenAndGetText();
                                var rightAddress = rightChild.FindTokenAndGetText();

                                var addressesInRange = Utility.AddressesInRange(leftAddress, rightAddress).ToArray();
                                Logger.Log(LogItemType.Info,
                                    $"Found range in {address}: Adding {string.Join(", ", addressesInRange)}");
                                referencedCells.AddRange(addressesInRange);
                            }
                        }
                        break;
                    case "Cell":
                        var externalMatch = referenceNodeToExternalSheet.FirstOrDefault(e => e.node == node);
                        if (!externalMatch.Equals(default))
                        {
                            externalReferencedCells.Add((externalMatch.sheetName, node.FindTokenAndGetText()));
                        }
                        else
                        {
                            referencedCells.Add(node.FindTokenAndGetText());
                        }
                        break;
                    case "Prefix":
                        // SheetNameToken vs ParsedSheetNameToken
                        string sheetName = node.ChildNodes.Count == 2 
                            ? node.ChildNodes[1].FindTokenAndGetText() 
                            : node.FindTokenAndGetText();
                        // remove ! at end of sheet name
                        sheetName = sheetName.FormatSheetName();

                        // reference must be as far above as possible, to parse e.g. 'Sheet'!A1:A8
                        ParseTreeNode reference = node.Parent(parseTree);
                        var traverse = node;
                        while (traverse != parseTree)
                        {
                            traverse = traverse.Parent(parseTree);
                            if (traverse.Term.Name == "Reference")
                                reference = traverse;
                        }
                        MarkChildNodesAsExternal(referenceNodeToExternalSheet, reference, sheetName);
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

        private void MarkChildNodesAsExternal(HashSet<(ParseTreeNode, string)> externalList, ParseTreeNode node, string refName)
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

        // Remove all vertices that are not transitively reachable from any vertex in the given list
        public void PerformTransitiveFilter(List<Vertex> vertices)
        {
            var logItem = Logger.Log(LogItemType.Info, "Perform transitive filter...", true);
            Vertices = vertices.SelectMany(v => v.GetReachableVertices()).Distinct().ToList();
            logItem.AppendElapsedTime();

            PopulatedRows = Vertices.Select(v => v.Address.row).Distinct().ToList();
            PopulatedRows.Sort();
            PopulatedColumns = Vertices.Select(v => v.Address.col).Distinct().ToList();
            PopulatedColumns.Sort();

            // filter external vertices for those which have at least one parent still in the vertices list
            if (AllExternalVertices.Count == 0) return;
            var logItem2 = Logger.Log(LogItemType.Info, "Perform transitive filter for external cells...", true);
            ExternalVertices = AllExternalVertices.Where(v => v.Parents.Any(p => Vertices.Contains(p))).ToList();
            logItem2.AppendElapsedTime();
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
                var newClass = new GeneratedClass($"Class{vertex.StringAddress}", vertex, reachableVertices.ToList(), rnd);
                classesList.Add(newClass);
            }

            var logItem = Logger.Log(LogItemType.Info, "Applying topological sort...", true);
            foreach (var generatedClass in classesList)
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

            if (sharedVertices.Count > 0)
            {
                var globalClass = new GeneratedClass("Global", null, sharedVertices);
                globalClass.TopologicalSort();
                classesList.Add(globalClass);
            }

            if (ExternalVertices.Count > 0)
            {
                var externalClass = new GeneratedClass("External", null, ExternalVertices);
                classesList.Add(externalClass);
            }

            if (Vertices.Count + ExternalVertices.Count != classesList.Sum(l => l.Vertices.Count))
                Logger.Log(LogItemType.Error, "Error creating classes; number of vertices does not match number of vertices in classes");

            return classesList;
        }
    }
}