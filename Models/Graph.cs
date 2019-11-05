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

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Irony.Parsing;
using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation;
using Thesis.Models.VertexTypes;
using Thesis.ViewModels;
using XLParser;

namespace Thesis.Models
{
    public class Graph
    {
        public string WorksheetName { get; }
        public List<Vertex> Vertices { get; set; }
        public List<Vertex> ExternalVertices { get; set; }
        public Dictionary<string, Vertex> NameDictionary { get; set; } // user defined names
        public Dictionary<string, RangeVertex> RangeDictionary { get; set; }

        // for the toolbox
        public List<Vertex> AllVertices { get; set; }

        // For layouting purposes
        public int[] PopulatedRows { get; set; }
        public int[] PopulatedColumns { get; set; }
        public List<LabelGenerator.Region> Regions { get; set; }

        public Graph(string worksheetName)
        {
            WorksheetName = worksheetName;
            Vertices = new List<Vertex>();
            ExternalVertices = new List<Vertex>();
            NameDictionary = new Dictionary<string, Vertex>();
            RangeDictionary = new Dictionary<string, RangeVertex>();
        }

        public static Graph FromSpreadsheet(string worksheetName, 
            IRange cells,  
            Func<string, IRange> getRangeFunc,
            INames names)
        {
            Graph graph = new Graph(worksheetName);
            var verticesDict = new Dictionary<string, Vertex>();
            var externalVerticesDict = new Dictionary<(string sheetName, string address), Vertex>();

            // create a cell vertex for all cells in the spreadsheet
            foreach (var cell in cells.Cells)
            {
                var cellVertex = new CellVertex(cell);
                verticesDict.Add(cellVertex.StringAddress, cellVertex);
                graph.Vertices.Add(cellVertex);
            }

            Logger.Log(LogItemType.Info, $"Processing {names.Count} names in file...");

            graph.NameDictionary = GenerateNameDictionary(names, verticesDict, worksheetName);

            Logger.Log(LogItemType.Info, $"Considering {graph.Vertices.Count} vertices...");

            ProcessVertices(graph, getRangeFunc, verticesDict, externalVerticesDict, graph.NameDictionary);
            graph.ExternalVertices = externalVerticesDict.Values.ToList();

            if (graph.ExternalVertices.Count > 0)
                Logger.Log(LogItemType.Info, $"Discovered {graph.ExternalVertices.Count} external cells.");

            var rangeVertices = graph.RangeDictionary.Values.ToList();
            // add all range vertices
            graph.Vertices.AddRange(rangeVertices);

            // create regions
            var cellVertices = graph.Vertices.GetCellVertices();
            graph.Regions = LabelGenerator.CreateRegions(cellVertices);
            // generate labels
            LabelGenerator.GenerateLabelsFromRegions(cellVertices.Where(c => c.Region != null && c.Region is LabelGenerator.DataRegion));

            graph.AllVertices = graph.Vertices.ToList();

            graph.PerformTransitiveFilter(graph.GetOutputFields());

            // add parent/child for cells in ranges
            var vertexDict = graph.Vertices.OfType<CellVertex>().ToDictionary(v => v.GlobalAddress);
            foreach (var rangeVertex in rangeVertices)
            {
                foreach (var address in rangeVertex.GetAddresses())
                {
                    if (vertexDict.TryGetValue(address, out var vertex))
                    {
                        rangeVertex.Children.Add(vertex);
                        vertex.Parents.Add(rangeVertex);
                    }
                }
            }


            Logger.Log(LogItemType.Info,
                $"Filtered for reachable vertices from output fields. {graph.Vertices.Count} remaining");

            return graph;
        }

        /// <summary>
        /// Creates a dictionary which maps a name to the name's vertex.
        /// If the name is external, create a new external RangeVertex.
        /// If the name is internal, and
        /// - consists of a single cell: simply rename the existing vertex with the name's name
        /// - consists of multiple cells: create a new RangeVertex
        /// 
        /// Parent/child relationships are not established here, but after transitive filtering has been performed, to avoid
        /// adding relationships to vertices which are not referenced by output fields.
        /// </summary>
        private static Dictionary<string, Vertex> GenerateNameDictionary(INames names, Dictionary<string, Vertex> verticesDict, string worksheetName)
        {
            var nameDictionary = new Dictionary<string, Vertex>();
            foreach (NameImpl name in names)
            {
                var nameTitle = name.Name;
                var nameAddress = name.AddressLocal;

                if (nameAddress == null || name.Cells.Length == 0 || nameTitle.StartsWith("_xlnm")) continue;

                // if the named range does not apply globally (which would mean name.Worksheet = null),
                // but to a specific worksheet, check if is valid currently
                if (name.Worksheet != null && name.Worksheet.Name != worksheetName) continue;

                Vertex nameVertex;

                // external named ranges
                var nameWorksheetMatch = new Regex("'?([a-zA-Z0-9-+@#$^&()_,.! ]+)'?!").Match(name.Value);
                if (!nameWorksheetMatch.Success)
                {
                    Logger.Log(LogItemType.Warning, "Could not find worksheet for named range " + nameTitle);
                    continue;
                }

                var nameWorksheetName = nameWorksheetMatch.Groups[1].Value;
                if (nameWorksheetName != worksheetName)
                {
                    if (name.Cells.Length == 1)
                        nameVertex = new CellVertex(name.Cells[0], nameTitle);
                    else
                        nameVertex = new RangeVertex(name.Cells, nameTitle, nameAddress);
                    nameVertex.MarkAsExternal(nameWorksheetName, nameTitle);
                }
                else
                {
                    if (name.Cells.Length == 1)
                    {
                        // simply assign the named range to the cell vertex
                        if (verticesDict.TryGetValue(nameAddress, out nameVertex))
                        {
                            nameVertex.Name = nameTitle;
                        }
                        else
                        {
                            // this name is not used anywhere
                            continue;
                        }
                    }
                    else
                    {
                        // contains more than one cell
                        nameVertex = new RangeVertex(name.Cells, nameTitle, nameAddress);
                    }
                }

                if (nameDictionary.ContainsKey(nameTitle))
                    Logger.Log(LogItemType.Warning, $"Name {nameTitle} already exists!");
                else
                    nameDictionary.Add(nameTitle, nameVertex);
            }

            return nameDictionary;
        }
        
        /// <summary>
        /// Runs <see cref="ProcessVertex"/> on a graph's cell vertices
        /// </summary>
        private static void ProcessVertices(Graph graph,
            Func<string, IRange> getRangeFunc,
            Dictionary<string, Vertex> verticesDict,
            Dictionary<(string sheetName, string address), Vertex> externalVerticesDict,
            Dictionary<string, Vertex> nameDictionary)
        {
            foreach (var cellVertex in graph.Vertices.GetCellVertices())
            {
                if (cellVertex.ParseTree == null) continue;
                try
                {
                    graph.ProcessVertex(cellVertex, getRangeFunc, verticesDict, externalVerticesDict, nameDictionary);
                }
                catch (Exception ex)
                {
                    Logger.Log(LogItemType.Error,
                        $"Error processing formula in {cellVertex.StringAddress} ({cellVertex.Formula}): {ex.GetType().Name} ({ex.Message})");
                }
            }
        }

        /// <summary>
        /// Recursively traverses the parse tree of each vertex and does the following, when it encounters a
        /// - cell in the current worksheet: establishes a parent/child relationship
        /// - cell in another worksheet: creates a new cell for the external cell, adds it to externalVerticesDict, and establishes a parent/child relationship
        /// - range: creates a new RangeVertex for the range, adds it to RangeDictionary, and establishes a parent/child relationship.
        /// - name: uses nameDictionary to find the named range, and establishes a parent/child relationship
        /// </summary>
        private void ProcessVertex(CellVertex cellVertex, 
                Func<string, IRange> getRangeFunc,
                Dictionary<string, Vertex> verticesDict,
                Dictionary<(string sheetName, string address), Vertex> externalVerticesDict,
                Dictionary<string, Vertex> nameDictionary)
        {
            var stack = new Stack<ParseTreeNode>();
            var visited = new HashSet<ParseTreeNode>();
            stack.Push(cellVertex.ParseTree);

            // maps a node to an external sheet
            var nodeToExternalSheetDictionary = new Dictionary<ParseTreeNode, string>();
            GenerateNodeToExternalSheetDictionary(cellVertex.ParseTree, cellVertex.ParseTree, nodeToExternalSheetDictionary);

            visited.Clear();
            stack.Push(cellVertex.ParseTree);

            while (stack.Count > 0)
            {
                var node = stack.Pop();
                if (visited.Contains(node)) continue;

                visited.Add(node);
                bool continueWithChildren = true;

                switch (node.Term.Name)
                {
                    case "ReferenceFunctionCall":
//                    case "VRange": not implemented yet
//                    case "HRange":
                    {
                        if (node.Term.Name == "ReferenceFunctionCall")
                        {
                            if (node.GetFunction() != ":")
                                break;
                        }
                        else
                        {
                            node = node.Parent(cellVertex.ParseTree);
                        }

                        string range = node.NodeToString(cellVertex.Formula);

                        if (!RangeDictionary.TryGetValue(range, out var rangeVertex))
                        {
                            var iRange = getRangeFunc(range);
                            var cells = iRange != null ? iRange.Cells : new IRange[0];
                            rangeVertex = new RangeVertex(cells, range, range);
                            RangeDictionary.Add(range, rangeVertex);
                        }
                        if (nodeToExternalSheetDictionary.TryGetValue(node, out var sheetName))
                            rangeVertex.MarkAsExternal(sheetName);

                        rangeVertex.Parents.Add(cellVertex);
                        cellVertex.Children.Add(rangeVertex);

                        continueWithChildren = false;

                        break;
                    }
                    case "Cell":
                    {
                        string referencedCellAddress = node.FindTokenAndGetText();
                        if (nodeToExternalSheetDictionary.TryGetValue(node, out var externalSheetName))
                        {
                            if (externalVerticesDict.TryGetValue((externalSheetName, referencedCellAddress), out var externalCellVertex))
                            {
                                cellVertex.Children.Add(externalCellVertex);
                                externalCellVertex.Parents.Add(cellVertex);
                            }
                            else
                            {
                                var wholeAddress = node.Parent(cellVertex.ParseTree).NodeToString(cellVertex.Formula);
                                var externalCellIRange = getRangeFunc(wholeAddress);
                                if (externalCellIRange == null)
                                {
                                    Logger.Log(LogItemType.Warning, $"Could not get cell {externalSheetName}.{referencedCellAddress}");
                                }
                                else
                                {
                                    externalCellVertex = new CellVertex(externalCellIRange);
                                    externalCellVertex.MarkAsExternal(externalSheetName, referencedCellAddress);

                                    cellVertex.Children.Add(externalCellVertex);
                                    externalCellVertex.Parents.Add(cellVertex);
                                    externalVerticesDict.Add((externalSheetName, referencedCellAddress), externalCellVertex);
                                }
                            }
                        }
                        else
                        {
                            cellVertex.Children.Add(verticesDict[referencedCellAddress]);
                            verticesDict[referencedCellAddress].Parents.Add(cellVertex);
                        }

                        break;
                    }
                    case "NamedRange":
                    {
                        var name = node.FindTokenAndGetText();
                        if (nameDictionary.TryGetValue(name, out var nameVertex))
                        {
                            cellVertex.Children.Add(nameVertex);
                            nameVertex.Parents.Add(cellVertex);
                        }
                        else
                        {
                            Logger.Log(LogItemType.Warning, "Could not find named range " + name);
                        }
                        break;
                    }
                }


                if (node.Term.Name == "UDFunctionCall")
                {
                    Logger.Log(LogItemType.Warning,
                        $"Skipping user defined function {node.FindTokenAndGetText()} in {cellVertex.Address}. Declaring as constant!");
                    foreach (var children in cellVertex.Children)
                    {
                        children.Parents.Remove(cellVertex);
                    }
                    cellVertex.Children.Clear();
                    cellVertex.ParseTree = null;
                    return;
                }
                else if (continueWithChildren)
                {
                    for (var i = node.ChildNodes.Count - 1; i >= 0; i--)
                    {
                        stack.Push(node.ChildNodes[i]);
                    }
                }
            }

        }

        /// <summary>
        /// Recursively traverses down a ParseTreeNode, and populates the dictionary with entries of nodes which are part of an
        /// external sheet reference
        /// </summary>
        private void GenerateNodeToExternalSheetDictionary(ParseTreeNode node, ParseTreeNode rootNode,
            Dictionary<ParseTreeNode, string> nodeToExternalSheetDictionary)
        {
            if (node.Term.Name == "Prefix")
            {
                // SheetNameToken vs ParsedSheetNameToken
                string sheetName = node.ChildNodes.Count == 2
                    ? node.ChildNodes[1].FindTokenAndGetText()
                    : node.FindTokenAndGetText();
                // remove ! at end of sheet name
                sheetName = sheetName.FormatSheetName();

                // traverse up the parse tree until a Reference node is find
                ParseTreeNode referenceNode = node.Parent(rootNode);
                var traverse = node;
                while (traverse != rootNode)
                {
                    traverse = traverse.Parent(rootNode);
                    if (traverse.Term.Name == "Reference")
                        referenceNode = traverse;
                }

                // make all children of the reference node as external
                MarkChildNodesAsExternal(nodeToExternalSheetDictionary, referenceNode, sheetName);
            }
            else
            {
                foreach (var child in node.ChildNodes)
                    GenerateNodeToExternalSheetDictionary(child, rootNode, nodeToExternalSheetDictionary);
            }

        }

        /// <summary>
        /// Recursively traverses down a ParseTreeNode, and adds all child nodes to the dictionary
        /// </summary>
        private void MarkChildNodesAsExternal(Dictionary<ParseTreeNode, string> nodeToExternalSheetDictionary, ParseTreeNode node, string sheetName)
        {
            foreach (var child in node.ChildNodes)
            {
                if (nodeToExternalSheetDictionary.ContainsKey(child)) return;
                nodeToExternalSheetDictionary.Add(child, sheetName);
                MarkChildNodesAsExternal(nodeToExternalSheetDictionary, child, sheetName);
            }
        }

        public List<CellVertex> GetOutputFields()
        {
            return Vertices
                .OfType<CellVertex>()
                .Where(v => v.NodeType == NodeType.OutputField).ToList();
        }

        // Remove all vertices that are not transitively reachable from any vertex in the given list
        public void PerformTransitiveFilter(List<CellVertex> cellVertices)
        {
            var logItem = Logger.Log(LogItemType.Info, "Perform transitive filtering...", true);
            Vertices = cellVertices.SelectMany(v => v.GetReachableVertices()).Distinct().ToList();

            // filter external vertices for those which have at least one parent still in the vertices list
            ExternalVertices = cellVertices
                .SelectMany(v => v.GetReachableVertices(false)).Where(v => v.IsExternal)
                .Distinct()
                .ToList();
            logItem.AppendElapsedTime();

            // create PopulatedRows, PopulatedColumns array
            var cells = Vertices.OfType<CellVertex>().ToList();
            var populatedRows = cells.Select(v => v.Address.row).Distinct().ToList();
            foreach (var rangeVertex in RangeDictionary.Values)
                populatedRows.AddRange(rangeVertex.GetPopulatedRows());
            PopulatedRows = populatedRows.Distinct().OrderBy(x => x).ToArray();

            var populatedColumns = cells.Select(v => v.Address.col).Distinct().ToList();
            foreach (var rangeVertex in RangeDictionary.Values)
                populatedColumns.AddRange(rangeVertex.GetPopulatedColumns());
            PopulatedColumns = populatedColumns.Distinct().OrderBy(x => x).ToArray();
        }

        public Vertex GetVertexByAddress(string sheetName, int row, int col)
        {
            if (sheetName == WorksheetName)
            {
                return AllVertices
                    .OfType<CellVertex>()
                    .FirstOrDefault(v => v.Address.row == row && v.Address.col == col);
            }

            return ExternalVertices
                .FirstOrDefault(v => v.ExternalWorksheetName == sheetName &&
                                     (v is CellVertex cellVertex && cellVertex.Address.row == row && cellVertex.Address.col == col
                                     || v is RangeVertex rangeVertex && rangeVertex.GetAddressTuples().Contains((row, col))));
        }
    }
}