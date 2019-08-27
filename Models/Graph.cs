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
        public List<Vertex> Vertices { get; set; }
        public List<Vertex> ExternalVertices { get; set; }
        public Dictionary<string, Vertex> NameDictionary { get; set; } // user defined names
        public Dictionary<string, RangeVertex> RangeDictionary { get; set; }

        // toolbox
        public List<Vertex> AllVertices { get; set; }

        // For layouting purposes
        public List<int> PopulatedRows { get; set; }
        public List<int> PopulatedColumns { get; set; }

        public Graph()
        {
            Vertices = new List<Vertex>();
            ExternalVertices = new List<Vertex>();
            NameDictionary = new Dictionary<string, Vertex>();
            RangeDictionary = new Dictionary<string, RangeVertex>();
        }

        public static Graph FromSpreadsheet(string worksheetName, 
            IRange cells, 
            Func<string, string, IRange> getExternalCellFunc, 
            Func<string, IRange> getRangeFunc,
            INames names)
        {
            Graph graph = new Graph();
            var verticesDict = new Dictionary<string, Vertex>();
            var externalVerticesDict = new Dictionary<(string sheetName, string address), Vertex>();

            foreach (var cell in cells.Cells)
            {
                var cellVertex = new CellVertex(cell);

                verticesDict.Add(cellVertex.StringAddress, cellVertex);
                graph.Vertices.Add(cellVertex);
            }

            // assigns a name name to a Vertex
            foreach (NameImpl name in names)
            {
                var nameTitle = name.Name;
                var nameAddress = name.AddressLocal;

                if (nameAddress == null || name.Cells.Length == 0 || nameTitle == "_xlnm._FilterDatabase") continue;

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
                    nameVertex = new RangeVertex(name.Cells, nameTitle);
                    nameVertex.MarkAsExternal(nameWorksheetName, nameTitle);
                }
                else
                {
                    if (name.Cells.Length == 1)
                    {
                        // simply assign the named range to the cell vertex
                        if (verticesDict.TryGetValue(nameAddress, out nameVertex))
                        {
                            nameVertex.VariableName = nameTitle;
                        }
                        else
                        {
                            Logger.Log(LogItemType.Warning, "Could not find vertex for named range " + nameTitle);
                            continue;
                        }
                    }
                    else
                    {
                        // contains more than one cell
                        nameVertex = new RangeVertex(name.Cells, nameTitle);
                    }
                }

                if (graph.NameDictionary.ContainsKey(nameTitle))
                    Logger.Log(LogItemType.Warning, $"Name {nameTitle} already exists!");
                else
                    graph.NameDictionary.Add(nameTitle, nameVertex);
            }

            Logger.Log(LogItemType.Info, $"Considering {graph.Vertices.Count} vertices...");

            foreach (var cellVertex in graph.Vertices.OfType<CellVertex>())
            {
                if (cellVertex.ParseTree == null) continue;
                try
                {
                    var (referencedCells, externalReferencedCells, referencedNames) =
                        graph.GetListOfReferencedCells(
                            cellVertex,
                            getRangeFunc);
                    foreach (var cellAddress in referencedCells)
                    {
                        cellVertex.Children.Add(verticesDict[cellAddress]);
                        verticesDict[cellAddress].Parents.Add(cellVertex);
                    }
                    // process external cells
                    foreach(var (sheetName, address) in externalReferencedCells)
                    {
                        if (externalVerticesDict.TryGetValue((sheetName, address), out var externalCellVertex))
                        {
                            cellVertex.Children.Add(externalCellVertex);
                            externalCellVertex.Parents.Add(cellVertex);
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
                                externalCellVertex = new CellVertex(externalCellIRange);
                                externalCellVertex.MarkAsExternal(sheetName, address);

                                cellVertex.Children.Add(externalCellVertex);
                                externalCellVertex.Parents.Add(cellVertex);
                                graph.ExternalVertices.Add(externalCellVertex);
                                externalVerticesDict.Add((sheetName, address), externalCellVertex);
                            }
                        }
                    }

                    // process named ranges
                    foreach (var nameName in referencedNames)
                    {
                        if (graph.NameDictionary.TryGetValue(nameName, out var nameVertex))
                        {
                            cellVertex.Children.Add(nameVertex);
                            nameVertex.Parents.Add(cellVertex);
                        }
                        else
                        {
                            Logger.Log(LogItemType.Warning, "Could not find named range " + nameName);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.Log(LogItemType.Error,
                        $"Error processing formula in {cellVertex.StringAddress} ({cellVertex.Formula}): {ex.GetType().Name} ({ex.Message})");
                }
            }

            if (graph.ExternalVertices.Count > 0)
                Logger.Log(LogItemType.Info, $"Discovered {graph.ExternalVertices.Count} external cells.");

            // add all range vertices
            graph.Vertices.AddRange(graph.RangeDictionary.Values.ToList());

            graph.GenerateLabels();

            graph.AllVertices = graph.Vertices.ToList();

            graph.PerformTransitiveFilter(graph.GetOutputFields());

            Logger.Log(LogItemType.Info,
                $"Filtered for reachable vertices from output fields. {graph.Vertices.Count} remaining");

            return graph;
        }

        private void GenerateLabels()
        {
            var logItem = Logger.Log(LogItemType.Info, "Generating labels...", true);
            // Create labels
            Dictionary<(int row, int col), Label> labelDictionary = new Dictionary<(int row, int col), Label>();

            var cells = Vertices.GetCellVertices();
            foreach (var cell in cells)
            {
                Label label = new Label(cell);
                if (cell.CellType == CellType.Unknown && cell.NodeType == NodeType.None)
                {
                    label.Type = LabelType.None;
                }
                else if (cell.Children.Count == 0 && cell.Parents.Count == 0 && cell.CellType == CellType.Text)
                {
                    if (labelDictionary.TryGetValue((cell.Address.row - 1, cell.Address.col), out Label labelAbove)
                        && (labelAbove.Type == LabelType.Attribute || labelAbove.Type == LabelType.Header))
                    {
                        label.Type = LabelType.Attribute;
                        label.Text = cell.DisplayValue;
                        labelAbove.Type = LabelType.Attribute;
                    }
                    else
                    {
                        label.Type = LabelType.Header;
                        label.Text = cell.DisplayValue;
                    }
                }
                else
                {
                    label.Type = LabelType.Data;
                }

                cell.Label = label;
                labelDictionary.Add((cell.Address.row, cell.Address.col), label);
            }

            // assign attributes and headers for each data type
            foreach (var cell in cells)
            {
                if (!cell.IsSpreadsheetCell || cell.Label.Type != LabelType.Data) continue;

                (int row, int col) currentPos = cell.Address;

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
                        cell.Label.Attributes.Add(currentLabel);
                        distancesToAttribute.Add(distanceToAttribute);
                    }
                }

                // add headers
                currentPos = cell.Address;
                if (!foundAttribute)
                {
                    // no attributes, use first header on the top
                    while (currentPos.row-- > 1)
                    {
                        var currentLabel = labelDictionary[currentPos];
                        if (currentLabel.Type == LabelType.Header)
                        {
                            cell.Label.Headers.Add(currentLabel);
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
                            cell.Label.Headers.Add(currentLabel);
                        }
                    }
                }

                // do not override name if name was already assigned, e.g. per named range
                if (string.IsNullOrEmpty(cell.VariableName))
                    cell.Label.GenerateVariableName();
            }
            logItem.AppendElapsedTime();
        }

        // recursively gets list of referenced cells from parse tree using DFS
        private (List<string> referencedCells, List<(string, string)> externalReferencedCells, List<string> referencednames)
            GetListOfReferencedCells(CellVertex cellVertex, Func<string, IRange> getRangeFunc)
        {
            var referencedCells = new List<string>();
            var externalReferencedCells = new List<(string sheetName, string address)>();
            var referencedNames = new List<string>();

            var stack = new Stack<ParseTreeNode>();
            var visited = new HashSet<ParseTreeNode>();
            stack.Push(cellVertex.ParseTree);

            // maps a Reference node to an external sheet
            var referenceNodeToExternalSheet = new HashSet<(ParseTreeNode node, string sheetName)>();

            while (stack.Count > 0)
            {
                var node = stack.Pop();
                if (visited.Contains(node)) continue;

                visited.Add(node);

                bool processChildren = true;
                switch (node.Term.Name)
                {
                    case "ReferenceFunctionCall":
                        string range = node.NodeToString(cellVertex.Formula);

                        if (!RangeDictionary.TryGetValue(range, out var rangeVertex))
                        {
                            rangeVertex = new RangeVertex(getRangeFunc(range).Cells, range);
                            RangeDictionary.Add(range, rangeVertex);
                        }

                        rangeVertex.Parents.Add(cellVertex);
                        cellVertex.Children.Add(rangeVertex);

                        break;
                    case "Cell":
                        var externalMatch = referenceNodeToExternalSheet.FirstOrDefault(e => e.node == node);
                        if (!externalMatch.Equals(default))
                            externalReferencedCells.Add((externalMatch.sheetName, node.FindTokenAndGetText()));
                        else
                            referencedCells.Add(node.FindTokenAndGetText());
                        break;
                    case "NamedRange":
                        var name = node.FindTokenAndGetText();
                        referencedNames.Add(name);
                        processChildren = false;
                        break;
                    case "Prefix":
                        // SheetNameToken vs ParsedSheetNameToken
                        string sheetName = node.ChildNodes.Count == 2 
                            ? node.ChildNodes[1].FindTokenAndGetText() 
                            : node.FindTokenAndGetText();
                        // remove ! at end of sheet name
                        sheetName = sheetName.FormatSheetName();

                        // reference must be as far above as possible, to parse e.g. 'Sheet'!A1:A8
                        ParseTreeNode reference = node.Parent(cellVertex.ParseTree);
                        var traverse = node;
                        while (traverse != cellVertex.ParseTree)
                        {
                            traverse = traverse.Parent(cellVertex.ParseTree);
                            if (traverse.Term.Name == "Reference")
                                reference = traverse;
                        }
                        MarkChildNodesAsExternal(referenceNodeToExternalSheet, reference, sheetName);
                        break;
                }

                if (node.Term.Name == "UDFunctionCall")
                {
                    Logger.Log(LogItemType.Warning,
                        $"Skipping user defined function {node.FindTokenAndGetText()} in {cellVertex.Address}");
                }
                else if (processChildren)
                {
                    for (var i = node.ChildNodes.Count - 1; i >= 0; i--)
                    {
                        stack.Push(node.ChildNodes[i]);
                    }
                }
            }

            return (referencedCells, externalReferencedCells, referencedNames);
        }

        private void MarkChildNodesAsExternal(HashSet<(ParseTreeNode, string)> externalList, ParseTreeNode node, string refName)
        {
            foreach (var child in node.ChildNodes)
            {
                externalList.Add((child, refName));
                MarkChildNodesAsExternal(externalList, child, refName);
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
            var logItem = Logger.Log(LogItemType.Info, "Perform transitive filter...", true);
            Vertices = cellVertices.SelectMany(v => v.GetReachableVertices()).Distinct().ToList();
            logItem.AppendElapsedTime();

            var cells = Vertices.OfType<CellVertex>().ToList();
            PopulatedRows = cells.Select(v => v.Address.row).Distinct().ToList();
            PopulatedRows.Sort();
            PopulatedColumns = cells.Select(v => v.Address.col).Distinct().ToList();
            PopulatedColumns.Sort();

            // filter external vertices for those which have at least one parent still in the vertices list
            Logger.Log(LogItemType.Info, "Perform transitive filter for external cells...");
            ExternalVertices = cellVertices
                .SelectMany(v => v.GetReachableVertices(false)).Where(v => v.IsExternal)
                .Distinct()
                .ToList();
        }

        public CellVertex GetVertexByAddress(int row, int col)
        {
            return AllVertices
                .OfType<CellVertex>()
                .FirstOrDefault(v =>
                    v.Address.row == row &&
                    v.Address.col == col);
        }
    }
}