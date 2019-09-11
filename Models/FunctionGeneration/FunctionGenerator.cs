using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Irony.Parsing;
using XLParser;
using Thesis.Models.VertexTypes;
using Thesis.ViewModels;

namespace Thesis.Models.FunctionGeneration
{
    public class FunctionGenerator
    {
        private Dictionary<(string worksheet, string address), CellVertex> _addressToVertexDictionary;
        private Dictionary<string, RangeVertex> _rangeDictionary;
        private Dictionary<string, Vertex> _nameDictionary;

        public static FunctionGenerator Instantiate(Dictionary<(string worksheet, string address), CellVertex> addressToVertexDictionary, Dictionary<string, RangeVertex> rangeDictionary, Dictionary<string, Vertex> nameDictionary)
        {
            var formulaToFunctionConverter = new FunctionGenerator
            {
                _addressToVertexDictionary = addressToVertexDictionary,
                _rangeDictionary = rangeDictionary,
                _nameDictionary = nameDictionary
            };
            return formulaToFunctionConverter;
        }

        public Expression ConstantAndConstantFormulaVertexToExpression(CellVertex vertex)
        {
            if (vertex.Classification == Classification.Constant)
                return new Constant(vertex.Value, vertex.CellType);
            return ParseTreeNodeToExpression(vertex.ParseTree, vertex.Formula);
        }

        public FormulaFunction FormulaVertexToFunction(CellVertex formulaVertex)
        {
            return new FormulaFunction(formulaVertex,
                ParseTreeNodeToExpression(formulaVertex.ParseTree, formulaVertex.Formula));
        }

        public OutputFieldFunction OutputFieldVertexToFunction(CellVertex outputFieldVertex)
        {
            var verticesInFunction = outputFieldVertex.GetReachableVertices();

            var sortedVerticesInFunction = TopologicalSort(verticesInFunction);
            var statements = sortedVerticesInFunction
                .OfType<CellVertex>()
                .Where(vertex => Graph.FormulaFunctionDictionary.ContainsKey(vertex))
                .Select(vertex =>
                    new FunctionInvocationStatement(vertex.Name, vertex.CellType, Graph.FormulaFunctionDictionary[vertex]))
                .ToArray();

            return new OutputFieldFunction(outputFieldVertex, statements);
        }

        public List<Vertex> TopologicalSort(HashSet<Vertex> vertices)
        {
            var sortedVertices = new List<Vertex>();
            // only consider parents from the same class
            var verticesWithNoParents = vertices.Where(v => v.Parents.Count(vertices.Contains) == 0)
                .ToHashSet();

            var edges = new HashSet<(Vertex from, Vertex to)>();
            foreach (var (vertex, child) in vertices
                .SelectMany(vertex => vertex.Children
                    .Where(vertices.Contains)
                    .Select(child => (vertex, child))))
                edges.Add((vertex, child));

            while (verticesWithNoParents.Count > 0)
            {
                // get first vertex that has the lowest number of children
                // reverse first to preserve argument order
                var currentVertex = verticesWithNoParents.Reverse().OrderBy(v => v.Children.Count).First();

                verticesWithNoParents.Remove(currentVertex);
                // do not re-add deleted vertices
                if (vertices.Contains(currentVertex)) sortedVertices.Add(currentVertex);

                foreach (var childOfCurrentVertex in edges
                    .Where(edge => edge.from == currentVertex)
                    .Select(edge => edge.to)
                    .ToList())
                {
                    edges.Remove((currentVertex, childOfCurrentVertex));
                    if (edges.Count(x => x.to == childOfCurrentVertex) == 0)
                        verticesWithNoParents.Add(childOfCurrentVertex);
                }
            }

            if (edges.Count != 0)
                Logger.Log(LogItemType.Error, $"Error during topological sort! Graph has at least one cycle?");

            // reverse order
            sortedVertices.Reverse();
            return sortedVertices;
        }

        public Expression ParseTreeNodeToExpression(ParseTreeNode node, string formula)
        {
            // Non-Terminals
            if (node.Term is NonTerminal)
            {
                switch (node.Term.Name)
                {
                    case "Cell":
                    {
                        var cellVertex = _addressToVertexDictionary[(null, node.FindTokenAndGetText())];
                        return CellVertexToReference(cellVertex);
                    }
                    case "Constant":
                    {
                        string constant = node.FindTokenAndGetText();
                        if (constant.Contains("%") && !constant.Any(char.IsLetter))
                        {
                            // e.g. IF(A1>0, "1%", "2%")  (user entered "1%" instead of 1%)
                            return new Constant(constant, true);
                        }

                        switch (node.ChildNodes[0].Term.Name)
                        {
                            case "Bool":
                                return new Constant(bool.Parse(constant), CellType.Bool);
                            case "Number":
                                return new Constant(constant);
                            case "Text":
                                return new Constant(constant, CellType.Text);
                            case "Error":
                                return new Constant(null, CellType.Error);
                        }
                        throw new Exception("Unidentified constant term name");
                    }

                    case "FormulaWithEq":
                    {
                        return ParseTreeNodeToExpression(node.ChildNodes[1], formula);
                    }
                    case "Formula":
                    {
                        // for rule OpenParen + Formula + CloseParen
                        return ParseTreeNodeToExpression(
                            node.ChildNodes.Count == 3 ? node.ChildNodes[1] : node.ChildNodes[0], formula);
                    }
                    case "FunctionCall":
                    {
                        return FunctionToExpressionFunction(node.GetFunction(), node.GetFunctionArguments().ToArray(), formula);
                    }
                    case "NamedRange":
                    {
                        return NamedRangeToReference(node);
                    }
                    case "Reference":
                    {
                        if (node.ChildNodes.Count == 1)
                            return ParseTreeNodeToExpression(node.ChildNodes[0], formula);
                        // External cells
                        if (node.ChildNodes.Count == 2)
                        {
                            if (node.ChildNodes[1].Term.Name == "NamedRange")
                                return ParseTreeNodeToExpression(node.ChildNodes[1], formula);
                            var prefix = node.ChildNodes[0];
                            var sheetName = prefix.ChildNodes.Count == 2
                                ? prefix.ChildNodes[1].FindTokenAndGetText()
                                : prefix.FindTokenAndGetText();
                            sheetName = sheetName.FormatSheetName();
                            var address = node.ChildNodes[1].FindTokenAndGetText();

                            return CellVertexToReference(_addressToVertexDictionary[(sheetName, address)]);
                        }
                        
                        throw new NotImplementedException("Rule not implemented");
                    }
                    case "ReferenceItem":
                    {
                        return node.ChildNodes.Count == 1
                            ? ParseTreeNodeToExpression(node.ChildNodes[0], formula)
                            : throw new NotImplementedException("Rule not implemented");
                    }
                    case "ReferenceFunctionCall":
                    {
                        var function = node.GetFunction();
                        if (function != ":")
                            return FunctionToExpressionFunction(function, node.GetFunctionArguments().ToArray(), formula);

                        // node is a range
                        return RangeToRangeReference(node, formula);
                    }
                    default:
                    {
                        throw new Exception($"Parse token {node.Term.Name} is not implemented yet!");
                    }
                }
            }

            // Terminals
            switch (node.Term.Name)
            {
                // Not implemented
                default:
                {
                    throw new Exception($"Terminal {node.Term.Name} is not implemented yet!");
                }
            }
        }

        private ReferenceOrConstant CellVertexToReference(CellVertex cellVertex)
        {
            if (Graph.ConstantsAndConstantFormulas.Any(tuple => tuple.cellVertex == cellVertex))
                return new GlobalCellReference(cellVertex);
            return new InputReference(cellVertex);
        }

        private Function FunctionToExpressionFunction(string functionName, ParseTreeNode[] arguments, string formula)
        {
            return new Function(functionName, 
                arguments.Select(argument => ParseTreeNodeToExpression(argument, formula)).ToArray());
        }

        private GlobalRangeReference RangeToRangeReference(ParseTreeNode node, string formula)
        {
            var range = node.NodeToString(formula);
            if (_rangeDictionary.TryGetValue(range, out var rangeVertex))
                return new GlobalRangeReference(rangeVertex);
            throw new Exception($"Did not find variable for range {range}");
        }

        private ReferenceOrConstant NamedRangeToReference(ParseTreeNode node)
        {
            var namedRangeName = node.FindTokenAndGetText();
            if (!_nameDictionary.TryGetValue(namedRangeName, out var namedRangeVertex))
                throw new Exception($"Did not find variable for range {namedRangeName}");

            if (namedRangeVertex is CellVertex cellVertex)
                return new GlobalCellReference(cellVertex);
            return new GlobalRangeReference((RangeVertex) namedRangeVertex);

        }
    }
}
