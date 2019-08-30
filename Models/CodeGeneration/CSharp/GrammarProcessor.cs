using System.Linq;
using Irony.Parsing;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Thesis.Models.VertexTypes;
using XLParser;
using static Microsoft.CodeAnalysis.CSharp.SyntaxFactory;

namespace Thesis.Models.CodeGeneration.CSharp
{
    public partial class CSharpGenerator
    {
        private ExpressionSyntax TreeNodeToExpression(ParseTreeNode node, CellVertex currentVertex)
        {
            // Non-Terminals
            if (node.Term is NonTerminal nt)
            {
                switch (node.Term.Name)
                {
                    case "Cell":
                        return VariableReferenceFromAddressToExpression(node.FindTokenAndGetText(), currentVertex);
                    case "Constant":
                        string constant = node.FindTokenAndGetText();
                        if (constant.Contains("%") && !constant.Any(char.IsLetter))
                        {
                            // e.g. IF(A1>0, "1%", "2%")  (user entered "1%" instead of 1%)
                            currentVertex.CellType = CellType.Number;
                            constant = constant.Replace("%", "*0.01").Replace("\"", "");
                        }

                        if (node.ChildNodes[0].Term.Name == "Bool")
                            return ParseExpression(constant.ToLower());
                        return ParseExpression(constant);
                    case "FormulaWithEq":
                        return TreeNodeToExpression(node.ChildNodes[1], currentVertex);
                    case "Formula":
                        // for rule OpenParen + Formula + CloseParen
                        return TreeNodeToExpression(
                            node.ChildNodes.Count == 3 ? node.ChildNodes[1] : node.ChildNodes[0], currentVertex);
                    case "FunctionCall":
                        return FunctionToExpression(node.GetFunction(), node.GetFunctionArguments().ToArray(),
                            currentVertex);
                    case "NamedRange":
                        return NamedRangeToExpression(node, currentVertex);
                    case "Reference":
                        if (node.ChildNodes.Count == 1)
                            return TreeNodeToExpression(node.ChildNodes[0], currentVertex);
                        // External cells
                        if (node.ChildNodes.Count == 2)
                        {
                            var prefix = node.ChildNodes[0];
                            var sheetName = prefix.ChildNodes.Count == 2
                                ? prefix.ChildNodes[1].FindTokenAndGetText()
                                : prefix.FindTokenAndGetText();
                            sheetName = sheetName.FormatSheetName();
                            var address = node.ChildNodes[1].FindTokenAndGetText();
                            return MemberAccessExpression(
                                SyntaxKind.SimpleMemberAccessExpression,
                                IdentifierName("External"),
                                IdentifierName(Vertex.GenerateExternalVariableName(sheetName, address)));
                        }

                        return RuleNotImplemented(nt);
                    case "ReferenceItem":
                        return node.ChildNodes.Count == 1
                            ? TreeNodeToExpression(node.ChildNodes[0], currentVertex)
                            : RuleNotImplemented(nt);
                    case "ReferenceFunctionCall":
                        var function = node.GetFunction();
                        if (function != ":")
                            return FunctionToExpression(function, node.GetFunctionArguments().ToArray(), currentVertex);

                        // node is a range
                        return RangeToExpression(node, currentVertex);
                    case "HRange":
                    case "VRange":
                        return RangeToExpression(node.Parent(currentVertex.ParseTree), currentVertex);
                    case "UDFunctionCall":
                        // Not implemented
                        return CommentExpression($"User-defined function {node.GetFunction()} call here");
                    default:
                        return CommentExpression($"Parse token {node.Term.Name} is not implemented yet! {node}", true);
                }
            }

            // Terminals
            switch (node.Term.Name)
            {
                // Not implemented
                default:
                    return CommentExpression($"Terminal {node.Term.Name} is not implemented yet!", true);
            }
        }

        private ExpressionSyntax RuleNotImplemented(NonTerminal nt)
        {
            return CommentExpression($"Rule {nt.Rule} is not implemented for {nt.Name}!", true);
        }

        private CellType? GetType(ParseTreeNode node, Vertex currentVertex)
        {
            ParseTreeNode[] childNodesToContinueWith = node.ChildNodes.ToArray();

            if (node.Term.Name == "ReferenceFunctionCall")
            {
                var arguments = node.ChildNodes[1];
                if (arguments.ChildNodes.Count == 3) // IF
                    return IsSameType(arguments.ChildNodes[1], arguments.ChildNodes[2], currentVertex);
                if (arguments.ChildNodes.Count == 2)
                    return IsTypeBoolean(arguments.ChildNodes[1], currentVertex) ? CellType.Bool : (CellType?)null;
                return null;

            }
            if (node.Term.Name == "FunctionCall")
            {
                var function = node.GetFunction();
                if (_functionToCellTypeDictionary.TryGetValue(function, out var functionType)
                    && functionType != CellType.Unknown)
                    return functionType;
                childNodesToContinueWith = node.GetFunctionArguments().ToArray();
            }

            if (node.Term.Name == "UDFunctionCall")
            {
                return null;
            }

            var text = node.FindTokenAndGetText();
            if (node.Term.Name == "Cell")
            {
                if (!AddressToVertexDictionary.TryGetValue((currentVertex.WorksheetName, text), out var referencedVertex))
                    return null;
                return _useDynamic.Contains(referencedVertex) ? (CellType?)null : referencedVertex.CellType;
            }

            if (node.Term.Name == "Constant")
            {
                if (text.Contains("%")) return CellType.Number;
                if (text.ToLower().Contains("true") || text.ToLower().Contains("false")) return CellType.Bool;
                if (!text.Contains("\"")) return CellType.Number;
                return CellType.Text;
            }

            var childTypes = childNodesToContinueWith.Select(node1 => GetType(node1, currentVertex)).Distinct().ToList();
            if (childTypes.All(c => c.HasValue) && childTypes.Count == 1 && childTypes[0].Value != CellType.Unknown)
                return childTypes[0].Value;
            return null;
        }

        private CellType? IsSameType(ParseTreeNode a, ParseTreeNode b, Vertex currentVertex)
        {
            var typeA = GetType(a, currentVertex);
            var typeB = GetType(b, currentVertex);
            if (typeA.HasValue && typeA.Value != CellType.Unknown && typeB.HasValue && typeB.Value != CellType.Unknown
                && typeA.Value == typeB.Value)
            {
                return typeA.Value;
            }
            return null;
        }

        private bool IsTypeBoolean(ParseTreeNode a, Vertex currentVertex)
        {
            var typeOfIf = GetType(a, currentVertex);
            return typeOfIf.HasValue && typeOfIf.Value == CellType.Bool;
        }
    }
}