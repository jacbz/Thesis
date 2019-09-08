using System.Linq;
using Irony.Parsing;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Thesis.Models.FunctionGeneration;
using Thesis.Models.VertexTypes;
using XLParser;
using static Microsoft.CodeAnalysis.CSharp.SyntaxFactory;

namespace Thesis.Models.CodeGeneration.CSharp
{
    public partial class CSharpGenerator
    {
        private bool _forceConstantsIntoDecimal = false;

        private ExpressionSyntax RuleNotImplemented(NonTerminal nt)
        {
            return CommentExpression($"Rule {nt.Rule} is not implemented for {nt.Name}!", true);
        }

        private CellType? GetType(ParseTreeNode node, Vertex currentVertex)
        {
            //ParseTreeNode[] childNodesToContinueWith = node.ChildNodes.ToArray();

            //if (node.Term.Name == "ReferenceFunctionCall")
            //{
            //    var arguments = node.ChildNodes[1];
            //    if (arguments.ChildNodes.Count == 3) // IF
            //        return IsSameType(arguments.ChildNodes[1], arguments.ChildNodes[2], currentVertex);
            //    if (arguments.ChildNodes.Count == 2)
            //        return IsTypeBoolean(arguments.ChildNodes[1], currentVertex) ? CellType.Bool : (CellType?)null;
            //    return null;

            //}
            //if (node.Term.Name == "FunctionCall")
            //{
            //    var function = node.GetFunction();
            //    if (_functionToCellTypeDictionary.TryGetValue(function, out var functionType)
            //        && functionType != CellType.Unknown)
            //        return functionType;
            //    childNodesToContinueWith = node.GetFunctionArguments().ToArray();
            //}

            //if (node.Term.Name == "UDFunctionCall")
            //{
            //    return null;
            //}

            //var text = node.FindTokenAndGetText();
            //if (node.Term.Name == "Cell")
            //{
            //    if (!AddressToVertexDictionary.TryGetValue((currentVertex.WorksheetName, text), out var referencedVertex))
            //        return null;
            //    return _useDynamic.Contains(referencedVertex) ? (CellType?)null : referencedVertex.CellType;
            //}

            //if (node.Term.Name == "Constant")
            //{
            //    if (text.Contains("%")) return CellType.Number;
            //    if (text.ToLower().Contains("true") || text.ToLower().Contains("false")) return CellType.Bool;
            //    if (!text.Contains("\"")) return CellType.Number;
            //    return CellType.Text;
            //}

            //var childTypes = childNodesToContinueWith.Select(node1 => GetType(node1, currentVertex)).Distinct().ToList();
            //if (childTypes.All(c => c.HasValue) && childTypes.Count == 1 && childTypes[0].Value != CellType.Unknown)
            //    return childTypes[0].Value;
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