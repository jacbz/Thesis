using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Irony.Parsing;
using XLParser;

namespace Thesis.Models.CodeGenerators
{
    public class CSharpGenerator : CodeGenerator
    {
        public CSharpGenerator(GeneratedClass generatedClass, Dictionary<string, Vertex> addressToVertexDictionary)
            : base(generatedClass, addressToVertexDictionary)
        {
        }

        public override string ClassToCode()
        {
            var output = new List<string>();
            output.Add($"class {generatedClass.Name}");
            output.Add("{");

            var properties = new List<string>();
            var method = new List<string>();

            for (int i = generatedClass.Vertices.Count - 1; i >= 0; i--)
            {
                var vertex = generatedClass.Vertices[i];
                if (vertex.Type == CellType.Formula)
                    method.AddRange(VertexToCode(vertex));
                else
                    properties.AddRange(VertexToCode(vertex));
            }

            output.AddRange(FormatLines(properties, 1));
            output.Add("\n");

            output.Add(FormatLine($"{generatedClass.Name}()", 1));
            output.Add(FormatLine("{", 1));
            output.AddRange(FormatLines(method, 2));
            output.Add(FormatLine( "}", 1));
            output.Add("}");
            return string.Join("\n", output);
        }

        public override IEnumerable<string> VertexToCode(Vertex vertex)
        {
            var label = vertex.LabelOrAddress;

            switch (vertex.Type)
            {
                case CellType.Bool:
                    yield return $"static bool {label} = {vertex.Value};";
                    break;
                case CellType.Date:
                    yield return $"static DateTime {label} = DateTime.Parse({vertex.Value});";
                    break;
                case CellType.Formula:
                    yield return $"var {label} = {TreeNodeToCode(vertex.ParseTree, 0)};";
                    break;
                case CellType.Number:
                    yield return $"static decimal {label} = {vertex.Value};";
                    break;
                case CellType.Text:
                    yield return $"static string {label} = \"{vertex.Value}\";";
                    break;
            }
        }

        private string FormatAddress(string address)
        {
            if (addressToVertexDictionary.TryGetValue(address, out var vertex))
                return vertex.Class == generatedClass ? vertex.LabelOrAddress: $"{vertex.Class.Name}.{vertex.LabelOrAddress}";

            return $"/* Error: {address} not found*/";
        }

        private string TreeNodeToCode(ParseTreeNode node, int depth)
        {
            depth++;

            // Non-Terminals
            if (node.Term is NonTerminal nt)
            {
                switch (node.Term.Name)
                {
                    case "Cell":
                        return FormatAddress(node.FindTokenAndGetText());
                    case "Constant":
                        return node.FindTokenAndGetText();
                    case "FormulaWithEq":
                        return TreeNodeToCode(node.ChildNodes[1], depth);
                    case "Formula":
                        // for rule OpenParen + Formula + CloseParen
                        return TreeNodeToCode(node.ChildNodes.Count == 3 ? node.ChildNodes[1] : node.ChildNodes[0], depth);
                    case "FunctionCall":
                        return ParseFunction(node.GetFunction(), node.GetFunctionArguments().ToArray(), depth);
                    case "Reference":
                        if (node.ChildNodes.Count == 1)
                            return TreeNodeToCode(node.ChildNodes[0], depth);
                        if (node.ChildNodes.Count == 2)
                        {
                            var prefix = node.ChildNodes[0];
                            var refName = prefix.ChildNodes.Count == 2 ? prefix.ChildNodes[1].FindTokenAndGetText() : prefix.FindTokenAndGetText();
                            return $"ExternalRef(\"[{refName}\", {TreeNodeToCode(node.ChildNodes[1], depth)})";
                        }
                        return RuleNotImplemented(nt);
                    case "ReferenceItem":
                        return node.ChildNodes.Count == 1 ? TreeNodeToCode(node.ChildNodes[0], depth) : RuleNotImplemented(nt);
                    case "ReferenceFunctionCall":
                        return ParseFunction(node.GetFunction(), node.GetFunctionArguments().ToArray(), depth);
                    case "UDFunctionCall":
                        return $"{node.GetFunction()}({string.Join(", ", node.GetFunctionArguments().Select(a => TreeNodeToCode(a, depth)))})";
                    // Not implemented
                    default:
                        return $"/* Parse token {node.Term.Name} is not implemented yet! {node} */";
                }
            }

            // Terminals
            switch (node.Term.Name)
            {
                // Not implemented
                default:
                    return $"/* Terminal {node.Term.Name} is not implemented yet! */";
            }

        }

        private string RuleNotImplemented(NonTerminal nt)
        {
            return $"/* Rule {nt.Rule} is not implemented for {nt.Name}! */";
        }

        private string ParseFunction(string functionName, ParseTreeNode[] arguments, int depth)
        {
            switch (functionName)
            {
                // arithmetic
                case "+":
                    if (arguments.Length != 2) return FunctionError(functionName, arguments);
                    return TreeNodeToCode(arguments[0], depth) + " + " + TreeNodeToCode(arguments[1], depth);
                case "-":
                    if (arguments.Length != 2) return FunctionError(functionName, arguments);
                    return TreeNodeToCode(arguments[0], depth) + " - " + TreeNodeToCode(arguments[1], depth);
                case "*":
                    if (arguments.Length != 2) return FunctionError(functionName, arguments);
                    return TreeNodeToCode(arguments[0], depth) + " * " + TreeNodeToCode(arguments[1], depth);
                case "/":
                    if (arguments.Length != 2) return FunctionError(functionName, arguments);
                    return TreeNodeToCode(arguments[0], depth) + " / " + TreeNodeToCode(arguments[1], depth);
                case "%":
                    if (arguments.Length != 1) return FunctionError(functionName, arguments);
                    return TreeNodeToCode(arguments[0], depth) + " * 0.01";
                case "^":
                    if (arguments.Length != 2) return FunctionError(functionName, arguments);
                    return $"Math.Pow({TreeNodeToCode(arguments[0], depth)}, {TreeNodeToCode(arguments[1], depth)})";
                case "ROUND":
                    if (arguments.Length != 2) return FunctionError(functionName, arguments);
                    return $"Math.Round({TreeNodeToCode(arguments[0], depth)}, {TreeNodeToCode(arguments[1], depth)}, MidpointRounding.AwayFromZero)";
                case "SUM":
                    return string.Join(".Concat(", arguments.Select(a => TreeNodeToCode(a, depth))) 
                           + (arguments.Length > 1 ? ")" : "") + ".Sum()";

                // conditionals
                case "IF":
                    if (arguments.Length == 2)
                        return $"(({TreeNodeToCode(arguments[0], depth)})\n" +
                               FormatLine($"? {TreeNodeToCode(arguments[1], depth)}\n", depth / 2) +
                               FormatLine($": null)", depth / 2);
                    if (arguments.Length == 3)
                        return $"(({TreeNodeToCode(arguments[0], depth)})\n" +
                               FormatLine($"? {TreeNodeToCode(arguments[1], depth)}\n", depth / 2) +
                               FormatLine($": {TreeNodeToCode(arguments[2], depth)})", depth / 2);
                    return FunctionError(functionName, arguments);
                case "=":
                    if (arguments.Length != 2) return FunctionError(functionName, arguments);
                    return $"{TreeNodeToCode(arguments[0], depth)} == {TreeNodeToCode(arguments[1], depth)}";
                case "<>":
                    if (arguments.Length != 2) return FunctionError(functionName, arguments);
                    return $"{TreeNodeToCode(arguments[0], depth)} != {TreeNodeToCode(arguments[1], depth)}";
                case "<":
                    if (arguments.Length != 2) return FunctionError(functionName, arguments);
                    return $"{TreeNodeToCode(arguments[0], depth)} < {TreeNodeToCode(arguments[1], depth)}";
                case "<=":
                    if (arguments.Length != 2) return FunctionError(functionName, arguments);
                    return $"{TreeNodeToCode(arguments[0], depth)} <= {TreeNodeToCode(arguments[1], depth)}";
                case ">=":
                    if (arguments.Length != 2) return FunctionError(functionName, arguments);
                    return $"{TreeNodeToCode(arguments[0], depth)} >= {TreeNodeToCode(arguments[1], depth)}";
                case ">":
                    if (arguments.Length != 2) return FunctionError(functionName, arguments);
                    return $"{TreeNodeToCode(arguments[0], depth)} > {TreeNodeToCode(arguments[1], depth)}";

                // strings
                case "&":
                case "CONCATENATE":
                    if (arguments.Length < 2 ) return FunctionError(functionName, arguments);
                    return string.Join(" + ", arguments.Select(a => TreeNodeToCode(a, depth)));

                // ranges
                case ":":
                    if (arguments.Length != 2) return FunctionError(functionName, arguments);
                    return $"GetRange({TreeNodeToCode(arguments[0], depth)}, {TreeNodeToCode(arguments[1], depth)})";

                // other
                case "TODAY":
                    return "DateTime.Now";
                case "SECOND":
                case "MINUTE":
                case "HOUR":
                case "DAY":
                case "MONTH":
                case "YEAR":
                    if (arguments.Length != 1) return FunctionError(functionName, arguments);
                    return $"DateTime.Parse({TreeNodeToCode(arguments[0], depth)})." + functionName.Substring(0,1) + functionName.Substring(1).ToLower();

                default:
                    return $"/* Function {functionName} not implemented yet! Args: \n " +
                           $"{string.Join("\n", arguments.Select(a => TreeNodeToCode(a, depth)))} */";
            }
        }

        private string FunctionError(string functionName, ParseTreeNode[] arguments)
        {
            return "Error in " + functionName;
        }

//        private List<string> Error(ParseTreeNode node)
//        {
//            ExcelFormulaGrammar
//            return new List<string> { $"// Parsing rules for {node.ToString()} not implemented yet! {node.ToString()}" };
//        }

        public string FormatLine(string s, int indentLevel = 1)
        {
            return Indent(indentLevel) + s;
        }

        public IEnumerable<string> FormatLines(IEnumerable<string> list, int indentLevel)
        {
            return list.Select(i => FormatLine(i.Replace("\n", "\n".PadRight(indentLevel * 4)), indentLevel));
        }

        public override string Indent(int level)
        {
            return "".PadLeft(level * 4);
        }
    }
}
