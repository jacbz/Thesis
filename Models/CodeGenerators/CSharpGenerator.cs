using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Irony.Parsing;
using Syncfusion.XlsIO;
using XLParser;

namespace Thesis.Models.CodeGenerators
{
    public class CSharpGenerator : CodeGenerator
    {
        public CSharpGenerator(List<GeneratedClass> generatedClasses, Dictionary<string, Vertex> addressToVertexDictionary)
            : base(generatedClasses, addressToVertexDictionary)
        {
        }

        private protected override string GetMainClass()
        {
            var output = new List<string>();
            output.Add("using System;\n");
            output.Add($"public class ThesisResult");
            output.Add("{");
            output.Add(FormatLine("static void Main(string[] args)", 1));
            output.Add(FormatLine("{", 1));

            output.AddRange(FormatLines(generatedClasses.Where(c => c.OutputVertex == null).Select(c => $"{c.Name} {c.Name.ToLower()} = new {c.Name}();"), 2));
            output.AddRange(FormatLines(generatedClasses.Where(c => c.OutputVertex != null).Select(c => $"{CellTypeToTypeString(c.OutputVertex.Type)} {c.OutputVertex.NameInCode} = {c.Name}.Calculate();"), 2));

            output.Add(FormatLine("}", 1));

            output.Add("}");
            return string.Join("\n", output);
        }

        private protected override string ClassToCode(GeneratedClass generatedClass)
        {
            var output = new List<string>();
            output.Add($"public class {generatedClass.Name}");
            output.Add("{");

            var properties = new List<string>();
            var method = new List<string>();

            var outputVertex = generatedClass.OutputVertex;

            for (int i = generatedClass.Vertices.Count - 1; i >= 0; i--)
            {
                var vertex = generatedClass.Vertices[i];
                if (vertex.NodeType == NodeType.Constant)
                {
                    properties.Add(VertexToCode(vertex, generatedClass));
                }
                else
                {
                    if (outputVertex == null)
                    {
                        properties.Add($"public static {CellTypeToTypeString(vertex.Type)} {vertex.NameInCode};");
                    }
                    method.Add(VertexToCode(vertex, generatedClass));
                }
            }

            output.AddRange(FormatLines(properties, 1));
            output.Add("");

            output.Add(FormatLine($"public {(generatedClass.OutputVertex == null ? generatedClass.Name : "static " + CellTypeToTypeString(outputVertex.Type) + " Calculate")}()", 1));
            output.Add(FormatLine("{", 1));
            output.AddRange(FormatLines(method, 2));
            if (outputVertex != null)
                output.Add(FormatLine($"return {outputVertex.NameInCode};", 2));
            output.Add(FormatLine( "}", 1));
            output.Add("}");
            return string.Join("\n", output);
        }

        public string VertexToCode(Vertex vertex, GeneratedClass generatedClass)
        {
            if (vertex.NodeType == NodeType.Constant)
            {
                string accessModifier = generatedClass.OutputVertex != null ? "static " : "public static ";
                switch (vertex.Type)
                {
                    case CellType.Bool:
                        return accessModifier + $"{CellTypeToTypeString(vertex.Type)} {vertex.NameInCode} = {vertex.Value};";
                    case CellType.Number:
                        string value = vertex.Value.ToString().Replace(",", ".");
                        if (value.Contains("%")) value = value.Replace("%", " * 0.01");
                        return accessModifier + $"{CellTypeToTypeString(vertex.Type)} {vertex.NameInCode} = {value};";
                    case CellType.Text:
                        return accessModifier + $"{CellTypeToTypeString(vertex.Type)} {vertex.NameInCode} = \"{vertex.Value}\";";
                    case CellType.Date:
                        return accessModifier + $"{CellTypeToTypeString(vertex.Type)} {vertex.NameInCode} = DateTime.Parse({vertex.Value});";
                    case CellType.Unknown:
                        return accessModifier + $"{CellTypeToTypeString(vertex.Type)} {vertex.NameInCode} = null;";
                    default:
                        return "";
                }
            }

            return $"{CellTypeToTypeString(vertex.Type)} {vertex.NameInCode} = {TreeNodeToCode(vertex.ParseTree, vertex, 0)};";
        }

        public static string CellTypeToTypeString(CellType cellType)
        {
            switch (cellType)
            {
                case CellType.Bool:
                    return "bool";
                case CellType.Date:
                    return "DateTime";
                case CellType.Number:
                    return "double";
                case CellType.Text:
                    return "string";
                case CellType.Unknown:
                    return "object";
                default:
                    return "";
            }
        }
        public static string CellTypeToNullValue(CellType cellType)
        {
            switch (cellType)
            {
                case CellType.Bool:
                    return "false";
                case CellType.Unknown:
                case CellType.Date:
                    return "null";
                case CellType.Number:
                    return "0";
                case CellType.Text:
                default:
                    return "";
            }
        }

        private string FormatAddress(string address, Vertex vertex)
        {
            if (addressToVertexDictionary.TryGetValue(address, out var v))
                return vertex.Class == v.Class ? v.NameInCode: $"{v.Class.Name}.{v.NameInCode}";

            return $"/* Error: {address} not found */";
        }

        private string TreeNodeToCode(ParseTreeNode node, Vertex vertex, int depth)
        {
            depth++;

            // Non-Terminals
            if (node.Term is NonTerminal nt)
            {
                switch (node.Term.Name)
                {
                    case "Cell":
                        return FormatAddress(node.FindTokenAndGetText(), vertex);
                    case "Constant":
                        return node.FindTokenAndGetText();
                    case "FormulaWithEq":
                        return TreeNodeToCode(node.ChildNodes[1], vertex, depth);
                    case "Formula":
                        // for rule OpenParen + Formula + CloseParen
                        return TreeNodeToCode(node.ChildNodes.Count == 3 ? node.ChildNodes[1] : node.ChildNodes[0], vertex, depth);
                    case "FunctionCall":
                        return ParseFunction(node.GetFunction(), node.GetFunctionArguments().ToArray(), vertex, depth);
                    case "Reference":
                        if (node.ChildNodes.Count == 1)
                            return TreeNodeToCode(node.ChildNodes[0], vertex, depth);
                        if (node.ChildNodes.Count == 2)
                        {
                            var prefix = node.ChildNodes[0];
                            var refName = prefix.ChildNodes.Count == 2 ? prefix.ChildNodes[1].FindTokenAndGetText() : prefix.FindTokenAndGetText();
                            return $"ExternalRef(\"[{refName}\", {TreeNodeToCode(node.ChildNodes[1], vertex, depth)})";
                        }
                        return RuleNotImplemented(nt);
                    case "ReferenceItem":
                        return node.ChildNodes.Count == 1 ? TreeNodeToCode(node.ChildNodes[0], vertex, depth) : RuleNotImplemented(nt);
                    case "ReferenceFunctionCall":
                        return ParseFunction(node.GetFunction(), node.GetFunctionArguments().ToArray(), vertex, depth);
                    case "UDFunctionCall":
                        return $"{node.GetFunction()}({string.Join(", ", node.GetFunctionArguments().Select(a => TreeNodeToCode(a, vertex, depth)))})";
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

        private string ParseFunction(string functionName, ParseTreeNode[] arguments, Vertex vertex, int depth)
        {
            switch (functionName)
            {
                // arithmetic
                case "+":
                    if (arguments.Length != 2) return FunctionError(functionName, arguments);
                    return TreeNodeToCode(arguments[0], vertex, depth) + " + " + TreeNodeToCode(arguments[1], vertex, depth);
                case "-":
                    if (arguments.Length != 2) return FunctionError(functionName, arguments);
                    return TreeNodeToCode(arguments[0], vertex, depth) + " - " + TreeNodeToCode(arguments[1], vertex, depth);
                case "*":
                    if (arguments.Length != 2) return FunctionError(functionName, arguments);
                    return TreeNodeToCode(arguments[0], vertex, depth) + " * " + TreeNodeToCode(arguments[1], vertex, depth);
                case "/":
                    if (arguments.Length != 2) return FunctionError(functionName, arguments);
                    return TreeNodeToCode(arguments[0], vertex, depth) + " / " + TreeNodeToCode(arguments[1], vertex, depth);
                case "%":
                    if (arguments.Length != 1) return FunctionError(functionName, arguments);
                    return TreeNodeToCode(arguments[0], vertex, depth) + " * 0.01";
                case "^":
                    if (arguments.Length != 2) return FunctionError(functionName, arguments);
                    return $"Math.Pow({TreeNodeToCode(arguments[0], vertex, depth)}, {TreeNodeToCode(arguments[1], vertex, depth)})";
                case "ROUND":
                    if (arguments.Length != 2) return FunctionError(functionName, arguments);
                    return $"Math.Round({TreeNodeToCode(arguments[0], vertex, depth)}, {TreeNodeToCode(arguments[1], vertex, depth)}, MidpointRounding.AwayFromZero)";
                case "SUM":
                    return string.Join(".Concat(", arguments.Select(a => TreeNodeToCode(a, vertex, depth))) 
                           + (arguments.Length > 1 ? ")" : "") + ".Sum()";

                // conditionals
                case "IF":
                    if (arguments.Length == 2)
                        return $"(({TreeNodeToCode(arguments[0], vertex, depth)})\n" +
                               FormatLine($"? {TreeNodeToCode(arguments[1], vertex, depth)}\n", depth / 2) +
                               FormatLine($": {CellTypeToNullValue(vertex.Type)})", depth / 2);
                    if (arguments.Length == 3)
                        return $"(({TreeNodeToCode(arguments[0], vertex, depth)})\n" +
                               FormatLine($"? {TreeNodeToCode(arguments[1], vertex, depth)}\n", depth / 2) +
                               FormatLine($": {TreeNodeToCode(arguments[2], vertex, depth)})", depth / 2);
                    return FunctionError(functionName, arguments);
                case "=":
                    if (arguments.Length != 2) return FunctionError(functionName, arguments);
                    return $"{TreeNodeToCode(arguments[0], vertex, depth)} == {TreeNodeToCode(arguments[1], vertex, depth)}";
                case "<>":
                    if (arguments.Length != 2) return FunctionError(functionName, arguments);
                    return $"{TreeNodeToCode(arguments[0], vertex, depth)} != {TreeNodeToCode(arguments[1], vertex, depth)}";
                case "<":
                    if (arguments.Length != 2) return FunctionError(functionName, arguments);
                    return $"{TreeNodeToCode(arguments[0], vertex, depth)} < {TreeNodeToCode(arguments[1], vertex, depth)}";
                case "<=":
                    if (arguments.Length != 2) return FunctionError(functionName, arguments);
                    return $"{TreeNodeToCode(arguments[0], vertex, depth)} <= {TreeNodeToCode(arguments[1], vertex, depth)}";
                case ">=":
                    if (arguments.Length != 2) return FunctionError(functionName, arguments);
                    return $"{TreeNodeToCode(arguments[0], vertex, depth)} >= {TreeNodeToCode(arguments[1], vertex, depth)}";
                case ">":
                    if (arguments.Length != 2) return FunctionError(functionName, arguments);
                    return $"{TreeNodeToCode(arguments[0], vertex, depth)} > {TreeNodeToCode(arguments[1], vertex, depth)}";

                // strings
                case "&":
                case "CONCATENATE":
                    if (arguments.Length < 2 ) return FunctionError(functionName, arguments);
                    return string.Join(" + ", arguments.Select(a => TreeNodeToCode(a, vertex, depth)));

                // ranges
                case ":":
                    if (arguments.Length != 2) return FunctionError(functionName, arguments);
                    return $"GetRange({TreeNodeToCode(arguments[0], vertex, depth)}, {TreeNodeToCode(arguments[1], vertex, depth)})";

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
                    return $"DateTime.Parse({TreeNodeToCode(arguments[0], vertex, depth)})." + functionName.Substring(0,1) + functionName.Substring(1).ToLower();

                default:
                    return $"/* Function {functionName} not implemented yet! Args: \n " +
                           $"{string.Join("\n", arguments.Select(a => TreeNodeToCode(a, vertex, depth)))} */";
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

        public static string FormatLine(string s, int indentLevel = 1)
        {
            return Indent(indentLevel) + s;
        }

        public static IEnumerable<string> FormatLines(IEnumerable<string> list, int indentLevel)
        {
            return list.Select(i => FormatLine(i.Replace("\n", "\n".PadRight(indentLevel * 4)), indentLevel));
        }

        public static string Indent(int level)
        {
            return "".PadLeft(level * 4);
        }
    }
}
