using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Irony.Parsing;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Formatting;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Formatting;
using static Microsoft.CodeAnalysis.CSharp.SyntaxFactory;
using XLParser;

namespace Thesis.Models.CodeGenerators
{
    public class CSharpGenerator : CodeGenerator
    {
        private HashSet<string> _usedVariableNames;

        public CSharpGenerator(List<GeneratedClass> generatedClasses, Dictionary<string, Vertex> addressToVertexDictionary)
            : base(generatedClasses, addressToVertexDictionary)
        {
            _usedVariableNames = new HashSet<string>();
        }

        public override async Task<string> GenerateCodeAsync(Dictionary<string, TestResult> testResults = null)
        {
            return await Task.Run(() => GenerateCode(testResults));
        }

        public string GenerateCode(Dictionary<string, TestResult> testResults = null)
        {
            Tester = new CSharpTester();

            // generate variable names for all
            GenerateVariableNamesForAll();
            VariableNameToVertexDictionary = GeneratedClasses.SelectMany(c => c.Vertices).ToDictionary(v => v.VariableName);

            // namespace Thesis
            var @namespace = NamespaceDeclaration(ParseName("Thesis")).NormalizeWhitespace();
            // using System;
            @namespace = @namespace.AddUsings(UsingDirective(ParseName("System")));

            // public class ThesisResult (Main class)
            var resultClass = GenerateResultClass(testResults);
            @namespace = @namespace.AddMembers(resultClass);

            var normalClasses = new List<MemberDeclarationSyntax>();
            var sharedClasses = new List<MemberDeclarationSyntax>();
            // normal classes must be processed first, as they can infer some Unknown types in shared classes
            foreach (var generatedClass in GeneratedClasses.OrderBy(v => v.IsSharedClass))
            {
                var newClass = GenerateClass(generatedClass, testResults);

                if (generatedClass.IsSharedClass)
                    sharedClasses.Add(newClass);
                else
                    normalClasses.Add(newClass);
            }

            // show shared classes on top for aesthetic reasons
            @namespace = @namespace
                .AddMembers(sharedClasses.ToArray())
                .AddMembers(normalClasses.ToArray());

            // format code and return as string
            var workspace = new AdhocWorkspace();
            var options = workspace.Options;
            var node = Formatter.Format(@namespace, workspace, options);
            var code = node.ToFullString();

            return code;
        }

        private ClassDeclarationSyntax GenerateClass(GeneratedClass generatedClass, Dictionary<string, TestResult> testResults = null)
        {
            // public class {generatedClass.Name}
            var newClass = ClassDeclaration(generatedClass.Name)
                .AddModifiers(Token(SyntaxKind.PublicKeyword));
            // shared classes are static
            if (generatedClass.IsSharedClass)
                newClass = newClass.AddModifiers(Token(SyntaxKind.StaticKeyword));

            // split vertex list into two
            var lookup = generatedClass.Vertices.ToLookup(v => v.NodeType == NodeType.Constant);
            var constants = lookup[true].ToList();
            var formulas = lookup[false].ToList();

            // generate assignments
            var statements = new List<StatementSyntax>();
            foreach (var formula in formulas)
            {
                var expression = TreeNodeToExpression(formula.ParseTree, formula, formula.CellType);

                if (generatedClass.IsSharedClass)
                {
                    // for shared classes: omit type name as already declared

                    // test
                    var identifier = testResults == null
                        ? IdentifierName(formula.VariableName)
                        : GenerateIdentifierWithComment(formula.VariableName,
                            testResults[formula.VariableName].ToString());

                    var assignmentExpression = AssignmentExpression(
                        SyntaxKind.SimpleAssignmentExpression,
                        identifier,
                        expression);
                    statements.Add(ExpressionStatement(assignmentExpression));
                }
                else
                {
                    // test
                    var type = testResults == null
                        ? ParseTypeName(GetTypeString(formula))
                        : GenerateTypeWithComment(GetTypeString(formula),
                            testResults[formula.VariableName].ToString());
                    var variableDeclaration = VariableDeclaration(type)
                        .AddVariables(VariableDeclarator(formula.VariableName)
                            .WithInitializer(
                                EqualsValueClause(expression)));
                    statements.Add(LocalDeclarationStatement(variableDeclaration));
                }
            }

            // generate fields
            // this happens after assignments generation because assignments generation can infer types
            // e.g. dynamic x = null can become int x = 0
            var fields = new List<MemberDeclarationSyntax>();
            foreach (var constant in constants)
            {
                // {type} {variableName} = {value};
                var expression = VertexValueToExpression(constant.CellType, constant.Value);
                var field =
                    FieldDeclaration(
                        VariableDeclaration(
                                ParseTypeName(GetTypeString(constant)))
                            .AddVariables(VariableDeclarator(constant.VariableName)
                                .WithInitializer(
                                    EqualsValueClause(expression))));
                if (generatedClass.IsSharedClass)
                    field = field.AddModifiers(
                        Token(SyntaxKind.PublicKeyword),
                        Token(SyntaxKind.StaticKeyword));
                else
                    field = field.AddModifiers(Token(SyntaxKind.PrivateKeyword));

                fields.Add(field);
            }

            // extra fields for shared classes
            if (generatedClass.IsSharedClass)
            {
                // add formula fields
                foreach (var formula in formulas)
                {
                    var formulaField = FieldDeclaration(
                            VariableDeclaration(ParseTypeName(GetTypeString(formula)))
                                .AddVariables(VariableDeclarator(formula.VariableName)))
                        .AddModifiers(
                            Token(SyntaxKind.PublicKeyword),
                            Token(SyntaxKind.StaticKeyword));
                    fields.Add(formulaField);
                }
            }

            // generate Calculate/Init method
            var method = GenerateMethod(generatedClass, statements);

            // add fields and method to class
            newClass = newClass.AddMembers(fields.ToArray()).AddMembers(method);

            Tester.ClassesCode.Add(new ClassCode(
                generatedClass.IsSharedClass,
                generatedClass.Name,
                newClass.NormalizeWhitespace().ToFullString(),
                string.Join("\n", fields.Select(f => f.NormalizeWhitespace().ToFullString())),
                string.Join("\n", statements.Select(f => f.NormalizeWhitespace().ToFullString()))));
            return newClass;
        }

        private IdentifierNameSyntax GenerateTypeWithComment(string typeString, string comment)
        {
            return IdentifierName(
                Identifier(
                    TriviaList(Comment(comment), LineFeed),
                    typeString,
                    TriviaList()));
        }

        private static IdentifierNameSyntax GenerateIdentifierWithComment(string variableName, string comment)
        {
            return IdentifierName(
                Identifier(
                    TriviaList(Comment(comment), LineFeed),
                    variableName,
                    TriviaList()));
        }

        private MemberDeclarationSyntax GenerateMethod(GeneratedClass generatedClass, List<StatementSyntax> statements)
        {
            if (generatedClass.IsSharedClass)
            {
                // public static void Init()
                var initMethod = MethodDeclaration(ParseTypeName("void"), "Init")
                    .AddModifiers(
                        Token(SyntaxKind.PublicKeyword),
                        Token(SyntaxKind.StaticKeyword))
                    .WithBody(Block(statements));
                return initMethod;
            }
            else
            {
                // return output vertex
                statements.Add(ReturnStatement(IdentifierName(generatedClass.OutputVertex.VariableName)));

                // public {type} Calculate()
                var outputField = generatedClass.OutputVertex;
                var calculateMethod = MethodDeclaration(ParseTypeName(GetTypeString(outputField)), "Calculate")
                    .AddModifiers(
                        Token(SyntaxKind.PublicKeyword))
                    .WithBody(Block(statements));
                return calculateMethod;
            }
        }

        private ClassDeclarationSyntax GenerateResultClass(Dictionary<string, TestResult> testResults = null)
        {
            var resultClass = ClassDeclaration("Result")
                .AddModifiers(Token(SyntaxKind.PublicKeyword));

            // public static void Main(string[] args)
            var mainMethod = MethodDeclaration(ParseTypeName("void"), "Main")
                .AddModifiers(
                    Token(SyntaxKind.PublicKeyword),
                    Token(SyntaxKind.StaticKeyword))
                .AddParameterListParameters(Parameter(Identifier("args"))
                    .WithType(ParseTypeName("string[]")));

            // method body
            var methodBody = new List<StatementSyntax>();
            foreach (var generatedClass in GeneratedClasses.OrderBy(v => !v.IsSharedClass))
            {
                if (generatedClass.IsSharedClass)
                {
                    // {classname}.Init()
                    methodBody.Add(ExpressionStatement(
                        InvocationExpression(
                            MemberAccessExpression(
                                SyntaxKind.SimpleMemberAccessExpression,
                                IdentifierName(generatedClass.Name),
                                IdentifierName("Init")))));
                }
                else
                {
                    // test
                    var type = testResults == null
                        ? ParseTypeName(GetTypeString(generatedClass.OutputVertex))
                        : GenerateTypeWithComment(GetTypeString(generatedClass.OutputVertex),
                            testResults[generatedClass.OutputVertex.VariableName].ToString());

                    // {type} {outputvertexname} = new {classname}().Calculate()
                    methodBody.Add(LocalDeclarationStatement(
                        VariableDeclaration(type)
                            .AddVariables(VariableDeclarator(generatedClass.OutputVertex.VariableName)
                                .WithInitializer(
                                    EqualsValueClause(
                                        InvocationExpression(
                                            MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                                                ObjectCreationExpression(
                                                        IdentifierName(generatedClass.Name))
                                                    .WithArgumentList(
                                                        ArgumentList()),
                                                IdentifierName("Calculate"))))))));
                }
            }

            mainMethod = mainMethod.WithBody(Block(methodBody));
            resultClass = resultClass.AddMembers(mainMethod);
            return resultClass;
        }

        private void GenerateVariableNamesForAll()
        {
            _usedVariableNames = CreateUsedVariableHashSet();

            foreach (var generatedClass in GeneratedClasses)
            {
                generatedClass.Name = GenerateUniqueName(generatedClass.Name, _usedVariableNames);

                foreach (var vertex in generatedClass.Vertices)
                {
                    vertex.VariableName = GenerateUniqueName(vertex.VariableName, _usedVariableNames);
                }
            }
        }

        private ExpressionSyntax TreeNodeToExpression(ParseTreeNode node, Vertex rootVertex, CellType type)
        {
            // Non-Terminals
            if (node.Term is NonTerminal nt)
            {
                switch (node.Term.Name)
                {
                    case "Cell":
                        return FormatVariableReference(node.FindTokenAndGetText(), rootVertex, type);
                    case "Constant":
                        string constant = node.FindTokenAndGetText();
                        if (constant.Contains("%"))
                        {
                            // e.g. IF(A1>0, "1%", "2%")  (user entered "1%" instead of 1%)
                            rootVertex.CellType = CellType.Number;
                            constant = constant.Replace("%", "*0.01").Replace("\"", "");
                        }
                        return ParseExpression(constant);
                    case "FormulaWithEq":
                        return TreeNodeToExpression(node.ChildNodes[1], rootVertex, type);
                    case "Formula":
                        // for rule OpenParen + Formula + CloseParen
                        return TreeNodeToExpression(node.ChildNodes.Count == 3 ? node.ChildNodes[1] : node.ChildNodes[0], rootVertex, type);
                    case "FunctionCall":
                        return FunctionToExpression(node.GetFunction(), node.GetFunctionArguments().ToArray(), rootVertex, type);
                    case "Reference":
                        if (node.ChildNodes.Count == 1)
                            return TreeNodeToExpression(node.ChildNodes[0], rootVertex, type);
                        if (node.ChildNodes.Count == 2)
                        {
                            var prefix = node.ChildNodes[0];
                            var refName = prefix.ChildNodes.Count == 2 ? prefix.ChildNodes[1].FindTokenAndGetText() : prefix.FindTokenAndGetText();
                            return ParseExpression($"ExternalRef(\"[{refName}\", {TreeNodeToExpression(node.ChildNodes[1], rootVertex, type)})");
                        }
                        return RuleNotImplemented(nt);
                    case "ReferenceItem":
                        return node.ChildNodes.Count == 1 ? TreeNodeToExpression(node.ChildNodes[0], rootVertex, type) : RuleNotImplemented(nt);
                    case "ReferenceFunctionCall":
                        return FunctionToExpression(node.GetFunction(), node.GetFunctionArguments().ToArray(), rootVertex, type);
                    case "UDFunctionCall":
                        return ParseExpression($"{node.GetFunction()}({string.Join(", ", node.GetFunctionArguments().Select(a => TreeNodeToExpression(a, rootVertex, type)))})");
                    // Not implemented
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

        private ExpressionSyntax FunctionToExpression(string functionName, ParseTreeNode[] arguments, Vertex rootVertex,
            CellType type)
        {
            switch (functionName)
            {
                // arithmetic
                case "+":
                    // can be either string concat or addition
                    return GetBinaryExpression(functionName, arguments, rootVertex, type);
                case "-":
                case "*":
                case "/":
                    return GetBinaryExpression(functionName, arguments, rootVertex, CellType.Number);
                case "%":
                    if (arguments.Length != 1) return FunctionError(functionName, arguments);
                    // arguments[0] * 0.01
                    return BinaryExpression(
                        SyntaxKind.MultiplyExpression,
                        TreeNodeToExpression(arguments[0], rootVertex, CellType.Number),
                        LiteralExpression(SyntaxKind.NumericLiteralExpression, Literal(0.01)));
                case "^":
                    if (arguments.Length != 2) return FunctionError(functionName, arguments);
                    return ParseExpression($"Math.Pow({TreeNodeToExpression(arguments[0], rootVertex, CellType.Number)}, {TreeNodeToExpression(arguments[1], rootVertex, CellType.Number)})");
                case "ROUND":
                    if (arguments.Length != 2) return FunctionError(functionName, arguments);
                    return ParseExpression($"Math.Round({TreeNodeToExpression(arguments[0], rootVertex, CellType.Number)}, {TreeNodeToExpression(arguments[1], rootVertex, CellType.Number)}, MidpointRounding.AwayFromZero)");
                case "SUM":
                    return ParseExpression(string.Join(".Concat(", arguments.Select(a => TreeNodeToExpression(a, rootVertex, CellType.Number)))
                           + (arguments.Length > 1 ? ")" : "") + ".Sum()");

                // conditionals
                case "IF":
                    if (arguments.Length != 2 && arguments.Length != 3)
                        return FunctionError(functionName, arguments);
                    return ConditionalExpression(
                        ParenthesizedExpression(TreeNodeToExpression(arguments[0], rootVertex, type)),
                        ParenthesizedExpression(TreeNodeToExpression(arguments[1], rootVertex, type)),
                        arguments.Length == 3
                                    ? ParenthesizedExpression(TreeNodeToExpression(arguments[2], rootVertex, type))
                                    : CellTypeToNullExpression(rootVertex.CellType));

                case "=":
                    // cannot infer type from root vertex when entering a comparison
                    return GetBinaryExpression(functionName, arguments, rootVertex, CellType.Unknown);
                case "<>":
                case "<":
                case "<=":
                case ">=":
                case ">":
                    return GetBinaryExpression(functionName, arguments, rootVertex, CellType.Number);

                // strings
                case "&":
                case "CONCATENATE":
                    if (arguments.Length < 2) return FunctionError(functionName, arguments);
                    return ParseExpression(string.Join(" + ",
                        arguments.Select(a => TreeNodeToExpression(a, rootVertex, CellType.Text))));

                // TODO: ranges
                case ":":
                    if (arguments.Length != 2) return FunctionError(functionName, arguments);
                    return ParseExpression($"GetRange({TreeNodeToExpression(arguments[0], rootVertex, type)}, {TreeNodeToExpression(arguments[1], rootVertex, type)})");

                // other
                case "TODAY":
                    return ParseExpression("DateTime.Now");
                case "SECOND":
                case "MINUTE":
                case "HOUR":
                case "DAY":
                case "MONTH":
                case "YEAR":
                    if (arguments.Length != 1) return FunctionError(functionName, arguments);
                    return ParseExpression(TreeNodeToExpression(arguments[0], rootVertex, CellType.Date) + "." 
                        + functionName.Substring(0, 1) + functionName.Substring(1).ToLower());

                default:
                    return CommentExpression($"Function {functionName} not implemented yet! Args: " +
                        $"{string.Join("\n", arguments.Select(a => TreeNodeToExpression(a, rootVertex, type)))}", true);
            }
        }

        private readonly SortedList<string, (SyntaxKind syntaxKind, bool parenthesize)> operators 
            = new SortedList<string, (SyntaxKind syntaxKind, bool parenthesize)>
        {
            {"+", (SyntaxKind.AddExpression, false)},
            {"-", (SyntaxKind.SubtractExpression, false)},
            {"/", (SyntaxKind.DivideExpression, true)},
            {"*", (SyntaxKind.MultiplyExpression, true)},
            {"=", (SyntaxKind.EqualsExpression, false)},
            {"<>", (SyntaxKind.NotEqualsExpression, false)},
            {"<", (SyntaxKind.LessThanExpression, false)},
            {"<=", (SyntaxKind.LessThanOrEqualExpression, false)},
            {">=", (SyntaxKind.GreaterThanOrEqualExpression, false)},
            {">", (SyntaxKind.GreaterThanExpression, false)},

        };

        private ExpressionSyntax GetBinaryExpression(string functionName, ParseTreeNode[] arguments, Vertex vertex,
            CellType type)
        {
            if (arguments.Length != 2) return FunctionError(functionName, arguments);

            SyntaxKind syntaxKind = operators[functionName].syntaxKind;

            var leftExpression = TreeNodeToExpression(arguments[0], vertex, type);
            var rightExpression = TreeNodeToExpression(arguments[1], vertex, type);

            if (operators[functionName].parenthesize)
            {
                leftExpression = ParenthesizedExpression(leftExpression);
                rightExpression = ParenthesizedExpression(rightExpression);
            }

            return BinaryExpression( syntaxKind, leftExpression, rightExpression);
        }

        private ExpressionSyntax FormatVariableReference(string address, Vertex rootVertex, CellType type)
        {
            if (AddressToVertexDictionary.TryGetValue(address, out var vertex))
            {
                if (type != CellType.Unknown && vertex.CellType != type)
                {
                    vertex.CellType = type;
                    if (type == CellType.Text) vertex.Value = "";
                    if (type == CellType.Number) vertex.Value = 0;
                    if (type == CellType.Bool) vertex.Value = "false";
                }

                return rootVertex.Class == vertex.Class
                    ? (ExpressionSyntax)IdentifierName(vertex.VariableName)
                    : MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                    IdentifierName(vertex.Class.Name), IdentifierName(vertex.VariableName));
            }

            return CommentExpression($"{address} not found", true);
        }

        private ExpressionSyntax RuleNotImplemented(NonTerminal nt)
        {
            return CommentExpression($"Rule {nt.Rule} is not implemented for {nt.Name}!", true);
        }

        private ExpressionSyntax FunctionError(string functionName, ParseTreeNode[] arguments)
        {
            return CommentExpression($"Function {functionName} has incorrect number of arguments ({arguments.Length})", true);
        }


        private ExpressionSyntax CommentExpression(string comment, bool isError = false)
        {
            return IdentifierName(
                Identifier(
                    TriviaList(Comment("/* " + (isError ? "ERROR: " : "") + comment + " */")),
                    "",
                    TriviaList()));
        }

        private string GenerateUniqueName(string variableName, HashSet<string> usedVariableNames)
        {
            variableName = GenerateNonDuplicateName(usedVariableNames, variableName);
            usedVariableNames.Add(variableName);
            return variableName;
        }

        private HashSet<string> CreateUsedVariableHashSet()
        {
            return new[]
            {
                "bool", "byte", "sbyte", "short", "ushort", "int", "uint", "long", "ulong", "double", "float",
                "decimal", "string", "char", "void", "object", "typeof", "sizeof", "null", "true", "false", "if",
                "else", "while", "for", "foreach", "do", "switch", "case", "default", "lock", "try", "throw", "catch",
                "finally", "goto", "break", "continue", "return", "public", "private", "internal", "protected", "static",
                "readonly", "sealed", "const", "fixed", "stackalloc", "volatile", "new", "override", "abstract",
                "virtual", "event", "extern", "ref", "out", "in", "is", "as", "params", "__arglist", "__makeref",
                "__reftype", "__refvalue", "this", "base", "namespace", "using", "class", "struct", "interface",
                "enum", "delegate", "checked", "unchecked", "unsafe", "operator", "implicit", "explicit"
            }.ToHashSet();
        }

        private string GetTypeString(Vertex vertex)
        {
            switch (vertex.CellType)
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
                    return "dynamic";
                default:
                    return "";
            }
        }

        public static ExpressionSyntax CellTypeToNullExpression(CellType cellType)
        {
            switch (cellType)
            {
                case CellType.Bool:
                    return LiteralExpression(SyntaxKind.FalseLiteralExpression);
                case CellType.Unknown:
                case CellType.Date:
                    return LiteralExpression(SyntaxKind.NullLiteralExpression);
                case CellType.Number:
                    return LiteralExpression(SyntaxKind.NumericLiteralExpression, Literal(0));
                case CellType.Text:
                default:
                    return LiteralExpression(SyntaxKind.StringLiteralExpression, Literal(""));
            }
        }

        private ExpressionSyntax VertexValueToExpression(CellType cellType, dynamic vertexValue)
        {
            switch (cellType)
            {
                case CellType.Bool:
                    return LiteralExpression(vertexValue ? SyntaxKind.TrueLiteralExpression : SyntaxKind.FalseLiteralExpression);
                case CellType.Text:
                    return LiteralExpression(SyntaxKind.StringLiteralExpression, Literal(vertexValue));
                case CellType.Number:
                    return LiteralExpression(SyntaxKind.NumericLiteralExpression, Literal(vertexValue));
                case CellType.Date:
                    return ParseExpression($"DateTime.Parse(\"{vertexValue}\")");
                case CellType.Unknown:
                default:
                    return LiteralExpression(SyntaxKind.NullLiteralExpression);
            }
        }
    }
}
