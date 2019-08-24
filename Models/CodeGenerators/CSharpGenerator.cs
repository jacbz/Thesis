using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Irony.Parsing;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Formatting;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Thesis.ViewModels;
using static Microsoft.CodeAnalysis.CSharp.SyntaxFactory;
using XLParser;
using Formatter = Microsoft.CodeAnalysis.Formatting.Formatter;

namespace Thesis.Models.CodeGenerators
{
    public class CSharpGenerator : CodeGenerator
    {
        private HashSet<string> _usedVariableNames;

        // vertices in this list must have type dynamic
        private HashSet<Vertex> _useDynamic;

        public CSharpGenerator(List<GeneratedClass> generatedClasses, Dictionary<string, Vertex> addressToVertexDictionary)
            : base(generatedClasses, addressToVertexDictionary)
        {
        }

        public override async Task<string> GenerateCodeAsync(Dictionary<string, TestResult> testResults = null)
        {
            return await Task.Run(() => GenerateCode(testResults));
        }

        public string GenerateCode(Dictionary<string, TestResult> testResults = null)
        {
            Tester = new CSharpTester();
            _usedVariableNames = new HashSet<string>();
            _useDynamic = new HashSet<Vertex>();

            // generate variable names for all
            GenerateVariableNamesForAll();
            VariableNameToVertexDictionary = GeneratedClasses.SelectMany(c => c.Vertices).ToDictionary(v => v.VariableName);

            // namespace Thesis
            var @namespace = NamespaceDeclaration(ParseName("Thesis")).NormalizeWhitespace();
            // using System;
            @namespace = @namespace.AddUsings(
                UsingDirective(ParseName("System")),
                UsingDirective(ParseName("System.Linq")));

            // public class ThesisResult (Main class)
            var resultClass = GenerateResultClass(testResults);
            @namespace = @namespace.AddMembers(resultClass);

            var normalClasses = new List<MemberDeclarationSyntax>();
            var sharedClasses = new List<MemberDeclarationSyntax>();

            // shared classes first (must determine which types are dynamic)
            foreach (var generatedClass in GeneratedClasses.OrderBy(v => !v.IsSharedClass))
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
                var expression = TreeNodeToExpression(formula.ParseTree, formula);

                if (generatedClass.IsSharedClass)
                {
                    // for shared classes: omit type name as already declared

                    // test result comment
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
                    // test result comment
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

        private ExpressionSyntax TreeNodeToExpression(ParseTreeNode node, Vertex rootVertex)
        {
            // Non-Terminals
            if (node.Term is NonTerminal nt)
            {
                switch (node.Term.Name)
                {
                    case "Cell":
                        return FormatVariableReferenceFromAddress(node.FindTokenAndGetText(), rootVertex);
                    case "Constant":
                        string constant = node.FindTokenAndGetText();
                        if (constant.Contains("%") && !constant.Any(char.IsLetter))
                        {
                            // e.g. IF(A1>0, "1%", "2%")  (user entered "1%" instead of 1%)
                            rootVertex.CellType = CellType.Number;
                            constant = constant.Replace("%", "*0.01").Replace("\"", "");
                        }
                        return ParseExpression(constant);
                    case "FormulaWithEq":
                        return TreeNodeToExpression(node.ChildNodes[1], rootVertex);
                    case "Formula":
                        // for rule OpenParen + Formula + CloseParen
                        return TreeNodeToExpression(node.ChildNodes.Count == 3 ? node.ChildNodes[1] : node.ChildNodes[0], rootVertex);
                    case "FunctionCall":
                        return FunctionToExpression(node.GetFunction(), node.GetFunctionArguments().ToArray(), rootVertex);
                    case "Reference":
                        if (node.ChildNodes.Count == 1)
                            return TreeNodeToExpression(node.ChildNodes[0], rootVertex);
                        if (node.ChildNodes.Count == 2)
                        {
                            var prefix = node.ChildNodes[0];
                            var refName = prefix.ChildNodes.Count == 2 ? prefix.ChildNodes[1].FindTokenAndGetText() : prefix.FindTokenAndGetText();
                            return ParseExpression($"ExternalRef(\"[{refName}\", {TreeNodeToExpression(node.ChildNodes[1], rootVertex)})");
                        }
                        return RuleNotImplemented(nt);
                    case "ReferenceItem":
                        return node.ChildNodes.Count == 1 ? TreeNodeToExpression(node.ChildNodes[0], rootVertex) : RuleNotImplemented(nt);
                    case "ReferenceFunctionCall":
                        return FunctionToExpression(node.GetFunction(), node.GetFunctionArguments().ToArray(), rootVertex);
                    case "UDFunctionCall":
                        return ParseExpression($"{node.GetFunction()}({string.Join(", ", node.GetFunctionArguments().Select(a => TreeNodeToExpression(a, rootVertex)))})");
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

        private ExpressionSyntax FunctionToExpression(string functionName, ParseTreeNode[] arguments, Vertex currentVertex)
        {
            switch (functionName)
            {
                // arithmetic
                case "+":
                case "-":
                case "*":
                case "/":
                    return GenerateBinaryExpression(functionName, arguments, currentVertex);
                case "%":
                    if (arguments.Length != 1) return FunctionError(functionName, arguments);
                    // arguments[0] * 0.01
                    return BinaryExpression(
                        SyntaxKind.MultiplyExpression,
                        TreeNodeToExpression(arguments[0], currentVertex),
                        LiteralExpression(SyntaxKind.NumericLiteralExpression, Literal(0.01)));
                case "^":
                    if (arguments.Length != 2) return FunctionError(functionName, arguments);
                    return ParseExpression($"Math.Pow({TreeNodeToExpression(arguments[0], currentVertex)}, {TreeNodeToExpression(arguments[1], currentVertex)})");
                case "ROUND":
                    if (arguments.Length != 2) return FunctionError(functionName, arguments);
                    return ParseExpression($"Math.Round({TreeNodeToExpression(arguments[0], currentVertex)}, {TreeNodeToExpression(arguments[1], currentVertex)}, MidpointRounding.AwayFromZero)");
                case "SUM":
                    // Collection(...).Sum()
                    return InvocationExpression(MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                        CollectionOf(arguments.Select(a => TreeNodeToExpression(a, currentVertex)).ToArray()),
                            IdentifierName("Sum")));
                case "MIN":
                    // Collection(...).Min()
                    return InvocationExpression(MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                        CollectionOf(arguments.Select(a => TreeNodeToExpression(a, currentVertex)).ToArray()),
                        IdentifierName("Min")));
                case "MAX":
                    // Collection(...).Max()
                    return InvocationExpression(MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                        CollectionOf(arguments.Select(a => TreeNodeToExpression(a, currentVertex)).ToArray()),
                        IdentifierName("Max")));

                // logicial functions
                case "IF":
                    if (arguments.Length != 2 && arguments.Length != 3)
                        return FunctionError(functionName, arguments);

                    ExpressionSyntax condition = TreeNodeToExpression(arguments[0], currentVertex);
                    ExpressionSyntax whenTrue = TreeNodeToExpression(arguments[1], currentVertex);
                    ExpressionSyntax whenFalse;

                    // if the condition is not always a bool (e.g. dynamic), use Equals(cond, true)
                    // otherwise we might have a number as cond, and number can not be evaluated as bool
                    var conditionType = GetType(arguments[0]);
                    if (!conditionType.HasValue || conditionType.Value != CellType.Bool)
                    {
                        condition = InvocationExpression(IdentifierName("Equals"))
                            .AddArgumentListArguments(
                                Argument(condition),
                                Argument(LiteralExpression(SyntaxKind.TrueLiteralExpression)));
                    }


                    // check if there is a mismatch between type of whenTrue and whenFalse
                    bool argumentsHaveDifferentTypes;
                    if (arguments.Length == 3)
                    {
                        whenFalse = TreeNodeToExpression(arguments[2], currentVertex);
                        argumentsHaveDifferentTypes = IsSameType(arguments[1], arguments[2]) == null;
                    }
                    else
                    {
                        // if no else statement is given, Excel defaults to FALSE
                        whenFalse = LiteralExpression(SyntaxKind.FalseLiteralExpression);
                        argumentsHaveDifferentTypes = !IsTypeBoolean(arguments[1]);
                    }

                    // if there is a mismatch in argument types, the variable must be of type dynamic
                    if (argumentsHaveDifferentTypes)
                    {
                        _useDynamic.Add(currentVertex);
                        // we must cast to type (dynamic):
                        // https://stackoverflow.com/questions/57633328/change-a-dynamics-type-in-a-ternary-conditional-statement#57633386
                        whenFalse = CastExpression(IdentifierName("dynamic"), whenFalse);
                    }

                    return ParenthesizedExpression(ConditionalExpression(condition, whenTrue, whenFalse));
                case "AND":
                case "OR":
                case "XOR":
                    if (arguments.Length == 0) return FunctionError(functionName, arguments);
                    if (arguments.Length == 1) return TreeNodeToExpression(arguments[0], currentVertex);
                    return ParenthesizedExpression(FoldBinaryExpression(functionName, arguments, currentVertex));
                case "NOT":
                    if (arguments.Length != 1) return FunctionError(functionName, arguments);
                    return PrefixUnaryExpression(SyntaxKind.LogicalNotExpression, 
                        TreeNodeToExpression(arguments[0], currentVertex));
                case "TRUE":
                    return LiteralExpression(SyntaxKind.TrueLiteralExpression);
                case "FALSE":
                    return LiteralExpression(SyntaxKind.FalseLiteralExpression);

                case "=":
                    if (arguments.Length != 2) return FunctionError(functionName, arguments);
                    var leftExpression = TreeNodeToExpression(arguments[0], currentVertex);
                    var rightExpression = TreeNodeToExpression(arguments[1], currentVertex);

                    ExpressionSyntax equalsExpression;
                    var leftType = GetType(arguments[0]);
                    var rightType = GetType(arguments[1]);
                    if (leftType.HasValue && rightType.HasValue && leftType.Value == rightType.Value)
                    {
                        // Excel uses case insensitive string compare
                        if (leftType.Value == CellType.Text && rightType.Value == CellType.Text)
                        {
                            equalsExpression = InvocationExpression(
                                    MemberAccessExpression(
                                        SyntaxKind.SimpleMemberAccessExpression,
                                        ParenthesizedExpression(leftExpression),
                                        IdentifierName("CIEquals")))
                                .AddArgumentListArguments(Argument(rightExpression));
                        }
                        else
                        {
                            // use normal == compare for legibility
                            equalsExpression = BinaryExpression(SyntaxKind.EqualsExpression,
                                leftExpression, rightExpression);
                        }
                    }
                    else
                    {
                        // if types different, use our custom Equals method (== would throw exception if types are different)
                        equalsExpression = InvocationExpression(IdentifierName("Equals"))
                            .AddArgumentListArguments(Argument(leftExpression), Argument(rightExpression));
                    }
                    return equalsExpression;
                case "<>":
                case "<":
                case "<=":
                case ">=":
                case ">":
                    return GenerateBinaryExpression(functionName, arguments, currentVertex);

                // strings
                case "&":
                case "CONCATENATE":
                    if (arguments.Length < 2) return FunctionError(functionName, arguments);
                    return ParseExpression(string.Join(" + ",
                        arguments.Select(a => TreeNodeToExpression(a, currentVertex))));

                case ":":
                    if (arguments.Length != 2) return FunctionError(functionName, arguments);
                    return GetRangeExpression(arguments[0], arguments[1], currentVertex);

                // other
                case "DATE":
                    if (arguments.Length != 3) return FunctionError(functionName, arguments);
                    return ObjectCreationExpression(
                            IdentifierName("DateTime"))
                        .AddArgumentListArguments(
                            Argument(TreeNodeToExpression(arguments[0], currentVertex)),
                            Argument(TreeNodeToExpression(arguments[1], currentVertex)),
                            Argument(TreeNodeToExpression(arguments[2], currentVertex))
                            );
                case "TODAY":
                    // DateTime.Now
                    return MemberAccessExpression(
                        SyntaxKind.SimpleMemberAccessExpression,
                        IdentifierName("DateTime"),
                        IdentifierName("Now"));
                case "SECOND":
                case "MINUTE":
                case "HOUR":
                case "DAY":
                case "MONTH":
                case "YEAR":
                    if (arguments.Length != 1) return FunctionError(functionName, arguments);
                    return ParseExpression(TreeNodeToExpression(arguments[0], currentVertex) + "."
                        + functionName.Substring(0, 1) + functionName.Substring(1).ToLower());

                default:
                    return CommentExpression($"Function {functionName} not implemented yet! Args: " +
                        $"{string.Join("\n", arguments.Select(a => TreeNodeToExpression(a, currentVertex)))}", true);
            }
        }


        private readonly Dictionary<string, CellType> _functionToCellTypeDictionary = new Dictionary<string, CellType>()
        {
            { "+", CellType.Unknown },
            { "-", CellType.Number },
            { "*", CellType.Number },
            { "/", CellType.Number },
            { "%", CellType.Number },
            { "^", CellType.Number },
            { "ROUND", CellType.Number },
            { "SUM", CellType.Number },
            { "MIN", CellType.Number },
            { "MAX", CellType.Number },

            { "IF", CellType.Unknown },
            { "AND", CellType.Bool },
            { "NOT", CellType.Bool },
            { "OR", CellType.Bool },
            { "XOR", CellType.Bool },
            { "TRUE", CellType.Bool },
            { "FALSE", CellType.Bool },

            { "=", CellType.Bool },
            { "<>", CellType.Number },
            { "<", CellType.Number },
            { "<=", CellType.Number },
            { ">=", CellType.Number },
            { ">", CellType.Number },

            { "&", CellType.Text },
            { "CONCATENATE", CellType.Text },
            { ":", CellType.Unknown },

            { "DATE", CellType.Date },
            { "SECOND", CellType.Date },
            { "MINUTE", CellType.Date },
            { "HOUR", CellType.Date },
            { "DAY", CellType.Date },
            { "MONTH", CellType.Date },
            { "TODAY", CellType.Date },
        };

        // gets the type of a node
        // if multiple types are found, or type is unknown or dynamic, return null
        private CellType? GetType(ParseTreeNode node)
        {
            if (node.Term.Name == "ReferenceFunctionCall")
            {
                var arguments = node.ChildNodes[1];
                if (arguments.ChildNodes.Count == 3)
                    return IsSameType(arguments.ChildNodes[1], arguments.ChildNodes[2]);
                if (arguments.ChildNodes.Count == 2)
                    return IsTypeBoolean(arguments.ChildNodes[1]) ? CellType.Bool : (CellType?)null;
                return null;

            }
            if (node.Term.Name == "FunctionCall")
            {
                var function = node.GetFunction();
                if (_functionToCellTypeDictionary.TryGetValue(function, out var functionType)
                    && functionType != CellType.Unknown)
                    return functionType;
            }

            // TODO Implement
            if (node.Term.Name == "Prefix")
            {
                return null;
            }
            // TODO Implement
            if (node.Term.Name == "UDFunctionCall")
            {
                return null;
            }

            var text = node.FindTokenAndGetText();
            if (node.Term.Name == "Cell")
            {
                if (!AddressToVertexDictionary.TryGetValue(text, out var vertex))
                    return null;
                return _useDynamic.Contains(vertex) ? (CellType?)null : vertex.CellType;
            }

            if (node.Term.Name == "Constant")
            {
                if (text.Contains("%")) return CellType.Number;
                if (text.ToLower().Contains("true") || text.ToLower().Contains("false")) return CellType.Bool;
                if (!text.Contains("\"")) return CellType.Number;
                return CellType.Text;
            }

            var childTypes = node.ChildNodes.Select(GetType).Distinct().ToList();
            if (childTypes.All(c => c.HasValue) && childTypes.Count == 1 && childTypes[0].Value != CellType.Unknown)
                return childTypes[0].Value;
            return null;
        }

        // checks if two parse tree nodes have the same type
        private CellType? IsSameType(ParseTreeNode a, ParseTreeNode b)
        {
            var typeA = GetType(a);
            var typeB = GetType(b);
            if (typeA.HasValue && typeA.Value != CellType.Unknown && typeB.HasValue && typeB.Value != CellType.Unknown
                && typeA.Value == typeB.Value)
            {
                return typeA.Value;
            }
            return null;
        }

        private bool IsTypeBoolean(ParseTreeNode a)
        {
            var typeOfIf = GetType(a);
            return typeOfIf.HasValue && typeOfIf.Value == CellType.Bool;
        }

        // A1:B2 -> new[]{A1,A2,B1,B2}
        private ExpressionSyntax GetRangeExpression(ParseTreeNode left, ParseTreeNode right, Vertex currentVertex)
        {
            if (left.ChildNodes.Count != 1 || right.ChildNodes.Count != 1)
                return CommentExpression($"Error while parsing range");


            var leftChild = left.ChildNodes[0];
            var rightChild = right.ChildNodes[0];
            if (leftChild.Term.Name != "Cell" || rightChild.Term.Name != "Cell")
            {
                return CommentExpression("Range with more than 2 components not implemented yet!");
            }

            var leftAddress = leftChild.FindTokenAndGetText();
            var rightAddress = rightChild.FindTokenAndGetText();

            var entries = new List<ExpressionSyntax>();
            foreach (var address in Utility.AddressesInRange(leftAddress, rightAddress))
            {
                if (AddressToVertexDictionary.TryGetValue(address, out var vertex))
                {
                    entries.Add(FormatVariableReference(vertex, currentVertex));
                }
            }

            // remove last comma
            if (entries.Count > 0) entries.RemoveAt(entries.Count - 1);

            // create custom collection
            return CollectionOf(entries.ToArray());
        }

        private ExpressionSyntax CollectionOf(params ExpressionSyntax[] expressions)
        {
            // avoid Collection(Collection)
            if (expressions.Length == 1 && expressions[0] is InvocationExpressionSyntax inv
                && ((IdentifierNameSyntax)inv.Expression).Identifier.Text == "Collection")
                return expressions[0];

            return InvocationExpression(
                    IdentifierName("Collection"))
                .AddArgumentListArguments(expressions.Select(Argument).ToArray());
        }

        private readonly SortedList<string, (SyntaxKind syntaxKind, bool parenthesize)> _binaryOperators
            = new SortedList<string, (SyntaxKind syntaxKind, bool parenthesize)>
        {
            {"+", (SyntaxKind.AddExpression, false)},
            {"-", (SyntaxKind.SubtractExpression, false)},
            {"/", (SyntaxKind.DivideExpression, true)},
            {"*", (SyntaxKind.MultiplyExpression, true)},
            {"<>", (SyntaxKind.NotEqualsExpression, false)},
            {"<", (SyntaxKind.LessThanExpression, false)},
            {"<=", (SyntaxKind.LessThanOrEqualExpression, false)},
            {">=", (SyntaxKind.GreaterThanOrEqualExpression, false)},
            {">", (SyntaxKind.GreaterThanExpression, false)},
            {"AND", (SyntaxKind.LogicalAndExpression, true)},
            {"OR", (SyntaxKind.LogicalOrExpression, true)},
            {"XOR", (SyntaxKind.ExclusiveOrExpression, true)},
        };


        private ExpressionSyntax GenerateBinaryExpression(string functionName, ParseTreeNode[] arguments, Vertex vertex)
        {
            if (arguments.Length != 2) return FunctionError(functionName, arguments);
            return GenerateBinaryExpression(functionName, 
                TreeNodeToExpression(arguments[0], vertex), 
                TreeNodeToExpression(arguments[1], vertex));
        }

        private ExpressionSyntax GenerateBinaryExpression(string functionName,
            ExpressionSyntax leftExpression, ExpressionSyntax rightExpression)
        {
            SyntaxKind syntaxKind = _binaryOperators[functionName].syntaxKind;
            if (_binaryOperators[functionName].parenthesize)
            {
                leftExpression = ParenthesizedExpression(leftExpression);
                rightExpression = ParenthesizedExpression(rightExpression);
            }

            return BinaryExpression(syntaxKind, leftExpression, rightExpression);
        }

        // e.g. function &&, arg [a,b,c,d] => a && (b && (c && d))
        private ExpressionSyntax FoldBinaryExpression(string functionName, ParseTreeNode[] arguments, Vertex vertex)
        {
            var syntaxKind = _binaryOperators[functionName].syntaxKind;
            // do not parenthesize
            return arguments.Select(a => TreeNodeToExpression(a, vertex))
                .Aggregate((acc, right) => BinaryExpression(syntaxKind, acc, right));
        }

        private ExpressionSyntax FormatVariableReferenceFromAddress(string address, Vertex currentVertex)
        {
            if (AddressToVertexDictionary.TryGetValue(address, out var variableVertex))
            {
                return FormatVariableReference(variableVertex, currentVertex);
            }

            return CommentExpression($"{address} not found", true);
        }

        private static ExpressionSyntax FormatVariableReference(Vertex variableVertex, Vertex currentVertex)
        {
            return currentVertex.Class == variableVertex.Class
                ? (ExpressionSyntax) IdentifierName(variableVertex.VariableName)
                : MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                    IdentifierName(variableVertex.Class.Name), IdentifierName(variableVertex.VariableName));
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
            if (_useDynamic.Contains(vertex)) return "dynamic";

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
                    return "EmptyCell";
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
                    return ObjectCreationExpression(IdentifierName("EmptyCell")).WithArgumentList(ArgumentList());
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
                    return ObjectCreationExpression(IdentifierName("EmptyCell")).WithArgumentList(ArgumentList());
                default:
                    return LiteralExpression(SyntaxKind.NullLiteralExpression);
            }
        }
    }
}
