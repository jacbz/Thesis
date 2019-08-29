using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Irony.Parsing;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Thesis.Models.VertexTypes;
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

        public CSharpGenerator(ClassCollection classCollection, 
            Dictionary<(string worksheet, string address), CellVertex> addressToVertexDictionary, 
            Dictionary<string, RangeVertex> rangeDictionary, 
            Dictionary<string, Vertex> nameDictionary) : 
            base(classCollection, addressToVertexDictionary, rangeDictionary, nameDictionary)
        {
        }

        public override async Task<Code> GenerateCodeAsync(Dictionary<string, TestResult> testResults = null)
        {
            return await Task.Run(() => GenerateCode(testResults));
        }

        public Code GenerateCode(Dictionary<string, TestResult> testResults = null)
        {
            _usedVariableNames = new HashSet<string>();
            _useDynamic = new HashSet<Vertex>();

            // generate variable names for all
            GenerateVariableNamesForAll();

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
            var staticClasses = new List<MemberDeclarationSyntax>();
            var classesCode = new List<ClassCode>();

            // static classes first (must determine which types are dynamic)
            foreach (var generatedClass in ClassCollection.Classes.OrderBy(v => !v.IsStaticClass))
            {
                var (newClass, classCode) = GenerateClass(generatedClass, testResults);
                classesCode.Add(classCode);

                if (generatedClass.IsStaticClass)
                    staticClasses.Add(newClass);
                else
                    normalClasses.Add(newClass);
            }

            // show static classes on top for aesthetic reasons
            @namespace = @namespace
                .AddMembers(staticClasses.ToArray())
                .AddMembers(normalClasses.ToArray());

            // format code and return as string
            var workspace = new AdhocWorkspace();
            var options = workspace.Options;
            var node = Formatter.Format(@namespace, workspace, options);
            var sourceCode = node.ToFullString();

            var variableNameToVertexDictionary = ClassCollection.Classes.SelectMany(c => c.Vertices).ToDictionary(v => v.VariableName);
            return new Code(sourceCode, variableNameToVertexDictionary, new CSharpTester(classesCode));
        }

        private (ClassDeclarationSyntax classDeclarationSyntax, ClassCode classCode)
            GenerateClass(Class @class, Dictionary<string, TestResult> testResults = null)
        {
            // public class {generatedClass.Name}
            var newClass = ClassDeclaration(@class.Name)
                .AddModifiers(Token(SyntaxKind.PublicKeyword));
            // static classes are static
            if (@class.IsStaticClass)
                newClass = newClass.AddModifiers(Token(SyntaxKind.StaticKeyword));

            // split vertex list into two
            var lookup = @class.Vertices
                .ToLookup(v => v.IsExternal || v is CellVertex c && c.NodeType == NodeType.Constant);
            var constants = lookup[true].ToList();
            var formulas = lookup[false].ToList();

            // generate assignments
            var statements = new List<StatementSyntax>();
            foreach (var formula in formulas)
            {
                var expression = formula is CellVertex cellVertex 
                    ? TreeNodeToExpression(cellVertex.ParseTree, cellVertex)
                    : RangeVertexToExpression(formula as RangeVertex);

                if (@class.IsStaticClass)
                {
                    // for static classes: omit type name as already declared

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
                var expression = VertexValueToExpression(constant);
                var field =
                    FieldDeclaration(
                        VariableDeclaration(
                                ParseTypeName(GetTypeString(constant)))
                            .AddVariables(VariableDeclarator(constant.VariableName)
                                .WithInitializer(
                                    EqualsValueClause(expression))));
                if (@class.IsStaticClass)
                    field = field.AddModifiers(
                        Token(SyntaxKind.PublicKeyword),
                        Token(SyntaxKind.StaticKeyword));
                else
                    field = field.AddModifiers(Token(SyntaxKind.PrivateKeyword));

                fields.Add(field);
            }

            // extra fields for static classes
            if (@class.IsStaticClass)
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

            newClass = newClass.AddMembers(fields.ToArray());

            if (statements.Count > 0)
            {
                // generate Calculate/Init method
                var method = GenerateMethod(@class, statements);

                // add fields and method to class
                newClass = newClass.AddMembers(method);
            }

            var classCode = new ClassCode(
                @class.IsStaticClass,
                @class.Name,
                newClass.NormalizeWhitespace().ToFullString(),
                string.Join("\n", fields.Select(f => f.NormalizeWhitespace().ToFullString())),
                string.Join("\n", statements.Select(f => f.NormalizeWhitespace().ToFullString())));
            return (newClass, classCode);
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

        private MemberDeclarationSyntax GenerateMethod(Class @class, List<StatementSyntax> statements)
        {
            if (@class.IsStaticClass)
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
                statements.Add(ReturnStatement(IdentifierName(@class.OutputVertex.VariableName)));

                // public {type} Calculate()
                var outputField = @class.OutputVertex;
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
            foreach (var generatedClass in ClassCollection.Classes
                .Where(c => c.Vertices.Count(v => !v.IsExternal && 
                                                  v is CellVertex cellVertex && cellVertex.NodeType != NodeType.Constant) > 0)
                .OrderBy(v => !v.IsStaticClass))
            {
                if (generatedClass.IsStaticClass)
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
            foreach (var generatedClass in ClassCollection.Classes)
            {
                generatedClass.Name = GenerateUniqueName(generatedClass.Name, _usedVariableNames);
                foreach (var vertex in generatedClass.Vertices)
                {
                    vertex.VariableName = GenerateUniqueName(vertex.VariableName.MakeNameVariableConform(), _usedVariableNames);
                }
            }
        }

        private ExpressionSyntax RangeVertexToExpression(RangeVertex rangeVertex)
        {
            switch (rangeVertex.Type)
            {
                case RangeVertex.RangeType.Empty:
                {
                    return MatrixOf();
                }
                case RangeVertex.RangeType.Single:
                {
                    return CellVertexToConstantOrVariable(rangeVertex.GetSingleElement(), rangeVertex);
                }
                case RangeVertex.RangeType.Column:
                {
                    var columnArray = rangeVertex.GetColumnArray();
                    return ColumnOf(columnArray
                        .Select(cellVertex => CellVertexToConstantOrVariable(cellVertex, rangeVertex))
                        .ToArray());
                }
                case RangeVertex.RangeType.Row:
                {
                    var rowArray = rangeVertex.GetRowArray();
                    return RowOf(rowArray
                        .Select(cellVertex => CellVertexToConstantOrVariable(cellVertex, rangeVertex))
                        .ToArray());
                }
                case RangeVertex.RangeType.Matrix:
                {
                    var matrix = rangeVertex.GetMatrixArray();
                    return MatrixOf(matrix
                        .Select(rowArray => RowOf(rowArray.Select(cellVertex => CellVertexToConstantOrVariable(cellVertex, rangeVertex))
                            .ToArray()))
                        .ToArray());
                }
                default:
                    return null;
            }
        }

        // if a cell vertex already exists, use that one, else format as constant
        private ExpressionSyntax CellVertexToConstantOrVariable(CellVertex cellVertex, Vertex currentVertex)
        {
            return AddressToVertexDictionary.TryGetValue((currentVertex.WorksheetName, cellVertex.StringAddress), out var c) 
                ? VariableReferenceToExpression(c, currentVertex) 
                : VertexValueToExpression(cellVertex);
        }

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
                        if (node.ChildNodes[0].Term.Name == "Bool") return ParseExpression(constant.ToLower());
                        return ParseExpression(constant);
                    case "FormulaWithEq":
                        return TreeNodeToExpression(node.ChildNodes[1], currentVertex);
                    case "Formula":
                        // for rule OpenParen + Formula + CloseParen
                        return TreeNodeToExpression(node.ChildNodes.Count == 3 ? node.ChildNodes[1] : node.ChildNodes[0], currentVertex);
                    case "FunctionCall":
                        return FunctionToExpression(node.GetFunction(), node.GetFunctionArguments().ToArray(), currentVertex);
                    case "NamedRange":
                        return NamedRangeToExpression(node, currentVertex);
                    case "Reference":
                        if (node.ChildNodes.Count == 1)
                            return TreeNodeToExpression(node.ChildNodes[0], currentVertex);
                        // External cells
                        if (node.ChildNodes.Count == 2)
                        {
                            var prefix = node.ChildNodes[0];
                            var sheetName = prefix.ChildNodes.Count == 2 ? prefix.ChildNodes[1].FindTokenAndGetText() : prefix.FindTokenAndGetText();
                            sheetName = sheetName.FormatSheetName();
                            var address = node.ChildNodes[1].FindTokenAndGetText();
                            return MemberAccessExpression(
                                SyntaxKind.SimpleMemberAccessExpression,
                                IdentifierName("External"),
                                IdentifierName(Vertex.GenerateExternalVariableName(sheetName, address)));
                        }
                        return RuleNotImplemented(nt);
                    case "ReferenceItem":
                        return node.ChildNodes.Count == 1 ? TreeNodeToExpression(node.ChildNodes[0], currentVertex) : RuleNotImplemented(nt);
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

        private ExpressionSyntax FunctionToExpression(string functionName, ParseTreeNode[] arguments, CellVertex currentVertex)
        {
            switch (functionName)
            {
                // arithmetic
                case "+":
                case "-":
                {
                    if (arguments.Length == 1)
                        return PrefixUnaryExpression( SyntaxKind.UnaryMinusExpression,
                            TreeNodeToExpression(arguments[0], currentVertex));

                    if (arguments.Length != 2) return FunctionError(functionName, arguments);

                    // adding a Date and a number yields Date.AddDays(number)
                    var typeLeft = GetType(arguments[0], currentVertex);
                    var typeRight = GetType(arguments[1], currentVertex);
                    if (typeLeft.HasValue && typeRight.HasValue)
                    {
                        var leftExpr = TreeNodeToExpression(arguments[0], currentVertex);
                        var rightExpr = TreeNodeToExpression(arguments[1], currentVertex);
                        if (typeLeft.Value == CellType.Date)
                        {
                            return InvocationExpression(MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                                    leftExpr, IdentifierName("AddDays")))
                                .AddArgumentListArguments(Argument(rightExpr));
                        }
                        if (typeRight.Value == CellType.Date)
                        {
                            return InvocationExpression(MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                                    rightExpr, IdentifierName("AddDays")))
                                .AddArgumentListArguments(Argument(leftExpr));
                        }
                    }

                    return GenerateBinaryExpression(functionName, arguments, currentVertex);
                }

                case "*":
                case "/":
                {
                    return GenerateBinaryExpression(functionName, arguments, currentVertex);
                }
                case "%":
                {
                    if (arguments.Length != 1) return FunctionError(functionName, arguments);
                    // arguments[0] * 0.01
                    return BinaryExpression(
                        SyntaxKind.MultiplyExpression,
                        TreeNodeToExpression(arguments[0], currentVertex),
                        LiteralExpression(SyntaxKind.NumericLiteralExpression, Literal(0.01)));
                }
                case "^":
                {
                    if (arguments.Length != 2) return FunctionError(functionName, arguments);
                    return ClassFunctionCall("Math", "Pow", 
                        TreeNodeToExpression(arguments[0], currentVertex), TreeNodeToExpression(arguments[1], currentVertex));
                }
                case "ROUND":
                {
                    return RoundFunction("Round", arguments, currentVertex);
                }
                case "ROUNDUP":
                {
                    return RoundFunction("RoundUp", arguments, currentVertex);
                }
                case "ROUNDDOWN":
                {
                    return RoundFunction("RoundDown", arguments, currentVertex);
                }
                case "SUM":
                case "MIN":
                case "MAX":
                case "COUNT":
                case "AVERAGE":
                {
                    // Collection(...).Sum/Min/Max/Count/Average()
                    return InvocationExpression(MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                        CollectionOf(arguments.Select(a => TreeNodeToExpression(a, currentVertex)).ToArray()),
                        IdentifierName(functionName.ToTitleCase())));
                }

                // reference functions
                case "VLOOKUP":
                case "HLOOKUP":
                    {
                    if (arguments.Length != 3 && arguments.Length != 4)
                        return FunctionError(functionName, arguments);

                    var matrix = RangeOrNamedRangeToExpression(arguments[1], currentVertex);

                    var lookupValue = TreeNodeToExpression(arguments[0], currentVertex);
                    var columnIndex = TreeNodeToExpression(arguments[2], currentVertex);

                    var function = functionName == "VLOOKUP" ? "VLookUp" : "HLookUp";
                    var expression = InvocationExpression(
                            MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                                matrix,
                                IdentifierName(function)))
                        .AddArgumentListArguments(
                            Argument(lookupValue),
                            Argument(columnIndex));

                    if (arguments.Length == 4)
                        expression = expression.AddArgumentListArguments(
                            Argument(TreeNodeToExpression(arguments[3], currentVertex)));
                    return expression;
                }
                case "CHOOSE":
                {
                    if (arguments.Length < 2)
                        return FunctionError(functionName, arguments);
                    var collection = CollectionOf(arguments
                        .Skip(1)
                        .Select(a => TreeNodeToExpression(a, currentVertex))
                        .ToArray());

                    // collection[arguments[0]]
                    return ElementAccessExpression(collection)
                        .WithArgumentList(BracketedArgumentList(SingletonSeparatedList(
                                    Argument(TreeNodeToExpression(arguments[0], currentVertex)))));
                }
                case "INDEX":
                {
                    if (arguments.Length != 2 && arguments.Length != 3)
                        return FunctionError(functionName, arguments);
                    var matrixOrCollection = RangeOrNamedRangeToExpression(arguments[0], currentVertex);

                    var argumentList = new List<SyntaxNodeOrToken>();
                    argumentList.Add(Argument(TreeNodeToExpression(arguments[1], currentVertex)));
                    if (arguments.Length == 3)
                    {
                        argumentList.Add(Token(SyntaxKind.CommaToken));
                        argumentList.Add(Argument(TreeNodeToExpression(arguments[2], currentVertex)));
                    }

                    return ElementAccessExpression(matrixOrCollection)
                        .WithArgumentList(BracketedArgumentList(SeparatedList<ArgumentSyntax>(argumentList.ToArray())));
                }
                case "MATCH":
                {
                    if (arguments.Length != 2 && arguments.Length != 3)
                        return FunctionError(functionName, arguments);
                    var collection = RangeOrNamedRangeToExpression(arguments[1], currentVertex);

                    var expression = InvocationExpression(
                            MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                                collection,
                                IdentifierName("Match")))
                        .AddArgumentListArguments(
                            Argument(TreeNodeToExpression(arguments[0], currentVertex)));
                    if (arguments.Length == 3)
                        expression = expression.AddArgumentListArguments(
                            Argument(TreeNodeToExpression(arguments[2], currentVertex)));

                    return expression;
                }

                // logical functions
                case "IF":
                {
                    if (arguments.Length != 2 && arguments.Length != 3)
                        return FunctionError(functionName, arguments);

                    ExpressionSyntax condition = TreeNodeToExpression(arguments[0], currentVertex);
                    ExpressionSyntax whenTrue = TreeNodeToExpression(arguments[1], currentVertex);
                    ExpressionSyntax whenFalse;

                    // if the condition is not always a bool (e.g. dynamic), use Compare(cond, true)
                    // otherwise we might have a number as cond, and number can not be evaluated as bool
                    var conditionType = GetType(arguments[0], currentVertex);
                    if (!conditionType.HasValue || conditionType.Value != CellType.Bool)
                    {
                        condition = InvocationExpression(IdentifierName("Compare"))
                            .AddArgumentListArguments(
                                Argument(condition),
                                Argument(LiteralExpression(SyntaxKind.TrueLiteralExpression)));
                    }


                    // check if there is a mismatch between type of whenTrue and whenFalse
                    bool argumentsHaveDifferentTypes;
                    if (arguments.Length == 3)
                    {
                        whenFalse = TreeNodeToExpression(arguments[2], currentVertex);
                        argumentsHaveDifferentTypes = IsSameType(arguments[1], arguments[2], currentVertex) == null;
                    }
                    else
                    {
                        // if no else statement is given, Excel defaults to FALSE
                        whenFalse = LiteralExpression(SyntaxKind.FalseLiteralExpression);
                        argumentsHaveDifferentTypes = !IsTypeBoolean(arguments[1], currentVertex);
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
                }
                case "AND":
                case "OR":
                case "XOR":
                {
                    if (arguments.Length == 0) return FunctionError(functionName, arguments);
                    if (arguments.Length == 1) return TreeNodeToExpression(arguments[0], currentVertex);
                    return ParenthesizedExpression(FoldBinaryExpression(functionName, arguments, currentVertex));
                }
                case "NOT":
                {
                    if (arguments.Length != 1) return FunctionError(functionName, arguments);
                    return PrefixUnaryExpression(SyntaxKind.LogicalNotExpression, 
                        TreeNodeToExpression(arguments[0], currentVertex));
                }
                case "TRUE":
                {
                    return LiteralExpression(SyntaxKind.TrueLiteralExpression);
                }
                case "FALSE":
                {
                    return LiteralExpression(SyntaxKind.FalseLiteralExpression);
                }

                // IS functions
                case "ISBLANK":
                {
                    if (arguments.Length != 1) return FunctionError(functionName, arguments);
                    var leftExpression = TreeNodeToExpression(arguments[0], currentVertex);
                    var rightExpression = IdentifierName("EmptyCell");
                    return GenerateBinaryExpression(functionName, leftExpression, rightExpression);
                }
                case "ISLOGICAL":
                {
                    if (arguments.Length != 1) return FunctionError(functionName, arguments);
                    var leftExpression = TreeNodeToExpression(arguments[0], currentVertex);
                    var rightExpression = PredefinedType( Token(SyntaxKind.BoolKeyword));
                    return GenerateBinaryExpression(functionName, leftExpression, rightExpression);
                }
                case "ISNONTEXT":
                {
                    // !(argument[0] is string)
                    return PrefixUnaryExpression(
                        SyntaxKind.LogicalNotExpression,
                        ParenthesizedExpression(FunctionToExpression("ISTEXT", arguments, currentVertex)));
                }
                case "ISNUMBER":
                {
                    return InvocationExpression(IdentifierName("IsNumeric"))
                        .AddArgumentListArguments(Argument(TreeNodeToExpression(arguments[0], currentVertex)));
                }
                case "ISTEXT":
                {
                    if (arguments.Length != 1) return FunctionError(functionName, arguments);
                    var leftExpression = TreeNodeToExpression(arguments[0], currentVertex);
                    var rightExpression = PredefinedType(Token(SyntaxKind.StringKeyword));
                    return GenerateBinaryExpression(functionName, leftExpression, rightExpression);
                }
                case "ISERR":
                case "ISERROR":
                case "ISNA":
                {
                    if (arguments.Length != 1) return FunctionError(functionName, arguments);
                    var leftExpression = TreeNodeToExpression(arguments[0], currentVertex);
                    var rightExpression = IdentifierName("FormulaError");
                    return GenerateBinaryExpression(functionName, leftExpression, rightExpression);
                }

                // comparators
                case "=":
                {
                    if (arguments.Length != 2) return FunctionError(functionName, arguments);
                    var leftExpression = TreeNodeToExpression(arguments[0], currentVertex);
                    var rightExpression = TreeNodeToExpression(arguments[1], currentVertex);

                    ExpressionSyntax equalsExpression;
                    var leftType = GetType(arguments[0], currentVertex);
                    var rightType = GetType(arguments[1], currentVertex);
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
                        // if types different, use our custom Compare method (== would throw exception if types are different)
                        equalsExpression = InvocationExpression(IdentifierName("Compare"))
                            .AddArgumentListArguments(Argument(leftExpression), Argument(rightExpression));
                    }

                    return equalsExpression;
                }
                case "<>":
                {
                    if (arguments.Length != 2) return FunctionError(functionName, arguments);

                    // see above code for "=", only if types match and are not strings, use ==
                    var leftType = GetType(arguments[0], currentVertex);
                    var rightType = GetType(arguments[1], currentVertex);
                    if (leftType.HasValue && rightType.HasValue && leftType.Value == rightType.Value &&
                        (leftType.Value != CellType.Text || rightType.Value != CellType.Text))
                    {
                        return GenerateBinaryExpression(functionName, arguments, currentVertex);
                    }

                    return PrefixUnaryExpression(SyntaxKind.LogicalNotExpression,
                        FunctionToExpression("=", arguments, currentVertex));
                }
                case "<":
                case "<=":
                case ">=":
                case ">":
                {
                    return GenerateBinaryExpression(functionName, arguments, currentVertex);
                }

                // strings
                case "&":
                case "CONCATENATE":
                {
                    if (arguments.Length < 2) return FunctionError(functionName, arguments);
                    return FoldBinaryExpression("+", arguments, currentVertex);
                }

                // other
                case "DATE":
                {
                    if (arguments.Length != 3) return FunctionError(functionName, arguments);
                    return ObjectCreationExpression(
                            IdentifierName("DateTime"))
                        .AddArgumentListArguments(
                            Argument(TreeNodeToExpression(arguments[0], currentVertex)),
                            Argument(TreeNodeToExpression(arguments[1], currentVertex)),
                            Argument(TreeNodeToExpression(arguments[2], currentVertex))
                        );
                }
                case "TODAY":
                {
                    // DateTime.Now
                    return MemberAccessExpression(
                        SyntaxKind.SimpleMemberAccessExpression,
                        IdentifierName("DateTime"),
                        IdentifierName("Now"));
                }
                case "SECOND":
                case "MINUTE":
                case "HOUR":
                case "DAY":
                case "MONTH":
                case "YEAR":
                {
                    if (arguments.Length != 1) return FunctionError(functionName, arguments);
                    return MemberAccessExpression(
                        SyntaxKind.SimpleMemberAccessExpression,
                        TreeNodeToExpression(arguments[0], currentVertex),
                        IdentifierName(functionName.ToTitleCase()));
                }

                default:
                {
                    return CommentExpression($"Function {functionName} not implemented yet! Args: " +
                                             $"{string.Join("\n", arguments.Select(a => TreeNodeToExpression(a, currentVertex)))}", true);
                }
            }
        }

        private readonly Dictionary<string, CellType> _functionToCellTypeDictionary = new Dictionary<string, CellType>
        {
            { "+", CellType.Unknown },
            { "-", CellType.Number },
            { "*", CellType.Number },
            { "/", CellType.Number },
            { "%", CellType.Number },
            { "^", CellType.Number },
            { "ROUND", CellType.Number },
            { "ROUNDUP", CellType.Number },
            { "ROUNDDOWN", CellType.Number },
            { "SUM", CellType.Number },
            { "MIN", CellType.Number },
            { "MAX", CellType.Number },
            { "COUNT", CellType.Number },
            { "AVERAGE", CellType.Number },

            { "HLOOKUP", CellType.Unknown },
            { "VLOOKUP", CellType.Unknown },
            { "CHOOSE", CellType.Unknown },
            { "MATCH", CellType.Number },
            { "INDEX", CellType.Unknown },

            { "IF", CellType.Unknown },
            { "AND", CellType.Bool },
            { "NOT", CellType.Bool },
            { "OR", CellType.Bool },
            { "XOR", CellType.Bool },
            { "TRUE", CellType.Bool },
            { "FALSE", CellType.Bool },
            
            { "ISBLANK", CellType.Bool },
            { "ISLOGICAL", CellType.Bool },
            { "ISNOTEXT", CellType.Bool },
            { "ISNUMBER", CellType.Bool },
            { "ISTEXT", CellType.Bool },
            { "ISERR", CellType.Unknown },
            { "ISERROR", CellType.Unknown },
            { "ISNA", CellType.Unknown },

            { "=", CellType.Bool },
            { "<>", CellType.Bool },
            { "<", CellType.Number },
            { "<=", CellType.Number },
            { ">=", CellType.Number },
            { ">", CellType.Number },

            { "&", CellType.Text },
            { "CONCATENATE", CellType.Text },

            { "DATE", CellType.Date },
            { "SECOND", CellType.Date },
            { "MINUTE", CellType.Date },
            { "HOUR", CellType.Date },
            { "DAY", CellType.Date },
            { "MONTH", CellType.Date },
            { "TODAY", CellType.Date },
        };

        private ExpressionSyntax RoundFunction(string roundFunction, ParseTreeNode[] arguments, CellVertex currentVertex)
        {
            if (arguments.Length != 2) return FunctionError(roundFunction, arguments);
            return InvocationExpression(IdentifierName(roundFunction))
                .AddArgumentListArguments(
                    Argument(TreeNodeToExpression(arguments[0], currentVertex)),
                    Argument(TreeNodeToExpression(arguments[1], currentVertex)));
        }

        // gets the type of a node
        // if multiple types are found, or type is unknown or dynamic, return null
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

        // checks if two parse tree nodes have the same type
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
        
        private ExpressionSyntax CollectionOf(params ExpressionSyntax[] expressions)
        {
            // avoid Collection(Collection)
            if (expressions.Length == 1 && expressions[0] is InvocationExpressionSyntax inv
                                        && inv.Expression is MemberAccessExpressionSyntax maes
                                        && ((IdentifierNameSyntax)maes.Expression).Identifier.Text == "Collection")
                return expressions[0];

            return ClassFunctionCall("Collection", "Of", expressions);
        }

        private ExpressionSyntax MatrixOf(params ExpressionSyntax[] expressions)
        {
            return ClassFunctionCall("Matrix", "Of", expressions);
        }

        private ExpressionSyntax ColumnOf(params ExpressionSyntax[] expressions)
        {
            return ClassFunctionCall("Column", "Of", expressions);
        }

        private ExpressionSyntax RowOf(params ExpressionSyntax[] expressions)
        {
            return ClassFunctionCall("Row", "Of", expressions);
        }

        private ExpressionSyntax ClassFunctionCall(string className, string functionName, params ExpressionSyntax[] expressions)
        {
            return InvocationExpression(
                    MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                        IdentifierName(className),
                        IdentifierName(functionName)))
                .AddArgumentListArguments(expressions.Select(Argument).ToArray());
        }

        private readonly SortedList<string, (SyntaxKind syntaxKind, bool parenthesize)> _binaryOperators
            = new SortedList<string, (SyntaxKind syntaxKind, bool parenthesize)>
        {
            {"+", (SyntaxKind.AddExpression, false)},
            {"-", (SyntaxKind.SubtractExpression, false)},
            {"/", (SyntaxKind.DivideExpression, true)},
            {"*", (SyntaxKind.MultiplyExpression, true)},

            {"ISBLANK", (SyntaxKind.IsExpression, false)},
            {"ISLOGICAL", (SyntaxKind.IsExpression, false)},
            {"ISNOTEXT", (SyntaxKind.IsExpression, false)},
            {"ISNUMBER", (SyntaxKind.IsExpression, false)},
            {"ISTEXT", (SyntaxKind.IsExpression, false)},
            {"ISERR", (SyntaxKind.IsExpression, false)},
            {"ISERROR", (SyntaxKind.IsExpression, false)},
            {"ISNA", (SyntaxKind.IsExpression, false)},

            {"<>", (SyntaxKind.NotEqualsExpression, false)},
            {"<", (SyntaxKind.LessThanExpression, false)},
            {"<=", (SyntaxKind.LessThanOrEqualExpression, false)},
            {">=", (SyntaxKind.GreaterThanOrEqualExpression, false)},
            {">", (SyntaxKind.GreaterThanExpression, false)},
            {"AND", (SyntaxKind.LogicalAndExpression, true)},
            {"OR", (SyntaxKind.LogicalOrExpression, true)},
            {"XOR", (SyntaxKind.ExclusiveOrExpression, true)},
        };

        private ExpressionSyntax GenerateBinaryExpression(string functionName, ParseTreeNode[] arguments, CellVertex vertex)
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
        private ExpressionSyntax FoldBinaryExpression(string functionName, ParseTreeNode[] arguments, CellVertex vertex)
        {
            var syntaxKind = _binaryOperators[functionName].syntaxKind;
            // do not parenthesize
            return arguments.Select(a => TreeNodeToExpression(a, vertex))
                .Aggregate((acc, right) => BinaryExpression(syntaxKind, acc, right));
        }

        private ExpressionSyntax VariableReferenceFromAddressToExpression(string address, Vertex currentVertex)
        {
            if (AddressToVertexDictionary.TryGetValue((currentVertex.WorksheetName, address)
                , out var variableVertex))
            {
                return VariableReferenceToExpression(variableVertex, currentVertex);
            }

            return CommentExpression($"{address} not found", true);
        }

        private static ExpressionSyntax VariableReferenceToExpression(Vertex variableVertex, Vertex currentVertex)
        {
            return currentVertex.Class == variableVertex.Class
                ? (ExpressionSyntax) IdentifierName(variableVertex.VariableName)
                : MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                    IdentifierName(variableVertex.Class.Name), IdentifierName(variableVertex.VariableName));
        }

        private ExpressionSyntax RangeOrNamedRangeToExpression(ParseTreeNode node, CellVertex currentVertex)
        {
            if (node.Term.Name == "NamedRange")
                return NamedRangeToExpression(node, currentVertex);
            if (node.Term.Name == "ReferenceFunctionCall" && node.GetFunction() == ":" ||
                node.ChildNodes.Count == 2 && (node.ChildNodes[1].Term.Name == "VRange" || node.ChildNodes[1].Term.Name == "HRange"))
                return RangeToExpression(node, currentVertex);
            if (node.ChildNodes.Count == 1)
                return RangeOrNamedRangeToExpression(node.ChildNodes[0], currentVertex);
            return CommentExpression($"Argument is not a (named) range!", true);
        }

        // node is ReferenceFunctionCall with function :
        private ExpressionSyntax RangeToExpression(ParseTreeNode node, CellVertex currentVertex)
        {
            var range = node.NodeToString(currentVertex.Formula);
            if (RangeDictionary.TryGetValue(range, out var rangeVertex))
                return VariableReferenceToExpression(rangeVertex, currentVertex);
            return CommentExpression($"Did not find variable for range {range}", true);
        }

        // node is NamedRange
        private ExpressionSyntax NamedRangeToExpression(ParseTreeNode node, CellVertex currentVertex)
        {
            var namedRangeName = node.FindTokenAndGetText();
            return NameDictionary.TryGetValue(namedRangeName, out var namedRangeVertex)
                ? VariableReferenceToExpression(namedRangeVertex, currentVertex)
                : CommentExpression($"Did not find variable for named range {namedRangeName}", true);
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

            CellVertex cellVertex;
            if (vertex is RangeVertex rangeVertex)
            {
                if (rangeVertex.Type == RangeVertex.RangeType.Empty)
                    return "Matrix";
                if (rangeVertex.Type == RangeVertex.RangeType.Single)
                    cellVertex = rangeVertex.GetSingleElement();
                else
                    return rangeVertex.Type.ToString();
            }
            else
            {
                cellVertex = (CellVertex)vertex;
            }

            switch (cellVertex.CellType)
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

        private ExpressionSyntax VertexValueToExpression(Vertex vertex)
        {
            CellVertex cellVertex;

            if (vertex is RangeVertex rangeVertex)
            {
                if (rangeVertex.Type == RangeVertex.RangeType.Single)
                    cellVertex = rangeVertex.GetSingleElement();
                else
                    return RangeVertexToExpression(rangeVertex);
            }
            else
            {
                cellVertex = (CellVertex) vertex;
            }

            var vertexValue = cellVertex.Value;
            switch (cellVertex.CellType)
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
