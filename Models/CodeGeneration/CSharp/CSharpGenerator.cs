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
using Syncfusion.Windows.Shared;
using Thesis.Models.FunctionGeneration;
using Thesis.Models.VertexTypes;
using static Microsoft.CodeAnalysis.CSharp.SyntaxFactory;
using Formatter = Microsoft.CodeAnalysis.Formatting.Formatter;

namespace Thesis.Models.CodeGeneration.CSharp
{
    public partial class CSharpGenerator : CodeGenerator
    {
        protected override string[] BlockedVariableNames =>
            new[]
            {
                "Matrix", "Collection", "Row", "Column", "EmptyCell", "Empty",
                "bool", "byte", "sbyte", "short", "ushort", "int", "uint", "long", "ulong", "double", "float",
                "decimal", "string", "char", "void", "object", "typeof", "sizeof", "null", "true", "false", "if",
                "else", "while", "for", "foreach", "do", "switch", "case", "default", "lock", "try", "throw",
                "catch", "finally", "goto", "break", "continue", "return", "public", "private", "internal", "protected",
                "static", "readonly", "sealed", "const", "fixed", "stackalloc", "volatile", "new", "override", "abstract",
                "virtual", "event", "extern", "ref", "out", "in", "is", "as", "params", "__arglist", "__makeref",
                "__reftype", "__refvalue", "this", "base", "namespace", "using", "class", "struct", "interface",
                "enum", "delegate", "checked", "unchecked", "unsafe", "operator", "implicit", "explicit"
            };

        public CSharpGenerator(Graph graph, 
            Dictionary<(string worksheet, string address), CellVertex> addressToVertexDictionary, 
            Dictionary<string, RangeVertex> rangeDictionary, 
            Dictionary<string, Vertex> nameDictionary) : 
            base(graph, addressToVertexDictionary, rangeDictionary, nameDictionary)
        {
        }

        public override async Task<Code> GenerateCodeAsync(TestResults testResults = null)
        {
            return await Task.Run(() => GenerateCode(testResults));
        }

        private FieldDeclarationSyntax DeclareFieldWithExpression(string name, Expression expression)
        {
            return FieldDeclaration(
                VariableDeclaration(IdentifierName(CellTypeToTypeString(expression.GetCellType())))
                    .AddVariables(
                        VariableDeclarator(
                                Identifier(name))
                            .WithInitializer(
                                EqualsValueClause(ExpressionToCode(expression)))));
        }

        private FieldDeclarationSyntax DeclareFieldWithRangeVertex(RangeVertex rangeVertex)
        {
            return FieldDeclaration(
                VariableDeclaration(IdentifierName(GetTypeString(rangeVertex)))
                    .AddVariables(
                        VariableDeclarator(
                                Identifier(rangeVertex.Name))
                            .WithInitializer(
                                EqualsValueClause(RangeVertexToExpression(rangeVertex)))));
        }

        public MemberDeclarationSyntax FormulaFunctionToCode(FormulaFunction formulaFunction)
        {
            var methodDeclaration = MethodDeclaration(
                IdentifierName(CellTypeToTypeString(formulaFunction.ReturnType)), Identifier(formulaFunction.Name))
                .AddParameterListParameters(formulaFunction.Parameters.Select(inputReference =>
                    Parameter(Identifier(inputReference.VariableName))
                        .WithType(IdentifierName(
                            inputReference is InputReference cellReference
                                ? CellTypeToTypeString(cellReference.InputType)
                                : "Matrix"))).ToArray())
                .WithBody(Block(ReturnStatement(ExpressionToCode(formulaFunction.Expression))));
            return methodDeclaration;
        }

        public StatementSyntax StatementToCode(Statement statement)
        {
            var rightSideExpression = statement is FunctionInvocationStatement functionInvocationStatement
                ? FunctionInvocationStatementToCode(functionInvocationStatement)
                : throw new Exception("Wrong statement");

            return LocalDeclarationStatement(VariableDeclaration(
                   IdentifierName(CellTypeToTypeString(statement.VariableType)))
                .AddVariables(
                    VariableDeclarator(Identifier(statement.VariableName))
                        .WithInitializer(EqualsValueClause(rightSideExpression))));
        }

        public ExpressionSyntax FunctionInvocationStatementToCode(
            FunctionInvocationStatement functionInvocationStatement)
        {
            var parameters = functionInvocationStatement.Parameters;
            if (parameters.Length == 2 &&
                _binaryOperators.ContainsKey(functionInvocationStatement.FunctionName))
                return BinaryExpression(_binaryOperators[functionInvocationStatement.FunctionName],
                    IdentifierName(parameters[0].VariableName),
                    IdentifierName(parameters[1].VariableName));
            return InvocationExpression(IdentifierName(functionInvocationStatement.FunctionName))
                .AddArgumentListArguments(functionInvocationStatement.GetParameterNames()
                    .Select(parameterName => Argument(IdentifierName(parameterName)))
                    .ToArray());
        }

        public MemberDeclarationSyntax OutputFieldFunctionToCode(OutputFieldFunction function)
        {
            var statements = function.Statements.Select(StatementToCode).ToList();
            statements.Add(ReturnStatement(IdentifierName(function.Statements.Last().VariableName)));

            var methodDeclaration = MethodDeclaration(
                    IdentifierName(CellTypeToTypeString(function.ReturnType)), Identifier(function.Name))
                .AddParameterListParameters(function.Parameters.Select(inputReference =>
                    Parameter(Identifier(inputReference.VariableName))
                        .WithType(IdentifierName(CellTypeToTypeString(inputReference.InputType)))).ToArray())
                .WithBody(Block(statements.ToArray()));
            return methodDeclaration;
        }

        public Code GenerateCode(TestResults testResults = null)
        {
            var constantsAndFormulaConstants = Graph.ConstantsAndConstantFormulas
                .Select(tuple => DeclareFieldWithExpression(tuple.cellVertex.Name, tuple.expression))
                .ToArray<MemberDeclarationSyntax>();

            var ranges = Graph.RangeDictionary.Values
                .OrderBy(v => v.Name)
                .Select(DeclareFieldWithRangeVertex)
                .ToArray<MemberDeclarationSyntax>();

            var outputFields = Graph.OutputFieldFunctionDictionary.Values
                .Distinct()
                .OrderBy(v => v.Name)
                .Select(OutputFieldFunctionToCode)
                .ToArray();

            var functionList = Graph.FormulaFunctionDictionary.Values
                .Distinct()
                .OrderBy(v => v.Name)
                .Select(FormulaFunctionToCode).ToArray();

            var compilationUnit = CompilationUnit();
            
            if (constantsAndFormulaConstants.Length > 0)
                compilationUnit = compilationUnit
                    .AddMembers(IncompleteMember(IdentifierName("COMMENT1")))
                    .AddMembers(constantsAndFormulaConstants);

            if (ranges.Length > 0)
                compilationUnit = compilationUnit
                    .AddMembers(IncompleteMember(IdentifierName("COMMENT2")))
                    .AddMembers(ranges);

            if (outputFields.Length > 0)
                compilationUnit = compilationUnit
                    .AddMembers(IncompleteMember(IdentifierName("COMMENT3")))
                    .AddMembers(outputFields);

            if (functionList.Length > 0)
                compilationUnit = compilationUnit
                    .AddMembers(IncompleteMember(IdentifierName("COMMENT4")))
                    .AddMembers(functionList);

            // format with Roslyn formatter
            var workspace = new AdhocWorkspace();
            var options = workspace.Options;
            var node = Formatter.Format(compilationUnit, workspace, options);
            var sourceCode = node.ToFullString();

            sourceCode = sourceCode
                .Replace("COMMENT1 ", "// Global constants and formula constants\r\n")
                .Replace("COMMENT2 ", "// Referenced ranges\r\n")
                .Replace("COMMENT3 ", "// Output fields\r\n")
                .Replace("COMMENT4 ", "// Formulas\r\n");

            // format with our custom formatter
            sourceCode = FormatCode(sourceCode);

            return new Code(sourceCode, null, null);

            //_useDynamic = new HashSet<Vertex>();

            //// generate variable names for all
            //GenerateVariableNamesForAll();

            //// namespace Thesis
            //var @namespace = NamespaceDeclaration(ParseName("Thesis")).NormalizeWhitespace();
            //// using System;
            //@namespace = @namespace.AddUsings(
            //    UsingDirective(ParseName("System")),
            //    UsingDirective(ParseName("System.Linq")));

            //// public class ThesisResult (Main class)
            //var resultClass = GenerateResultClass(testResults);
            //@namespace = @namespace.AddMembers(resultClass);

            //var normalClasses = new List<MemberDeclarationSyntax>();
            //var staticClasses = new List<MemberDeclarationSyntax>();
            //var classesCode = new List<ClassCode>();

            //// static classes first (must determine which types are dynamic)
            //foreach (var generatedClass in ClassCollection.Classes.OrderBy(v => !v.IsStaticClass))
            //{
            //    var (newClass, classCode) = GenerateClass(generatedClass, testResults);
            //    classesCode.Add(classCode);

            //    if (generatedClass.IsStaticClass)
            //        staticClasses.Add(newClass);
            //    else
            //        normalClasses.Add(newClass);
            //}

            //// show static classes on top for aesthetic reasons
            //@namespace = @namespace
            //    .AddMembers(staticClasses.ToArray())
            //    .AddMembers(normalClasses.ToArray());

            //// format with Roslyn formatter
            //var workspace = new AdhocWorkspace();
            //var options = workspace.Options;
            //var node = Formatter.Format(@namespace, workspace, options);
            //var sourceCode = node.ToFullString();

            //// format with our custom formatter
            //sourceCode = FormatCode(sourceCode);

            //var variableNameToVertexDictionary = ClassCollection.Classes
            //    .SelectMany(c => c.Vertices)
            //    .ToDictionary(v => (v.Class.Name, v.Name));
            //return new Code(sourceCode, variableNameToVertexDictionary, new CSharpTester(classesCode));
        }


        private TypeSyntax GenerateTypeSyntaxWithTestResult(Vertex vertex, TestResults testResults)
        {
            var typeString = GetTypeString(vertex);
            if (testResults == null)
                return ParseTypeName(typeString);
            var testResult = testResults[vertex];
            if (testResult == null || testResult.TestResultType == TestResultType.Ignore)
                return ParseTypeName(typeString);

            return IdentifierName(
                Identifier(
                    TriviaList(Comment(testResult.ToString()), LineFeed),
                    typeString,
                    TriviaList()));
        }

        private static IdentifierNameSyntax GenerateIdentifierNameSyntaxWithTestResult(Vertex vertex, TestResults testResults)
        {
            if (testResults == null)
                return IdentifierName(vertex.Name);
            var testResult = testResults[vertex];
            if (testResult == null || testResult.TestResultType == TestResultType.Ignore)
                return IdentifierName(vertex.Name);

            return IdentifierName(
                Identifier(
                    TriviaList(Comment(testResult.ToString()), LineFeed),
                    vertex.Name,
                    TriviaList()));
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

        private string RangeToVariableName(Expression expression)
        {
            return expression is GlobalReference gr
                ? gr.ReferencedVertex.Name.ToString()
                : throw new Exception("Unidentified reference");
        }

        private static ExpressionSyntax VariableReferenceToExpression(Vertex variableVertex, Vertex currentVertex)
        {
            return currentVertex.Class == variableVertex.Class
                ? (ExpressionSyntax) IdentifierName(variableVertex.Name)
                : MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                    IdentifierName(variableVertex.Class.Name), IdentifierName(variableVertex.Name));
        }

        private ExpressionSyntax CellVertexToConstantOrVariable(CellVertex cellVertex, Vertex currentVertex)
        {
            return AddressToVertexDictionary.TryGetValue((currentVertex.WorksheetName, cellVertex.StringAddress), out var c) 
                ? VariableReferenceToExpression(c, currentVertex) 
                : VertexValueToExpression(cellVertex);
        }

        private string GetTypeString(Vertex vertex)
        {
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

            return CellTypeToTypeString(cellVertex.CellType);
        }

        private string CellTypeToTypeString(CellType cellType)
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
                case CellType.Error:
                case CellType.Unknown:
                    return "dynamic";
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
                    return InvocationExpression(MemberAccessExpression( SyntaxKind.SimpleMemberAccessExpression,
                            IdentifierName("DateTime"), IdentifierName("Parse")))
                        .AddArgumentListArguments(Argument(LiteralExpression(SyntaxKind.StringLiteralExpression,
                            Literal(vertexValue.ToString()))));
                case CellType.Error:
                    return ObjectCreationExpression(IdentifierName("FormulaError"))
                        .AddArgumentListArguments(Argument(ParseExpression(vertexValue)));
                case CellType.Unknown:
                    return IdentifierName("Empty");
                default:
                    return LiteralExpression(SyntaxKind.NullLiteralExpression);
            }
        }

        private ExpressionSyntax ConstantToExpression(Constant constant)
        {
            switch (constant.ConstantType)
            {
                case CellType.Bool:
                    return LiteralExpression(constant.ConstantValue ? SyntaxKind.TrueLiteralExpression : SyntaxKind.FalseLiteralExpression);
                case CellType.Text:
                    return LiteralExpression(SyntaxKind.StringLiteralExpression, Literal(constant.ConstantValue));
                case CellType.Number:
                    return LiteralExpression(SyntaxKind.NumericLiteralExpression, Literal(constant.ConstantValue));
                case CellType.Date:
                    return InvocationExpression(MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                            IdentifierName("DateTime"), IdentifierName("Parse")))
                        .AddArgumentListArguments(Argument(LiteralExpression(SyntaxKind.StringLiteralExpression,
                            Literal(constant.ConstantValue.ToString()))));
                case CellType.Error:
                    return ObjectCreationExpression(IdentifierName("FormulaError"))
                        .AddArgumentListArguments(Argument(ParseExpression(constant.ConstantValue)));
                case CellType.Unknown:
                    return IdentifierName("Empty");
                default:
                    return LiteralExpression(SyntaxKind.NullLiteralExpression);
            }
        }
        private ExpressionSyntax CommentExpression(string comment, bool isError = false)
        {
            return IdentifierName(
                Identifier(
                    TriviaList(Comment("/* " + (isError ? "ERROR: " : "") + comment + " */")),
                    "",
                    TriviaList()));
        }
    }
}
