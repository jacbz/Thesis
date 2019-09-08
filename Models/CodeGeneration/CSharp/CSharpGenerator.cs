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
        protected string[] BlockedVariableNames =>
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

            // construct variable name -> Vertex list dictionary
            var variableNameToVerticesDictionary = new Dictionary<string, List<Vertex>>();
            foreach(var cellVertex in Graph.ConstantsAndConstantFormulas.Select(t => t.cellVertex).Distinct())
                variableNameToVerticesDictionary.Add(cellVertex.Name, new List<Vertex> { cellVertex });
            foreach(var kvp in Graph.FormulaFunctionDictionary)
            {
                var name = kvp.Value.Name;
                if (variableNameToVerticesDictionary.ContainsKey(name))
                    variableNameToVerticesDictionary[name].Add(kvp.Key);
                else
                    variableNameToVerticesDictionary.Add(name, new List<Vertex> { kvp.Key });
            }
            foreach (var kvp in Graph.OutputFieldFunctionDictionary)
            {
                var name = kvp.Value.Name;
                if (variableNameToVerticesDictionary.ContainsKey(name))
                    variableNameToVerticesDictionary[name].Add(kvp.Key);
                else
                    variableNameToVerticesDictionary.Add(name, new List<Vertex> { kvp.Key });
            }

            return new Code(sourceCode, variableNameToVerticesDictionary, null);
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


        private ExpressionSyntax CellVertexToConstantOrVariable(CellVertex cellVertex, Vertex currentVertex)
        {
            return AddressToVertexDictionary.TryGetValue((currentVertex.WorksheetName, cellVertex.StringAddress), out var c)
                ? IdentifierName(c.Name)
                : ConstantToExpression(cellVertex);
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

        private ExpressionSyntax ConstantToExpression(Constant constant)
        {
            return ConstantToExpression(constant.ConstantValue, constant.ConstantType);
        }

        private ExpressionSyntax ConstantToExpression(CellVertex cellVertex)
        {
            return ConstantToExpression(cellVertex.Value, cellVertex.CellType);
        }

        private ExpressionSyntax ConstantToExpression(dynamic constantValue, CellType constantType)
        {
            switch (constantType)
            {
                case CellType.Bool:
                    return LiteralExpression(constantValue ? SyntaxKind.TrueLiteralExpression : SyntaxKind.FalseLiteralExpression);
                case CellType.Text:
                    return LiteralExpression(SyntaxKind.StringLiteralExpression, Literal(constantValue));
                case CellType.Number:
                    return LiteralExpression(SyntaxKind.NumericLiteralExpression, Literal(constantValue));
                case CellType.Date:
                    return InvocationExpression(MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                            IdentifierName("DateTime"), IdentifierName("Parse")))
                        .AddArgumentListArguments(Argument(LiteralExpression(SyntaxKind.StringLiteralExpression,
                            Literal(constantValue.ToString()))));
                case CellType.Error:
                    return ObjectCreationExpression(IdentifierName("FormulaError"))
                        .AddArgumentListArguments(Argument(ParseExpression(constantValue)));
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
