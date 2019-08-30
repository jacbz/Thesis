using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Irony.Parsing;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Thesis.Models.VertexTypes;
using XLParser;
using static Microsoft.CodeAnalysis.CSharp.SyntaxFactory;
using Formatter = Microsoft.CodeAnalysis.Formatting.Formatter;

namespace Thesis.Models.CodeGeneration.CSharp
{
    public partial class CSharpGenerator : CodeGenerator
    {
        // vertices in this list must have type dynamic
        private HashSet<Vertex> _useDynamic;

        protected override string[] LanguageKeywords =>
            new[]
            {
                "bool", "byte", "sbyte", "short", "ushort", "int", "uint", "long", "ulong", "double", "float",
                "decimal", "string", "char", "void", "object", "typeof", "sizeof", "null", "true", "false", "if",
                "else", "while", "for", "foreach", "do", "switch", "case", "default", "lock", "try", "throw",
                "catch", "finally", "goto", "break", "continue", "return", "public", "private", "internal", "protected",
                "static", "readonly", "sealed", "const", "fixed", "stackalloc", "volatile", "new", "override", "abstract",
                "virtual", "event", "extern", "ref", "out", "in", "is", "as", "params", "__arglist", "__makeref",
                "__reftype", "__refvalue", "this", "base", "namespace", "using", "class", "struct", "interface",
                "enum", "delegate", "checked", "unchecked", "unsafe", "operator", "implicit", "explicit"
            };

        public CSharpGenerator(ClassCollection classCollection, 
            Dictionary<(string worksheet, string address), CellVertex> addressToVertexDictionary, 
            Dictionary<string, RangeVertex> rangeDictionary, 
            Dictionary<string, Vertex> nameDictionary) : 
            base(classCollection, addressToVertexDictionary, rangeDictionary, nameDictionary)
        {
        }

        public override async Task<Code> GenerateCodeAsync(TestResults testResults = null)
        {
            return await Task.Run(() => GenerateCode(testResults));
        }

        public Code GenerateCode(TestResults testResults = null)
        {
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

            var variableNameToVertexDictionary = ClassCollection.Classes
                .SelectMany(c => c.Vertices)
                .ToDictionary(v => (v.Class.Name, v.Name));
            return new Code(sourceCode, variableNameToVertexDictionary, new CSharpTester(classesCode));
        }

        private (ClassDeclarationSyntax classDeclarationSyntax, ClassCode classCode)
            GenerateClass(Class @class, TestResults testResults = null)
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
                        ? IdentifierName(formula.Name)
                        : GenerateIdentifierWithComment(formula.Name,
                            testResults[formula].ToString());

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
                            testResults[formula].ToString());
                    var variableDeclaration = VariableDeclaration(type)
                        .AddVariables(VariableDeclarator(formula.Name)
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
                            .AddVariables(VariableDeclarator(constant.Name)
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
                                .AddVariables(VariableDeclarator(formula.Name)))
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
                statements.Add(ReturnStatement(IdentifierName(@class.OutputVertex.Name)));

                // public {type} Calculate()
                var outputField = @class.OutputVertex;
                var calculateMethod = MethodDeclaration(ParseTypeName(GetTypeString(outputField)), "Calculate")
                    .AddModifiers(
                        Token(SyntaxKind.PublicKeyword))
                    .WithBody(Block(statements));
                return calculateMethod;
            }
        }

        private ClassDeclarationSyntax GenerateResultClass(TestResults testResults = null)
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
                            testResults[generatedClass.OutputVertex].ToString());

                    // {type} {outputvertexname} = new {classname}().Calculate()
                    methodBody.Add(LocalDeclarationStatement(
                        VariableDeclaration(type)
                            .AddVariables(VariableDeclarator(generatedClass.OutputVertex.Name)
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

        private ExpressionSyntax RangeToExpression(ParseTreeNode node, CellVertex currentVertex)
        {
            var range = node.NodeToString(currentVertex.Formula);
            if (RangeDictionary.TryGetValue(range, out var rangeVertex))
                return VariableReferenceToExpression(rangeVertex, currentVertex);
            return CommentExpression($"Did not find variable for range {range}", true);
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
                        .Select<CellVertex[], ExpressionSyntax>(rowArray => RowOf(rowArray.Select(cellVertex => CellVertexToConstantOrVariable(cellVertex, rangeVertex))
                            .ToArray()))
                        .ToArray());
                }
                default:
                    return null;
            }
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

        // node is NamedRange
        private ExpressionSyntax NamedRangeToExpression(ParseTreeNode node, CellVertex currentVertex)
        {
            var namedRangeName = node.FindTokenAndGetText();
            return NameDictionary.TryGetValue(namedRangeName, out var namedRangeVertex)
                ? VariableReferenceToExpression(namedRangeVertex, currentVertex)
                : CommentExpression($"Did not find variable for named range {namedRangeName}", true);
        }

        // if a cell vertex already exists, use that one, else format as constant

        // gets the type of a node
        // if multiple types are found, or type is unknown or dynamic, return null

        // checks if two parse tree nodes have the same type

        // e.g. function &&, arg [a,b,c,d] => a && (b && (c && d))

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
