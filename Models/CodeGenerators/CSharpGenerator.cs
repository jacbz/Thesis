using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Irony.Parsing;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
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

        public override string GenerateCode()
        {
            // namespace Thesis
            var @namespace = SyntaxFactory.NamespaceDeclaration(SyntaxFactory.ParseName("Thesis")).NormalizeWhitespace();
            // using System;
            @namespace = @namespace.AddUsings(SyntaxFactory.UsingDirective(SyntaxFactory.ParseName("System")));

            // public class ThesisResult (Main class)
            var resultClass = SyntaxFactory
                .ClassDeclaration("Result")
                .AddModifiers(SyntaxFactory.Token(SyntaxKind.PublicKeyword));
            var syntax = SyntaxFactory.ParseStatement("// TBD");
            var mainMethod = SyntaxFactory.MethodDeclaration(SyntaxFactory.ParseTypeName("void"), "Main")
                .AddModifiers(SyntaxFactory.Token(SyntaxKind.PublicKeyword))
                .AddModifiers(SyntaxFactory.Token(SyntaxKind.StaticKeyword))
                .AddParameterListParameters(SyntaxFactory
                    .Parameter(SyntaxFactory.Identifier("args"))
                    .WithType(SyntaxFactory.ParseTypeName("string[]")))
                .WithBody(SyntaxFactory.Block(syntax));
            resultClass = resultClass.AddMembers(mainMethod);
            @namespace = @namespace.AddMembers(resultClass);

            foreach(var generatedClass in generatedClasses)
            {
                var vertices = generatedClass.Vertices.ToList();
                vertices.Reverse(); // as topological resulted in the output field being at the bottom
                var usedVariableNames = new HashSet<string>();

                // public class {generatedClass.Name}
                var newClass = SyntaxFactory
                    .ClassDeclaration(generatedClass.Name)
                    .AddModifiers(SyntaxFactory.Token(SyntaxKind.PublicKeyword));


                // add fields
                var lookup = vertices.ToLookup(v => v.NodeType == NodeType.Constant); // split list into two
                var constants = lookup[true].ToList();
                var formulas = lookup[false].ToList();

                foreach (var vertex in constants)
                {
                    var variableName = vertex.VariableName;
                    variableName = GenerateNonDuplicateName(usedVariableNames, variableName);
                    usedVariableNames.Add(variableName);

                    // {type} {variableName} = {value};
                    var field =
                        SyntaxFactory.FieldDeclaration(
                                SyntaxFactory.VariableDeclaration(
                                        SyntaxFactory.ParseTypeName(GetTypeString(vertex)))
                                    .AddVariables(SyntaxFactory
                                        .VariableDeclarator(variableName)
                                        .WithInitializer(SyntaxFactory
                                            .EqualsValueClause(SyntaxFactory
                                                .ParseExpression(FormatVertexValue(vertex.CellType, vertex.Value))))))
                            .AddModifiers(SyntaxFactory.Token(SyntaxKind.PrivateKeyword));
                    newClass = newClass.AddMembers(field);
                }

                var statements = new List<StatementSyntax>();
                if (generatedClass.IsSharedClass)
                {

                }
                else
                {
                    foreach (var vertex in formulas)
                    {

                    }

                    // public static {type} Calculate()
                    var outputField = generatedClass.OutputVertex;
                    var calculateMethod = SyntaxFactory
                        .MethodDeclaration(SyntaxFactory
                            .ParseTypeName(GetTypeString(outputField)), "Calculate")
                        .AddModifiers(SyntaxFactory.Token(SyntaxKind.PublicKeyword))
                        .AddModifiers(SyntaxFactory.Token(SyntaxKind.StaticKeyword))
                        .WithBody(SyntaxFactory.Block(statements));
                    newClass = newClass.AddMembers(calculateMethod);
                }


                @namespace = @namespace.AddMembers(newClass);
            }

            var code = @namespace
                .NormalizeWhitespace()
                .ToFullString();
            return code;
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
                    return "object";
                default:
                    return "";
            }
        }

        private string FormatVertexValue(CellType cellType, object vertexValue)
        {
            switch (cellType)
            {
                case CellType.Bool:
                    return vertexValue.ToString();
                case CellType.Text:
                    return $"\"{vertexValue}\"";
                case CellType.Number:
                    string value = vertexValue.ToString().Replace(",", ".");
                    if (value.Contains("%")) value = value.Replace("%", " * 0.01");
                    return value;
                case CellType.Date:
                    return $"DateTime.Parse({vertexValue});";
                case CellType.Unknown:
                default:
                    return "null";
            }
        }
    }
}
