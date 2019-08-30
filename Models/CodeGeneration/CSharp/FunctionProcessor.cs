﻿using System.Collections.Generic;
using System.Linq;
using Irony.Parsing;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Thesis.Models.VertexTypes;
using static Microsoft.CodeAnalysis.CSharp.SyntaxFactory;

namespace Thesis.Models.CodeGeneration.CSharp
{
    public partial class CSharpGenerator
    {
        private readonly Dictionary<string, CellType> _functionToCellTypeDictionary = new Dictionary<string, CellType>
        {
            {"+", CellType.Unknown},
            {"-", CellType.Number},
            {"*", CellType.Number},
            {"/", CellType.Number},
            {"%", CellType.Number},
            {"^", CellType.Number},
            {"ROUND", CellType.Number},
            {"ROUNDUP", CellType.Number},
            {"ROUNDDOWN", CellType.Number},
            {"SUM", CellType.Number},
            {"MIN", CellType.Number},
            {"MAX", CellType.Number},
            {"COUNT", CellType.Number},
            {"AVERAGE", CellType.Number},

            {"HLOOKUP", CellType.Unknown},
            {"VLOOKUP", CellType.Unknown},
            {"CHOOSE", CellType.Unknown},
            {"MATCH", CellType.Number},
            {"INDEX", CellType.Unknown},

            {"IF", CellType.Unknown},
            {"AND", CellType.Bool},
            {"NOT", CellType.Bool},
            {"OR", CellType.Bool},
            {"XOR", CellType.Bool},
            {"TRUE", CellType.Bool},
            {"FALSE", CellType.Bool},

            {"ISBLANK", CellType.Bool},
            {"ISLOGICAL", CellType.Bool},
            {"ISNOTEXT", CellType.Bool},
            {"ISNUMBER", CellType.Bool},
            {"ISTEXT", CellType.Bool},
            {"ISERR", CellType.Unknown},
            {"ISERROR", CellType.Unknown},
            {"ISNA", CellType.Unknown},

            {"=", CellType.Bool},
            {"<>", CellType.Bool},
            {"<", CellType.Number},
            {"<=", CellType.Number},
            {">=", CellType.Number},
            {">", CellType.Number},

            {"&", CellType.Text},
            {"CONCATENATE", CellType.Text},

            {"DATE", CellType.Date},
            {"SECOND", CellType.Date},
            {"MINUTE", CellType.Date},
            {"HOUR", CellType.Date},
            {"DAY", CellType.Date},
            {"MONTH", CellType.Date},
            {"TODAY", CellType.Date},
        };

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

        private ExpressionSyntax FunctionToExpression(string functionName, ParseTreeNode[] arguments,
            CellVertex currentVertex)
        {
            switch (functionName)
            {
                // arithmetic
                case "+":
                case "-":
                {
                    if (arguments.Length == 1)
                        return PrefixUnaryExpression(SyntaxKind.UnaryMinusExpression,
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
                            return InvocationExpression(MemberAccessExpression(
                                    SyntaxKind.SimpleMemberAccessExpression,
                                    leftExpr, IdentifierName("AddDays")))
                                .AddArgumentListArguments(Argument(rightExpr));
                        }

                        if (typeRight.Value == CellType.Date)
                        {
                            return InvocationExpression(MemberAccessExpression(
                                    SyntaxKind.SimpleMemberAccessExpression,
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
                        LiteralExpression(SyntaxKind.NumericLiteralExpression,
                            Literal(0.01)));
                }
                case "^":
                {
                    if (arguments.Length != 2) return FunctionError(functionName, arguments);
                    return ClassFunctionCall("Math", "Pow",
                        TreeNodeToExpression(arguments[0], currentVertex),
                        TreeNodeToExpression(arguments[1], currentVertex));
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
                    return InvocationExpression(MemberAccessExpression(
                        SyntaxKind.SimpleMemberAccessExpression,
                        CollectionOf(arguments.Select<ParseTreeNode, ExpressionSyntax>(a => TreeNodeToExpression(a, currentVertex)).ToArray()),
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
                        .Select<ParseTreeNode, ExpressionSyntax>(a => TreeNodeToExpression(a, currentVertex))
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
                        .WithArgumentList(
                            BracketedArgumentList(
                                SeparatedList<ArgumentSyntax>(argumentList.ToArray())));
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
                                Argument(
                                    LiteralExpression(SyntaxKind.TrueLiteralExpression)));
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

                    return ParenthesizedExpression(
                        ConditionalExpression(condition, whenTrue, whenFalse));
                }
                case "AND":
                case "OR":
                case "XOR":
                {
                    if (arguments.Length == 0) return FunctionError(functionName, arguments);
                    if (arguments.Length == 1) return TreeNodeToExpression(arguments[0], currentVertex);
                    return ParenthesizedExpression(FoldBinaryExpression(functionName, arguments,
                        currentVertex));
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
                    var rightExpression = PredefinedType(Token(SyntaxKind.BoolKeyword));
                    return GenerateBinaryExpression(functionName, leftExpression, rightExpression);
                }
                case "ISNONTEXT":
                {
                    // !(argument[0] is string)
                    return PrefixUnaryExpression(
                        SyntaxKind.LogicalNotExpression,
                        ParenthesizedExpression(FunctionToExpression("ISTEXT", arguments,
                            currentVertex)));
                }
                case "ISNUMBER":
                {
                    return InvocationExpression(IdentifierName("IsNumeric"))
                        .AddArgumentListArguments(
                            Argument(TreeNodeToExpression(arguments[0], currentVertex)));
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
                            .AddArgumentListArguments(Argument(leftExpression),
                                Argument(rightExpression));
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
                                             $"{string.Join("\n", arguments.Select<ParseTreeNode, ExpressionSyntax>(a => TreeNodeToExpression(a, currentVertex)))}",
                        true);
                }
            }
        }

        private ExpressionSyntax RoundFunction(string roundFunction, ParseTreeNode[] arguments,
            CellVertex currentVertex)
        {
            if (arguments.Length != 2) return FunctionError(roundFunction, arguments);
            return InvocationExpression(IdentifierName(roundFunction))
                .AddArgumentListArguments(
                    Argument(TreeNodeToExpression(arguments[0], currentVertex)),
                    Argument(TreeNodeToExpression(arguments[1], currentVertex)));
        }

        private ExpressionSyntax CollectionOf(params ExpressionSyntax[] expressions)
        {
            // avoid Collection(Collection)
            if (expressions.Length == 1 && expressions[0] is InvocationExpressionSyntax inv
                                        && inv.Expression is MemberAccessExpressionSyntax maes
                                        && ((IdentifierNameSyntax) maes.Expression).Identifier.Text == "Collection")
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

        private ExpressionSyntax ClassFunctionCall(string className, string functionName,
            params ExpressionSyntax[] expressions)
        {
            return InvocationExpression(
                    MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                        IdentifierName(className),
                        IdentifierName(functionName)))
                .AddArgumentListArguments(expressions.Select(Argument).ToArray());
        }

        private ExpressionSyntax GenerateBinaryExpression(string functionName, ParseTreeNode[] arguments,
            CellVertex vertex)
        {
            if (arguments.Length != 2) return FunctionError(functionName, arguments);
            return GenerateBinaryExpression(functionName,
                (ExpressionSyntax) TreeNodeToExpression(arguments[0], vertex),
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

        private ExpressionSyntax FoldBinaryExpression(string functionName, ParseTreeNode[] arguments, CellVertex vertex)
        {
            var syntaxKind = _binaryOperators[functionName].syntaxKind;
            // do not parenthesize
            return arguments.Select<ParseTreeNode, ExpressionSyntax>(a => TreeNodeToExpression(a, vertex))
                .Aggregate((acc, right) => BinaryExpression(syntaxKind, acc, right));
        }

        private ExpressionSyntax FunctionError(string functionName, ParseTreeNode[] arguments)
        {
            return CommentExpression($"Function {functionName} has incorrect number of arguments ({arguments.Length})", true);
        }
    }
}