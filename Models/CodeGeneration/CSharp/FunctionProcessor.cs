using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Thesis.Models.FunctionGeneration;
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
            {"ISNONTEXT", CellType.Bool},
            {"ISNUMBER", CellType.Bool},
            {"ISTEXT", CellType.Bool},
            {"ISERR", CellType.Unknown},
            {"ISERROR", CellType.Unknown},
            {"ISNA", CellType.Unknown},

            {"=", CellType.Bool},
            {"<>", CellType.Bool},
            {"<", CellType.Bool},
            {"<=", CellType.Bool},
            {">=", CellType.Bool},
            {">", CellType.Bool},

            {"&", CellType.Text},
            {"CONCATENATE", CellType.Text},

            {"DATE", CellType.Date},
            {"SECOND", CellType.Date},
            {"MINUTE", CellType.Date},
            {"HOUR", CellType.Date},
            {"DAY", CellType.Date},
            {"MONTH", CellType.Date},
            {"YEAR", CellType.Date},
            {"TODAY", CellType.Date},
        };

        private ExpressionSyntax ExpressionToCode(Expression expression)
        {
            switch (expression)
            {
                case Function function:
                    return FunctionToCode(function);
                case Reference reference:
                    return IdentifierName(reference.VariableName);
                case Constant constant:
                    var constantExpression = ConstantToExpression(constant);
                    if (_forceConstantsIntoDecimal && constant.ConstantValue is int && _forceConstantsIntoDecimal)
                        return CastExpression(PredefinedType( Token(SyntaxKind.DoubleKeyword)), constantExpression);
                    return constantExpression;
            }
            throw new Exception("Unidentified expression");
        }


        private ExpressionSyntax FunctionToCode(Function function)
        {
            return FunctionToCode(function.FunctionName, function.Arguments);
        }

        private ExpressionSyntax FunctionToCode(string functionName, Expression[] arguments)
        {
            switch (functionName)
            {
                // arithmetic
                case "+":
                case "-":
                    {
                        if (arguments.Length == 1)
                            return PrefixUnaryExpression(SyntaxKind.UnaryMinusExpression,
                                ExpressionToCode(arguments[0]));

                        if (arguments.Length != 2) return FunctionError(functionName, arguments);
                        var leftExpr = ExpressionToCode(arguments[0]);
                        var rightExpr = ExpressionToCode(arguments[1]);

                        // Date operations
                        var typeLeft = arguments[0].GetCellType();
                        var typeRight = arguments[1].GetCellType();
                        if (typeLeft == CellType.Date || typeRight == CellType.Date)
                        {
                            return GenerateDateArithmeticExpression(functionName, typeLeft, typeRight, leftExpr, rightExpr);
                        }

                        return GenerateBinaryExpression(functionName, leftExpr, rightExpr);
                    }

                case "*":
                    {
                        return GenerateBinaryExpression(functionName, arguments, true);
                    }
                case "/":
                    {
                        _forceConstantsIntoDecimal = true;
                        var expression = GenerateBinaryExpression(functionName, arguments, true);
                        _forceConstantsIntoDecimal = false;
                        return expression;
                    }
                case "%":
                    {
                        if (arguments.Length != 1) return FunctionError(functionName, arguments);
                        // arguments[0] / 100.0
                        return BinaryExpression(
                            SyntaxKind.DivideExpression,
                            ExpressionToCode(arguments[0]),
                            ParseExpression("100d"));
                    }
                case "^":
                    {
                        if (arguments.Length != 2) return FunctionError(functionName, arguments);
                        return ClassFunctionCall("Math", "Pow",
                            ExpressionToCode(arguments[0]),
                            ExpressionToCode(arguments[1]));
                    }
                case "ROUND":
                    {
                        return RoundFunction("Round", arguments);
                    }
                case "ROUNDUP":
                    {
                        return RoundFunction("RoundUp", arguments);
                    }
                case "ROUNDDOWN":
                    {
                        return RoundFunction("RoundDown", arguments);
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
                            CollectionOf(arguments.Select(ExpressionToCode).ToArray()),
                            IdentifierName(functionName.ToTitleCase())));
                    }

                // reference functions
                case "VLOOKUP":
                case "HLOOKUP":
                    {
                        if (arguments.Length != 3 && arguments.Length != 4)
                            return FunctionError(functionName, arguments);

                        var matrixName = RangeToVariableName(arguments[1]);

                        var lookupValue = ExpressionToCode(arguments[0]);
                        var columnIndex = ExpressionToCode(arguments[2]);

                        functionName = functionName == "VLOOKUP" ? "VLookUp" : "HLookUp";
                        var expression = InvocationExpression(
                                MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                                    IdentifierName(matrixName),
                                    IdentifierName(functionName)))
                            .AddArgumentListArguments(
                                Argument(lookupValue),
                                Argument(columnIndex));

                        if (arguments.Length == 4)
                            expression = expression.AddArgumentListArguments(
                                Argument(ExpressionToCode(arguments[3])));
                        return expression;
                    }
                case "CHOOSE":
                    {
                        if (arguments.Length < 2)
                            return FunctionError(functionName, arguments);
                        var collection = CollectionOf(arguments
                            .Skip(1)
                            .Select(ExpressionToCode)
                            .ToArray());

                        // collection[arguments[0]]
                        return ElementAccessExpression(collection)
                            .WithArgumentList(BracketedArgumentList(SingletonSeparatedList(
                                Argument(ExpressionToCode(arguments[0])))));
                    }
                case "INDEX":
                    {
                        if (arguments.Length != 2 && arguments.Length != 3)
                            return FunctionError(functionName, arguments);
                        var matrixOrCollection = RangeToVariableName(arguments[0]);

                        var argumentList = new List<SyntaxNodeOrToken>();
                        argumentList.Add(Argument(ExpressionToCode(arguments[1])));
                        if (arguments.Length == 3)
                        {
                            argumentList.Add(Token(SyntaxKind.CommaToken));
                            argumentList.Add(Argument(ExpressionToCode(arguments[2])));
                        }

                        return ElementAccessExpression(IdentifierName(matrixOrCollection))
                            .WithArgumentList(
                                BracketedArgumentList(
                                    SeparatedList<ArgumentSyntax>(argumentList.ToArray())));
                    }
                case "MATCH":
                    {
                        if (arguments.Length != 2 && arguments.Length != 3)
                            return FunctionError(functionName, arguments);
                        var collection = RangeToVariableName(arguments[1]);

                        var expression = InvocationExpression(
                                MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                                    IdentifierName(collection),
                                    IdentifierName("Match")))
                            .AddArgumentListArguments(
                                Argument(ExpressionToCode(arguments[0])));
                        if (arguments.Length == 3)
                            expression = expression.AddArgumentListArguments(
                                Argument(ExpressionToCode(arguments[2])));

                        return expression;
                    }

                // logical functions
                case "IF":
                    {
                        if (arguments.Length != 2 && arguments.Length != 3)
                            return FunctionError(functionName, arguments);

                        ExpressionSyntax condition = ExpressionToCode(arguments[0]);
                        ExpressionSyntax whenTrue = ExpressionToCode(arguments[1]);
                        ExpressionSyntax whenFalse;

                        // if the condition is not always a bool (e.g. dynamic), use Compare(cond, true)
                        // otherwise we might have a number as cond, and number can not be evaluated as bool
                        if (arguments[0].GetCellType() != CellType.Bool)
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
                            whenFalse = ExpressionToCode(arguments[2]);
                            argumentsHaveDifferentTypes = !arguments[1].IsSameTypeAndNotUnknown(arguments[2]);
                        }
                        else
                        {
                            // if no else statement is given, Excel defaults to FALSE
                            whenFalse = LiteralExpression(SyntaxKind.FalseLiteralExpression);
                            argumentsHaveDifferentTypes = arguments[1].GetCellType() != CellType.Bool;
                        }

                        // if there is a mismatch in argument types, the variable must be of type dynamic
                        if (argumentsHaveDifferentTypes)
                        {
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
                        if (arguments.Length == 1) return ExpressionToCode(arguments[0]);
                        return ParenthesizedExpression(FoldBinaryExpression(functionName, arguments));
                    }
                case "NOT":
                    {
                        if (arguments.Length != 1) return FunctionError(functionName, arguments);
                        return PrefixUnaryExpression(SyntaxKind.LogicalNotExpression,
                            ExpressionToCode(arguments[0]));
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
                        var leftExpression = ExpressionToCode(arguments[0]);
                        var rightExpression = IdentifierName("EmptyCell");
                        return GenerateBinaryExpression(functionName, leftExpression, rightExpression);
                    }
                case "ISLOGICAL":
                    {
                        if (arguments.Length != 1) return FunctionError(functionName, arguments);
                        var leftExpression = ExpressionToCode(arguments[0]);
                        var rightExpression = PredefinedType(Token(SyntaxKind.BoolKeyword));
                        return GenerateBinaryExpression(functionName, leftExpression, rightExpression);
                    }
                case "ISNONTEXT":
                    {
                        // !(argument[0] is string)
                        return PrefixUnaryExpression(
                            SyntaxKind.LogicalNotExpression,
                            ParenthesizedExpression(FunctionToCode("ISTEXT", arguments)));
                    }
                case "ISNUMBER":
                    {
                        return InvocationExpression(IdentifierName("IsNumeric"))
                            .AddArgumentListArguments(
                                Argument(ExpressionToCode(arguments[0])));
                    }
                case "ISTEXT":
                    {
                        if (arguments.Length != 1) return FunctionError(functionName, arguments);
                        var leftExpression = ExpressionToCode(arguments[0]);
                        var rightExpression = PredefinedType(Token(SyntaxKind.StringKeyword));
                        return GenerateBinaryExpression(functionName, leftExpression, rightExpression);
                    }
                case "ISERR":
                case "ISERROR":
                case "ISNA":
                    {
                        if (arguments.Length != 1) return FunctionError(functionName, arguments);
                        var leftExpression = ExpressionToCode(arguments[0]);
                        var rightExpression = IdentifierName("FormulaError");
                        return GenerateBinaryExpression(functionName, leftExpression, rightExpression);
                    }

                // comparators
                case "=":
                    {
                        if (arguments.Length != 2) return FunctionError(functionName, arguments);
                        var leftExpression = ExpressionToCode(arguments[0]);
                        var rightExpression = ExpressionToCode(arguments[1]);

                        ExpressionSyntax equalsExpression;
                        var leftType = arguments[0].GetCellType();
                        var rightType = arguments[1].GetCellType();
                        if (leftType == rightType)
                        {
                            // Excel uses case insensitive string compare
                            if (leftType == CellType.Text && rightType == CellType.Text)
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
                        var leftType = arguments[0].GetCellType();
                        var rightType = arguments[1].GetCellType();
                        if (leftType == rightType &&
                            (leftType != CellType.Text || rightType != CellType.Text))
                        {
                            return GenerateBinaryExpression(functionName, arguments);
                        }

                        return PrefixUnaryExpression(SyntaxKind.LogicalNotExpression,
                            FunctionToCode("=", arguments));
                    }
                case "<":
                case "<=":
                case ">=":
                case ">":
                    {
                        return GenerateBinaryExpression(functionName, arguments);
                    }

                // strings
                case "&":
                case "CONCATENATE":
                    {
                        if (arguments.Length < 2) return FunctionError(functionName, arguments);
                        return FoldBinaryExpression("+", arguments);
                    }

                // other
                case "DATE":
                    {
                        if (arguments.Length != 3) return FunctionError(functionName, arguments);
                        return ObjectCreationExpression(
                                IdentifierName("DateTime"))
                            .AddArgumentListArguments(
                                Argument(ExpressionToCode(arguments[0])),
                                Argument(ExpressionToCode(arguments[1])),
                                Argument(ExpressionToCode(arguments[2]))
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
                            ExpressionToCode(arguments[0]),
                            IdentifierName(functionName.ToTitleCase()));
                    }

                default:
                    {
                        return CommentExpression($"Function {functionName} not implemented yet!", true);
                    }
            }
        }

        private ExpressionSyntax GenerateDateArithmeticExpression(string functionName, CellType typeLeft, CellType typeRight,
            ExpressionSyntax leftExpr, ExpressionSyntax rightExpr)
        {
            if (typeLeft == CellType.Date && typeRight == CellType.Date)
            {
                // (left + right).Days
                return MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression, ParenthesizedExpression(BinaryExpression(
                            functionName == "+" ? SyntaxKind.AddExpression : SyntaxKind.SubtractExpression,
                            leftExpr,
                            rightExpr)),
                    IdentifierName("Days"));
            }

            var rightArgument = typeLeft == CellType.Date ? rightExpr : leftExpr;
            if (functionName == "-")
                rightArgument = PrefixUnaryExpression(SyntaxKind.UnaryMinusExpression, rightArgument);
            return InvocationExpression(MemberAccessExpression(
                    SyntaxKind.SimpleMemberAccessExpression,
                    typeLeft == CellType.Date ? leftExpr : rightExpr, IdentifierName("AddDays")))
                .AddArgumentListArguments(Argument(rightArgument));
        }

        private ExpressionSyntax RoundFunction(string roundFunction, Expression[] arguments)
        {
            if (arguments.Length != 2) return FunctionError(roundFunction, arguments);
            return InvocationExpression(IdentifierName(roundFunction))
                .AddArgumentListArguments(
                    Argument(ExpressionToCode(arguments[0])),
                    Argument(ExpressionToCode(arguments[1])));
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

        private ExpressionSyntax GenerateBinaryExpression(string functionName, Expression[] arguments, bool parenthesize = false)
        {
            if (arguments.Length != 2) return FunctionError(functionName, arguments);

            var leftExpression = ExpressionToCode(arguments[0]);
            var rightExpression = ExpressionToCode(arguments[1]);

            //TODO better parenthesization 
            if (parenthesize && arguments[0] is Function)
                leftExpression = ParenthesizedExpression(leftExpression);
            if (parenthesize && arguments[1] is Function)
                rightExpression = ParenthesizedExpression(rightExpression);

            return GenerateBinaryExpression(functionName,
                leftExpression,
                rightExpression);
        }

        private readonly SortedList<string, SyntaxKind> _binaryOperators
            = new SortedList<string, SyntaxKind>
            {
                {"+", SyntaxKind.AddExpression},
                {"-", SyntaxKind.SubtractExpression},
                {"/", SyntaxKind.DivideExpression},
                {"*", SyntaxKind.MultiplyExpression},

                {"ISBLANK", SyntaxKind.IsExpression},
                {"ISLOGICAL", SyntaxKind.IsExpression},
                {"ISNOTEXT", SyntaxKind.IsExpression},
                {"ISNUMBER", SyntaxKind.IsExpression},
                {"ISTEXT", SyntaxKind.IsExpression},
                {"ISERR", SyntaxKind.IsExpression},
                {"ISERROR", SyntaxKind.IsExpression},
                {"ISNA", SyntaxKind.IsExpression},

                {"<>", SyntaxKind.NotEqualsExpression},
                {"<", SyntaxKind.LessThanExpression},
                {"<=", SyntaxKind.LessThanOrEqualExpression},
                {">=", SyntaxKind.GreaterThanOrEqualExpression},
                {">", SyntaxKind.GreaterThanExpression},
                {"AND", SyntaxKind.LogicalAndExpression},
                {"OR", SyntaxKind.LogicalOrExpression},
                {"XOR", SyntaxKind.ExclusiveOrExpression},
            };

        private ExpressionSyntax GenerateBinaryExpression(string functionName,
            ExpressionSyntax leftExpression, ExpressionSyntax rightExpression)
        {
            SyntaxKind syntaxKind = _binaryOperators[functionName];
            return BinaryExpression(syntaxKind, leftExpression, rightExpression);
        }

        private ExpressionSyntax FoldBinaryExpression(string functionName, Expression[] arguments)
        {
            var syntaxKind = _binaryOperators[functionName];
            // do not parenthesize
            return arguments.Select(ExpressionToCode)
                .Aggregate((acc, right) => BinaryExpression(syntaxKind, acc, right));
        }

        private ExpressionSyntax FunctionError(string functionName, Expression[] arguments)
        {
            return CommentExpression($"Function {functionName} has incorrect number of arguments ({arguments.Length})", true);
        }
    }
}