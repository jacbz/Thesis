using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Thesis.Models.VertexTypes;

namespace Thesis.Models.FunctionGeneration
{
    public abstract class CellFunction
    {
        public Name Name { get; set; }
        public CellType ReturnType { get; protected set; }
        public InputReference[] Parameters { get; protected set; }

        protected CellFunction(CellVertex vertex)
        {
            Name = vertex.Name.Copy();
            ReturnType = vertex.CellType;
        }

        public abstract bool IsStructurallyEquivalentTo(CellFunction otherFunction);
    }

    public class FormulaFunction : CellFunction
    {
        public Expression Expression { get; }

        public FormulaFunction(CellVertex formulaVertex, Expression expression) : base(formulaVertex)
        {
            Expression = expression;
            Parameters = expression.GetLeafsOfType<InputReference>().ToArray();

            // make parameter names unique
            Name.MakeNamesUnique(Parameters.Select(par => par.VariableName));
        }

        public override bool IsStructurallyEquivalentTo(CellFunction otherFunction)
        {
            if (otherFunction is FormulaFunction otherFormulaFunction)
                return Expression.IsStructurallyEquivalentTo(otherFormulaFunction.Expression);
            return false;
        }
    }

    public class OutputFieldFunction : CellFunction
    {
        public Statement[] Statements { get; }

        public OutputFieldFunction(CellVertex outputFieldVertex, Statement[] statements) : base(outputFieldVertex)
        {
            Statements = statements;
            Name.MakeNamesUnique(statements.Select(statement => statement.VariableName));

            var statementVariableNames = Statements.Select(statement => statement.VariableName).ToHashSet();
            Parameters = Statements
                .OfType<FunctionInvocationStatement>()
                .SelectMany(statement => statement.Parameters)
                .Where(inputReference => !statementVariableNames.Contains(inputReference.VariableName))
                .ToArray();
            Name.MakeNamesUnique(Parameters.Select(par => par.VariableName));
        }

        public override bool IsStructurallyEquivalentTo(CellFunction otherFunction)
        {
            if (!(otherFunction is OutputFieldFunction otherOutputFieldFunction) ||
                Statements.Length != otherOutputFieldFunction.Statements.Length)
                return false;
            for(int i = 0; i < Statements.Length; i++)
                if (!Statements[i].IsStructurallyEquivalentTo(otherOutputFieldFunction.Statements[i]))
                    return false;

            return true;
        }
    }


    public abstract class Expression
    {
        public abstract CellType GetCellType();
        public abstract T[] GetLeafsOfType<T>();
        public abstract bool IsStructurallyEquivalentTo(Expression otherExpression);
    }

    public abstract class ReferenceOrConstant : Expression
    {
        public override T[] GetLeafsOfType<T>()
        {
            if (this is T)
                return new T[] { (T)(object)this };
            return new T[0];
        }
    }

    public abstract class Reference : ReferenceOrConstant
    {
        public Name VariableName { get; protected set; }

        protected Reference(Name variableName)
        {
            VariableName = variableName;
        }
    }

    public class GlobalReference : Reference
    {
        public CellVertex ReferencedVertex { get; }

        public GlobalReference(CellVertex cellVertex) : base(cellVertex.Name)
        {
            ReferencedVertex = cellVertex;
        }

        public override CellType GetCellType()
        {
            return ReferencedVertex.CellType;
        }

        public override bool IsStructurallyEquivalentTo(Expression otherExpression)
        {
            return otherExpression is GlobalReference otherGlobalReference &&
                   ReferencedVertex == otherGlobalReference.ReferencedVertex;
        }
    }

    public abstract class InputReference : Reference
    {
        protected InputReference(Name variableName) : base(variableName)
        {
        }
    }

    public class InputCellReference : InputReference
    {
        public CellType InputType { get; }

        public InputCellReference(CellVertex cellVertex) : base(cellVertex.Name)
        {
            InputType = cellVertex.CellType;
        }

        public override CellType GetCellType()
        {
            return InputType;
        }

        public override bool IsStructurallyEquivalentTo(Expression otherExpression)
        {
            return otherExpression is InputCellReference otherInputCellReference &&
                   InputType == otherInputCellReference.InputType;
        }
    }

    public class InputRangeReference : InputReference
    {
        public int ColumnCount { get; }
        public int RowCount { get; }

        public InputRangeReference(RangeVertex rangeVertex) : base(rangeVertex.Name)
        {
            ColumnCount = rangeVertex.ColumnCount;
            RowCount = rangeVertex.RowCount;
        }

        public override CellType GetCellType()
        {
            return CellType.Unknown;
        }

        public override bool IsStructurallyEquivalentTo(Expression otherExpression)
        {
            return otherExpression is InputRangeReference otherInputRangeReference &&
                   ColumnCount == otherInputRangeReference.ColumnCount &&
                   RowCount == otherInputRangeReference.RowCount;
        }
    }

    public class Constant : ReferenceOrConstant
    { 
        public dynamic ConstantValue { get; }
        public CellType ConstantType { get; }

        public Constant(string numberString)
        {
            ConstantType = CellType.Number;
            if (int.TryParse(numberString, out int result))
                ConstantValue = result;
            ConstantValue = double.Parse(numberString);
        }

        public Constant(dynamic value, CellType type)
        {
            ConstantValue = value;
            ConstantType = type;
        }

        public override CellType GetCellType()
        {
            return ConstantType;
        }

        public override bool IsStructurallyEquivalentTo(Expression otherExpression)
        {
            if (otherExpression is Constant otherConstant)
                return ConstantType == otherConstant.ConstantType && ConstantValue == otherConstant.ConstantValue;
            return false;
        }
    }

    public class Function : Expression
    {
        public string FunctionName { get; protected set; }
        public Expression[] Arguments { get; protected set; }

        public static Dictionary<string, CellType> FunctionToCellTypeDictionary = new Dictionary<string, CellType>
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

        public Function(string name, Expression[] arguments)
        {
            FunctionName = name;
            Arguments = arguments;
        }

        public override CellType GetCellType()
        {
            if (FunctionName == "IF")
            {
                var ifType = Arguments[1].GetCellType();
                return (Arguments.Length == 3
                    ? ifType == Arguments[2].GetCellType()
                    : ifType == CellType.Bool) ? ifType : CellType.Unknown;
            }
            return FunctionToCellTypeDictionary[FunctionName];
        }

        public override T[] GetLeafsOfType<T>()
        {
            return Arguments.SelectMany(function => function.GetLeafsOfType<T>()).ToArray();
        }

        public override bool IsStructurallyEquivalentTo(Expression otherExpression)
        {
            if (!(otherExpression is Function otherFunction) ||
                Arguments.Length != otherFunction.Arguments.Length) return false;

            for(int i = 0; i < Arguments.Length; i++)
                if (!Arguments[i].IsStructurallyEquivalentTo(otherFunction.Arguments[i]))
                    return false;

            return true;
        }
    }

    public abstract class Statement
    {
        public Name VariableName { get; }
        public CellType VariableType { get; }
        protected Statement(Name variableName, CellType variableType)
        {
            VariableName = variableName;
            VariableType = variableType;
        }
        public abstract bool IsStructurallyEquivalentTo(Statement otherStatement);
    }

    public class ConstantStatement : Statement
    {
        // for Constant, store entire expression
        public Constant Constant { get; }

        public ConstantStatement(Name variableName, CellType variableType, Constant constant) : base(variableName, variableType)
        {
            Constant = constant;
        }

        public override bool IsStructurallyEquivalentTo(Statement otherStatement)
        {
            return otherStatement is ConstantStatement otherConstantStatement &&
                  Constant.IsStructurallyEquivalentTo(otherConstantStatement.Constant);
        }
    }

    public class FunctionInvocationStatement : Statement
    {
        private readonly FormulaFunction _formulaFunction;

        // for Functions, only store the names of the parameters which will be used for invocation
        public Name FunctionName => _formulaFunction.Name;
        public InputReference[] Parameters => _formulaFunction.Parameters;

        public FunctionInvocationStatement(Name variableName, CellType variableType, FormulaFunction function) : base(variableName, variableType)
        {
            _formulaFunction = function;
        }

        public string[] GetParameterNames()
        {
            return Parameters.Select(parameter => parameter.VariableName.ToString()).ToArray();
        }

        public override bool IsStructurallyEquivalentTo(Statement otherStatement)
        {
            return otherStatement is FunctionInvocationStatement otherFunctionInvocationStatement &&
                   FunctionName == otherFunctionInvocationStatement.FunctionName &&
                   Parameters.Length == otherFunctionInvocationStatement.Parameters.Length;
        }
    }
}
