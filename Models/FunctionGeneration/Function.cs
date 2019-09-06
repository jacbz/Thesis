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
        public string Name { get; protected set; }
        public CellType ReturnType { get; protected set; }
        public InputReference[] Parameters { get; protected set; }

        protected CellFunction(CellVertex vertex)
        {
            Name = vertex.Name;
            ReturnType = vertex.CellType;
        }
    }

    public class FormulaFunction : CellFunction
    {
        public Expression Expression { get; }

        public FormulaFunction(CellVertex formulaVertex, Expression expression) : base(formulaVertex)
        {
            Expression = expression;
            Parameters = expression.GetLeafsOfType<InputReference>().ToArray();
        }
    }

    public class OutputFieldFunction : CellFunction
    {
        public Statement[] Statements { get; }

        public OutputFieldFunction(CellVertex outputFieldVertex, Statement[] statements) : base(outputFieldVertex)
        {
            Name = "Calculate" + Name.RaiseFirstCharacter();
            Statements = statements;

            var statementVariableNames = Statements.Select(statement => statement.VariableName).ToHashSet();
            Parameters = Statements
                .OfType<FunctionInvocationStatement>()
                .SelectMany(statement => statement.Parameters)
                .Where(inputReference => !statementVariableNames.Contains(inputReference.VariableName))
                .ToArray();
        }
    }


    public abstract class Expression
    {
        public abstract CellType GetCellType();
        public abstract T[] GetLeafsOfType<T>();
        public abstract string Name { get; }
    }

    public abstract class ReferenceOrConstant : Expression
    {
        private new CellType[] ParameterTypes = null;

        public override T[] GetLeafsOfType<T>()
        {
            if (this is T)
                return new T[] { (T)(object)this };
            return new T[0];
        }
    }

    public abstract class Reference : ReferenceOrConstant
    {
        public string VariableName { get; protected set; }
        public override string Name => "Get" + VariableName;
    }

    public class GlobalReference : Reference
    {
        public CellVertex ReferencedVertex { get; }

        public GlobalReference(CellVertex cellVertex)
        {
            ReferencedVertex = cellVertex;
            VariableName = cellVertex.Name;
        }

        public override CellType GetCellType()
        {
            return ReferencedVertex.CellType;
        }
    }

    public abstract class InputReference : Reference
    {
    }

    public class InputCellReference : InputReference
    {
        public CellType InputType { get; }

        public InputCellReference(CellVertex cellVertex)
        {
            InputType = cellVertex.CellType;
            VariableName = cellVertex.Name;
        }

        public override CellType GetCellType()
        {
            return InputType;
        }
    }

    public class InputRangeReference : InputReference
    {
        public int ColumnCount { get; }
        public int RowCount { get; }

        public InputRangeReference(RangeVertex rangeVertex)
        {
            ColumnCount = rangeVertex.ColumnCount;
            RowCount = rangeVertex.RowCount;
            VariableName = rangeVertex.Name;
        }

        public override CellType GetCellType()
        {
            return CellType.Unknown;
        }
    }

    public class Constant : ReferenceOrConstant
    { 
        public dynamic ConstantValue { get; }
        public CellType ConstantType { get; }
        public override string Name => "const";

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
    }

    public class Function : Expression
    {
        public string FunctionName { get; protected set; }
        public override string Name => FunctionName;
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
    }

    public abstract class Statement
    {
        public string VariableName { get; }
        public CellType VariableType { get; }
        protected Statement(string variableName, CellType variableType)
        {
            VariableName = variableName;
            VariableType = variableType;
        }
    }

    public class ConstantStatement : Statement
    {
        // for Constant, store entire expression
        public Constant Constant { get; }

        public ConstantStatement(string variableName, CellType variableType, Constant constant) : base(variableName, variableType)
        {
            Constant = constant;
        }
    }

    public class FunctionInvocationStatement : Statement
    {
        // for Functions, only store the names of the parameters which will be used for invocation
        public string FunctionName { get; }
        public InputReference[] Parameters { get; }

        public FunctionInvocationStatement(string variableName, CellType variableType, FormulaFunction function) : base(variableName, variableType)
        {
            FunctionName = function.Name;
            Parameters = function.Parameters;
        }

        public string[] GetParameterNames()
        {
            return Parameters.Select(parameter => parameter.VariableName).ToArray();
        }
    }
}
