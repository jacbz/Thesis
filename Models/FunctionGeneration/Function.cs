// Thesis - An Excel to code converter
// Copyright (C) 2019 Jacob Zhang
// 
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
// 
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
// 
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <https://www.gnu.org/licenses/>.

using System.Collections.Generic;
using System.Linq;
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
        }

        public abstract bool IsStructurallyEquivalentTo(CellFunction otherFunction);
    }

    public class FormulaFunction : CellFunction
    {
        public Expression Expression { get; }

        public FormulaFunction(CellVertex formulaVertex, Expression expression) : base(formulaVertex)
        {
            Name.IsFunction = true;
            Expression = expression;
            Parameters = expression.GetLeafsOfType<InputReference>().ToArray();
            ReturnType = expression.GetCellType();
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
            Name.IsOutputField = true;
            Statements = statements;

            var statementVariableNames = Statements.Select(statement => statement.VariableName.ToString()).ToHashSet();
            Parameters = Statements
                .OfType<FunctionInvocationStatement>()
                .SelectMany(statement => statement.Parameters)
                .Where(inputReference => !statementVariableNames.Contains(inputReference.VariableName.ToString()))
                .ToArray();

            ReturnType = statements.Last().VariableType;
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

        public bool IsSameTypeAndNotUnknown(Expression otherExpression)
        {
            var cellType = GetCellType();
            return cellType == otherExpression.GetCellType() && cellType != CellType.Unknown;
        }
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

    public abstract class GlobalReference : Reference
    {
        public Vertex ReferencedVertex { get; }

        protected GlobalReference(Vertex cellVertex) : base(cellVertex.Name)
        {
            ReferencedVertex = cellVertex;
        }

        public override bool IsStructurallyEquivalentTo(Expression otherExpression)
        {
            return otherExpression is GlobalReference otherGlobalReference &&
                   ReferencedVertex == otherGlobalReference.ReferencedVertex;
        }
    }

    public class GlobalCellReference : GlobalReference
    {
        public new CellVertex ReferencedVertex { get; }

        public GlobalCellReference(CellVertex cellVertex) : base(cellVertex)
        {
            ReferencedVertex = cellVertex;
        }

        public override CellType GetCellType()
        {
            return ReferencedVertex.CellType;
        }
    }

    public class GlobalRangeReference : GlobalReference
    {
        public new RangeVertex ReferencedVertex { get; }

        public GlobalRangeReference(RangeVertex rangeVertex) : base(rangeVertex)
        {
            ReferencedVertex = rangeVertex;
        }

        public override CellType GetCellType()
        {
            return CellType.Unknown;
        }
    }

    public class InputReference : Reference
    {
        public CellType InputType { get; }


        public InputReference(CellVertex cellVertex) : base(cellVertex.Name)
        {
            InputType = cellVertex.CellType;
        }

        protected InputReference(Name name, CellType inputType) : base(name)
        {
            InputType = inputType;
        }

        public InputReference Copy()
        {
            return new InputReference(VariableName.Copy(), InputType);
        }

        public override CellType GetCellType()
        {
            return InputType;
        }

        public override bool IsStructurallyEquivalentTo(Expression otherExpression)
        {
            return otherExpression is InputReference otherInputCellReference &&
                   InputType == otherInputCellReference.InputType;
        }
    }

    public class Constant : ReferenceOrConstant
    { 
        public dynamic ConstantValue { get; }
        public CellType ConstantType { get; }

        public Constant(string numberString, bool isPercentage = false)
        {
            ConstantType = CellType.Number;
            if (isPercentage)
                numberString = numberString.Replace("%", "").Replace("\"", "");
            if (int.TryParse(numberString, out int result))
                ConstantValue = result;
            ConstantValue = double.Parse(numberString);
            if (isPercentage)
                ConstantValue /= 100.0;
        }

        public Constant(dynamic value, CellType type)
        {
            ConstantValue = value;
            ConstantType = type;

            if (value is string s && s[0] == '"' && s[s.Length - 1] == '"')
                ConstantValue = s.Substring(1, s.Length - 2);
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
            {"TODAY", CellType.Date},
            {"SECOND", CellType.Number},
            {"MINUTE", CellType.Number},
            {"HOUR", CellType.Number},
            {"DAY", CellType.Number},
            {"MONTH", CellType.Number},
            {"YEAR", CellType.Number},
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
        public CellType VariableType { get; protected set; }
        protected Statement(Name variableName, CellType variableType)
        {
            VariableName = variableName;
            VariableType = variableType;
        }
        public abstract bool IsStructurallyEquivalentTo(Statement otherStatement);
    }

    //public class ConstantStatement : Statement
    //{
    //    // for Constant, store entire expression
    //    public Constant Constant { get; }

    //    public ConstantStatement(Name variableName, CellType variableType, Constant constant) : base(variableName, variableType)
    //    {
    //        Constant = constant;
    //    }

    //    public override bool IsStructurallyEquivalentTo(Statement otherStatement)
    //    {
    //        return otherStatement is ConstantStatement otherConstantStatement &&
    //              Constant.IsStructurallyEquivalentTo(otherConstantStatement.Constant);
    //    }
    //}

    public class FunctionInvocationStatement : Statement
    {
        // for Functions, only store the names of the parameters which will be used for invocation
        public Name FunctionName { get; }
        public InputReference[] Parameters;

        public FunctionInvocationStatement(Name variableName, CellType variableType, FormulaFunction function) : base(variableName, variableType)
        {
            if (function.ReturnType == CellType.Unknown)
                VariableType = CellType.Unknown;
            FunctionName = function.Name;
            Parameters = function.Parameters.ToArray();
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
