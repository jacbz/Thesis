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

using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Thesis.Models.VertexTypes;

namespace Thesis.Models.CodeGeneration
{
    public abstract class Tester
    {
        public List<ClassCode> ClassesCode { get; }
        public TestResults TestResults { get; set; }

        protected Tester(List<ClassCode> classesCode)
        {
            ClassesCode = classesCode;
        }

        public abstract Task PerformTestAsync();
        public abstract TestReport GenerateTestReport(Dictionary<(string className, string variableName), Vertex> variableNameToVertexDictionary);

        public static bool IsNumeric(object o) => o is byte || o is sbyte || o is ushort || o is uint || o is ulong || o is short || o is int || o is long || o is float || o is double || o is decimal;
    }

    public class TestReport
    {
        public string Code { get; set; }
        public int PassCount { get; }
        public int NullCount { get; }
        public int ValueMismatchCount { get; }
        public int TypeMismatchCount { get; }
        public int ErrorCount { get; }

        public TestReport(int passCount, int nullCount, int valueMismatchCount, int typeMismatchCount, int errorCount)
        {
            PassCount = passCount;
            NullCount = nullCount;
            ValueMismatchCount = valueMismatchCount;
            TypeMismatchCount = typeMismatchCount;
            ErrorCount = errorCount;
        }

        public override string ToString()
        {
            return $"Test results: {PassCount} Pass, {NullCount} were null, {ValueMismatchCount} Value Mismatch, {TypeMismatchCount} Type Mismatch, {ErrorCount} Error";
        }
    }

    public enum TestResultType
    {
        Pass,
        Null,
        TypeMismatch,
        ValueMismatch,
        Error,
        Ignore
    }

    public class TestResults
    {
        private readonly Dictionary<(string className, string variableName), TestResult> _testResults;

        public TestResults()
        {
            _testResults = new Dictionary<(string className, string variableName), TestResult>();
        }

        public void Add((string className, string variableName) tuple, TestResult testResult)
        {
            _testResults.Add(tuple, testResult);
        }

        public TestResult this[(string className, string variableName) tuple] =>
            _testResults.TryGetValue(tuple, out var testResult)
                ? testResult
                : null;

        public TestResult this[Vertex vertex]
        {
            get
            {
                var tuple = (vertex.Class.Name, vertex.Name);
                return _testResults.TryGetValue(tuple, out var testResult) 
                    ? testResult 
                    : null;
            }
        }
    }
    

    public class TestResult
    {
        public string ClassName { get; }
        public string VariableName { get; }
        public dynamic ActualValue { get; }
        public dynamic ExpectedValue { get; set; }
        public TestResultType TestResultType { get; set; }
        public CellType ExpectedValueType { get; set; }
        public Type ActualValueType { get; }
        public string Annotation { get; set; }

        public TestResult(string className, string variableName, dynamic actualValue, Type actualValueType)
        {
            ClassName = className;
            VariableName = variableName;
            ActualValue = actualValue;
            ActualValueType = actualValueType;
        }

        public void SetExpectedValue(CellVertex vertex)
        {
            if (vertex == null)
            {
                ExpectedValue = null;
                ExpectedValueType = CellType.Unknown;
            }
            else
            {
                ExpectedValue = vertex.Value;
                ExpectedValueType = vertex.CellType;
            }
        }

        public override string ToString()
        {
            return TestResultTypeToString() + (string.IsNullOrEmpty(Annotation) ? "" : $" ({Annotation})");
        }

        private string TestResultTypeToString()
        {
            switch (TestResultType)
            {
                case TestResultType.Pass:
                    return "//  PASS  Value: " + ActualValue;
                case TestResultType.TypeMismatch:
                    return $"//  FAIL  Expected: {ExpectedValue} ({ExpectedValue.GetType().Name}), " +
                           $"Actual: {ActualValue} ({ActualValue.GetType().Name})";
                case TestResultType.ValueMismatch:
                    return $"//  FAIL  Expected: {ExpectedValue}, Actual: {ActualValue}";
                case TestResultType.Error:
                    return $"//  COMPILE/RUNTIME ERROR  Expected: {ExpectedValue}";
                case TestResultType.Ignore:
                    return "//  IGNORE  Can not check this type for correctness";
                default:
                    return "// Error";
            }
        }
    }

    public class ClassCode
    {
        public bool IsStaticClass { get; }
        public string ClassName { get; }
        public string Code { get; }
        public string FieldsCode { get; }
        public string MethodBodyCode { get; }

        public ClassCode(bool isStaticClass, string className, string code, string fieldsCode, string methodBodyCode)
        {
            IsStaticClass = isStaticClass;
            ClassName = className;
            Code = code;
            FieldsCode = fieldsCode;
            MethodBodyCode = methodBodyCode;
        }
    }
}
