using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Thesis.ViewModels;

namespace Thesis.Models.CodeGenerators
{
    public abstract class Tester
    {
        public List<ClassCode> ClassesCode { get; set; }
        public Dictionary<string, TestResult> VariableToTestResultDictionary { get; set; }
        public abstract Task PerformTestAsync();
        public abstract TestReport GenerateTestReport(Dictionary<string, Vertex> variableNameToVertexDictionary);

        public static bool IsNumeric(object o) => o is byte || o is sbyte || o is ushort || o is uint || o is ulong || o is short || o is int || o is long || o is float || o is double || o is decimal;


        protected Tester()
        {
            ClassesCode = new List<ClassCode>();
        }
    }

    public class TestReport
    {
        public string Code { get; set; }
        public int PassCount { get; }
        public int ValueMismatchCount { get; }
        public int TypeMismatchCount { get; }
        public int SkippedCount { get; }

        public TestReport(int passCount, int valueMismatchCount, int typeMismatchCount, int skippedCount)
        {
            PassCount = passCount;
            ValueMismatchCount = valueMismatchCount;
            TypeMismatchCount = typeMismatchCount;
            SkippedCount = skippedCount;
        }

        public override string ToString()
        {
            return $"Test results: {PassCount} Pass, {ValueMismatchCount} Value Mismatch, {TypeMismatchCount} Type Mismatch, {SkippedCount} Skipped";
        }
    }

    public enum TestResultType
    {
        Pass,
        TypeMismatch,
        ValueMismatch,
        Skipped
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

    public class TestResult
    {
        public string ClassName { get; }
        public string VariableName { get; }
        public dynamic ActualValue { get; }
        public dynamic ExpectedValue { get; set; }
        public TestResultType TestResultType { get; set; }
        public Type ValueType { get; }
        public string Annotation { get; set; }

        public TestResult(string className, string variableName, dynamic actualValue, Type valueType)
        {
            ClassName = className;
            VariableName = variableName;
            ActualValue = actualValue;
            ValueType = valueType;
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
                case TestResultType.Skipped:
                    return $"//  SKIPPED  Expected: {ExpectedValue}";
                default:
                    return "// Error";
            }
        }
    }
}
