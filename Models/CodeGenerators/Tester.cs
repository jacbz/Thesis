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

        public static bool IsNumeric(object o) => o is byte || o is sbyte || o is ushort || o is uint || o is ulong || o is short || o is int || o is long || o is float || o is double || o is decimal;

        public TestReport GenerateTestReport(Dictionary<string, Vertex> variableNameToVertexDictionary)
        {
            Logger.DispatcherLog(LogItemType.Info, "Generating test report...");

            int passCount = 0, valueMismatchCount = 0, typeMismatchCount = 0, skippedCount = 0;

            foreach (var keyValuePair in variableNameToVertexDictionary)
            {
                var variableName = keyValuePair.Key;
                var vertex = keyValuePair.Value;

                if (vertex.NodeType == NodeType.Constant) continue;

                if (VariableToTestResultDictionary.TryGetValue(variableName, out var testResult))
                {
                    testResult.ExpectedValue = vertex.Value;

                    if (testResult.ExpectedValue.GetType() != testResult.ActualValue.GetType() &&
                        // two different numeric types are to be treated as equal types
                        !(IsNumeric(testResult.ExpectedValue) && IsNumeric(testResult.ActualValue)))
                    {
                        // test for strings with %, e.g. 0,02 should pass with the expected value was "2%"
                        if (vertex.CellType == CellType.Number &&
                            testResult.ExpectedValue is string percentNumber &&
                            percentNumber.Contains("%") &&
                            double.TryParse(percentNumber.Replace("%", ""), out double number))
                        {
                            if (IsNumeric(testResult.ActualValue) && testResult.ActualValue == number * 0.01)
                            {
                                testResult.TestResultType = TestResultType.Pass;
                                passCount++;
                                continue;
                            }
                        }

                        testResult.TestResultType = TestResultType.TypeMismatch;
                        typeMismatchCount++;
                    }
                    else if (testResult.ExpectedValue != testResult.ActualValue)
                    {
                        testResult.TestResultType = TestResultType.ValueMismatch;
                        valueMismatchCount++;
                    }
                    else
                    {
                        testResult.TestResultType = TestResultType.Pass;
                        passCount++;
                    }
                }
                else
                {
                    VariableToTestResultDictionary.Add(variableName,
                        new TestResult(vertex.Class.Name, vertex.VariableName, null, null)
                    {
                        ExpectedValue = vertex.Value,
                        TestResultType = TestResultType.Skipped
                    });
                    skippedCount++;

                }
            }

            return new TestReport(passCount, valueMismatchCount, typeMismatchCount, skippedCount);
        }

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
        public bool IsSharedClass { get; }
        public string ClassName { get; }
        public string Code { get; }
        public string FieldsCode { get; }
        public string MethodBodyCode { get; }

        public ClassCode(bool isSharedClass, string className, string code, string fieldsCode, string methodBodyCode)
        {
            IsSharedClass = isSharedClass;
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

        public TestResult(string className, string variableName, dynamic actualValue, Type valueType)
        {
            ClassName = className;
            VariableName = variableName;
            ActualValue = actualValue;
            ValueType = valueType;
        }

        public override string ToString()
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
