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

        public TestReport GenerateTestReport(Dictionary<string, Vertex> variableNameToVertexDictionary)
        {
            Logger.DispatcherLog(LogItemType.Info, "Generating test report...");

            int passCount = 0, valueMismatchCount = 0, typeMismatchCount = 0;
            foreach (var testResult in VariableToTestResultDictionary.Values)
            {
                var vertex = variableNameToVertexDictionary[testResult.VariableName];
                if (vertex.NodeType == NodeType.Constant) continue;

                testResult.ExpectedValue = vertex.Value;

                if (testResult.ExpectedValue.GetType() != testResult.ActualValue.GetType())
                {
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

            return new TestReport(passCount, valueMismatchCount, typeMismatchCount);
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

        public TestReport(int passCount, int valueMismatchCount, int typeMismatchCount)
        {
            PassCount = passCount;
            ValueMismatchCount = valueMismatchCount;
            TypeMismatchCount = typeMismatchCount;
        }

        public override string ToString()
        {
            return $"Test results: {PassCount} Pass, {ValueMismatchCount} Value Mismatch, {TypeMismatchCount} Type Mismatch";
        }
    }

    public enum TestResultType
    {
        Pass,
        TypeMismatch,
        ValueMismatch
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
                default:
                    return "// Error";
            }
        }
    }
}
