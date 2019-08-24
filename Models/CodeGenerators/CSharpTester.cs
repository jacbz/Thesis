using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Microsoft.CodeAnalysis.CSharp.Scripting;
using Microsoft.CodeAnalysis.Scripting;
using Microsoft.CSharp.RuntimeBinder;
using Thesis.ViewModels;

namespace Thesis.Models.CodeGenerators
{
    public class CSharpTester : Tester
    {
        private readonly ScriptOptions _scriptOptions = ScriptOptions.Default
            .WithImports("System", "System.Linq", "System.Collections.Generic")
            //.WithReferences(typeof(System.Linq.Enumerable).Assembly)
            // required for dynamic
            .WithReferences(typeof(RuntimeBinderException).GetTypeInfo().Assembly.Location);
        public override async Task PerformTestAsync()
        {
            var testResults = new Dictionary<string, TestResult>();

            var lookup = ClassesCode.ToLookup(c => c.IsStaticClass);
            var staticClasses = lookup[true];
            var normalClasses = lookup[false];

            Logger.DispatcherLog(LogItemType.Info, "Initializing Roslyn CSharp scripting engine...", true);

            // base code required for testing, such as the EmptyCell class
            var baseCode = Properties.Resources.CSharpTestingBase;

            // create a state with all static classes initiated
            ScriptState testState = await CSharpScript.RunAsync(baseCode, _scriptOptions);

            // test each static class separately
            foreach (var staticClass in staticClasses)
            {
                try
                {
                    var logItem2 = Logger.DispatcherLog(LogItemType.Info, "Testing class " + staticClass.ClassName, true);

                    var testStaticClassState = 
                        await testState.ContinueWithAsync(staticClass.FieldsCode);
                    testStaticClassState = 
                        await testStaticClassState.ContinueWithAsync(staticClass.MethodBodyCode);

                    foreach (var testResult in VariablesToTestResults(staticClass.ClassName, testStaticClassState))
                    {
                        testResults.Add(testResult.VariableName, testResult);
                    }
                    logItem2.DispatcherAppendElapsedTime();

                    // initialize the static classes state
                    testState = await testState.ContinueWithAsync(staticClass.Code);
                    if (!string.IsNullOrEmpty(staticClass.MethodBodyCode))
                    {
                        testState = await testState.ContinueWithAsync(staticClass.ClassName + ".Init();");
                    }

                }
                catch (Exception ex)
                {
                    LogException(ex, staticClass.ClassName);
                }
            }

            // test all normal classes on this state, separately
            foreach (var normalClass in normalClasses)
            {
                try
                {
                    var logItem2 = Logger.DispatcherLog(LogItemType.Info, "Testing class " + normalClass.ClassName, true);
                    var testNormalClassState = await testState.ContinueWithAsync(normalClass.FieldsCode);
                    testNormalClassState = await testNormalClassState.ContinueWithAsync(normalClass.MethodBodyCode);
                    foreach (var testResult in VariablesToTestResults(normalClass.ClassName, testNormalClassState))
                    {
                        testResults.Add(testResult.VariableName, testResult);
                    }
                    logItem2.DispatcherAppendElapsedTime();
                }
                catch (Exception ex)
                {
                    LogException(ex, normalClass.ClassName);
                }
            }

            VariableToTestResultDictionary = testResults;
        }

        private static void LogException(Exception ex, string className)
        {
            Logger.DispatcherLog(LogItemType.Error, ex.GetType().Name + " in " + className + ": " + ex.Message);
        }

        public IEnumerable<TestResult> VariablesToTestResults(string className, ScriptState state)
        {
            return state.Variables.Select(variable =>
                new TestResult(className, variable.Name, variable.Value, variable.Type));
        }

        public override TestReport GenerateTestReport(Dictionary<string, Vertex> variableNameToVertexDictionary)
        {
            Logger.DispatcherLog(LogItemType.Info, "Generating test report...");

            int passCount = 0, valueMismatchCount = 0, typeMismatchCount = 0, skippedCount = 0;

            foreach (var keyValuePair in variableNameToVertexDictionary)
            {
                var variableName = keyValuePair.Key;
                var vertex = keyValuePair.Value;

                if (vertex.NodeType == NodeType.Constant || vertex.NodeType == NodeType.External) continue;

                if (VariableToTestResultDictionary.TryGetValue(variableName, out var testResult))
                {
                    testResult.ExpectedValue = vertex.Value;
                    var actualValueIsNumeric = IsNumeric(testResult.ActualValue);
                    var expectedValueIsNumeric = IsNumeric(testResult.ExpectedValue);

                    if (testResult.ExpectedValue.GetType() != testResult.ActualValue.GetType() &&
                        // two different numeric types are to be treated as equal types
                        !(expectedValueIsNumeric && actualValueIsNumeric))
                    {
                        // test for strings with %, e.g. 0,02 should pass with the expected value was "2%"
                        if (vertex.CellType == CellType.Number &&
                            testResult.ExpectedValue is string percentNumber &&
                            percentNumber.Contains("%") &&
                            double.TryParse(percentNumber.Replace("%", ""), out double number))
                        {
                            if (actualValueIsNumeric && testResult.ActualValue == number * 0.01)
                            {
                                testResult.Annotation = "expected percentage but value matches that percentage";
                                testResult.TestResultType = TestResultType.Pass;
                                passCount++;
                                continue;
                            }
                        }

                        // handle EmptyCell
                        if (testResult.ActualValue.GetType().Name.Contains("EmptyCell"))
                        {
                            if (expectedValueIsNumeric && Equals(testResult.ActualValue, 0) ||
                               testResult.ExpectedValue is string s && s == "" ||
                               testResult.ExpectedValue is bool b && b == false)
                            {
                                testResult.Annotation = "empty cell";
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
                        // allow mismatch in DateTime Hour, Second as functions like TODAY() will yield different times
                        if (testResult.ExpectedValue is DateTime dt1 && testResult.ActualValue is DateTime dt2
                                                                     && dt1.Year == dt2.Year &&
                                                                     dt1.Month == dt2.Month && dt1.Day == dt2.Day)
                        {
                            testResult.Annotation = "ignored mismatch in hour/second";
                            testResult.TestResultType = TestResultType.Pass;
                            passCount++;
                            continue;
                        }

                        if (expectedValueIsNumeric && actualValueIsNumeric)
                        {
                            double epsilon = Math.Abs(Convert.ToDouble(testResult.ActualValue) - 
                                                      Convert.ToDouble(testResult.ExpectedValue));
                            if (epsilon <= 0.0000001)
                            {
                                testResult.Annotation = $"difference of {epsilon:E}";
                                testResult.TestResultType = TestResultType.Pass;
                                passCount++;
                                continue;
                            }
                        }

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

    }
}
