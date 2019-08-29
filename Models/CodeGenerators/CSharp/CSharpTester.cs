using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using Microsoft.CodeAnalysis.CSharp.Scripting;
using Microsoft.CodeAnalysis.Scripting;
using Microsoft.CSharp.RuntimeBinder;
using Thesis.Models.VertexTypes;
using Thesis.ViewModels;

namespace Thesis.Models.CodeGenerators.CSharp
{
    public class CSharpTester : Tester
    {
        private readonly ScriptOptions _scriptOptions = ScriptOptions.Default
            .WithImports("System", "System.Linq", "System.Collections.Generic")
            //.WithReferences(typeof(System.Linq.Enumerable).Assembly)
            // required for dynamic
            .WithReferences(typeof(RuntimeBinderException).GetTypeInfo().Assembly.Location);

        public CSharpTester(List<ClassCode> classesCode) : base(classesCode)
        {
        }

        public override async Task PerformTestAsync()
        {
            VariableToTestResultDictionary.Clear();

            var lookup = ClassesCode.ToLookup(c => c.IsStaticClass);
            var staticClasses = lookup[true];
            var normalClasses = lookup[false];

            Logger.Log(LogItemType.Info, "Initializing the Roslyn C# scripting engine...", true);

            // framework required for testing, such as the EmptyCell class
            var framework = Properties.Resources.CSharpTestingFramework;

            // create a state with all static classes initiated
            ScriptState testState;
            try
            {
                testState = await CSharpScript.RunAsync(framework, _scriptOptions);
            }
            catch (Exception ex)
            {
                Logger.Log(LogItemType.Error, "Error loading testing framework: " + ex.Message);
                return;
            }

            // test each static class separately
            foreach (var staticClass in staticClasses)
            {
                try
                {
                    var logItem2 = Logger.Log(LogItemType.Info, "Testing class " + staticClass.ClassName, true);

                    var testStaticClassState = 
                        await testState.ContinueWithAsync(staticClass.FieldsCode);
                    testStaticClassState = 
                        await testStaticClassState.ContinueWithAsync(staticClass.MethodBodyCode);

                    foreach (var testResult in VariablesToTestResults(staticClass.ClassName, testStaticClassState))
                    {
                        VariableToTestResultDictionary.Add(testResult.VariableName, testResult);
                    }
                    logItem2.AppendElapsedTime();

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
                    var logItem2 = Logger.Log(LogItemType.Info, "Testing class " + normalClass.ClassName, true);
                    var testNormalClassState = await testState.ContinueWithAsync(normalClass.FieldsCode);
                    testNormalClassState = await testNormalClassState.ContinueWithAsync(normalClass.MethodBodyCode);
                    foreach (var testResult in VariablesToTestResults(normalClass.ClassName, testNormalClassState))
                    {
                        VariableToTestResultDictionary.Add(testResult.VariableName, testResult);
                    }
                    logItem2.AppendElapsedTime();
                }
                catch (Exception ex)
                {
                    LogException(ex, normalClass.ClassName);
                }
            }
        }

        private static void LogException(Exception ex, string className)
        {
            Logger.Log(LogItemType.Error, ex.GetType().Name + " in " + className + ": " + ex.Message);
        }

        public IEnumerable<TestResult> VariablesToTestResults(string className, ScriptState state)
        {
            return state.Variables.Select(variable =>
                new TestResult(className, variable.Name, variable.Value, variable.Type));
        }

        public override TestReport GenerateTestReport(Dictionary<string, Vertex> variableNameToVertexDictionary)
        {
            Logger.Log(LogItemType.Info, "Generating test report...");

            int passCount = 0, nullCount = 0, valueMismatchCount = 0, typeMismatchCount = 0, errorCount = 0;

            foreach (var keyValuePair in variableNameToVertexDictionary)
            {
                var variableName = keyValuePair.Key;
                var vertex = keyValuePair.Value;

                if (VariableToTestResultDictionary.TryGetValue(variableName, out var testResult))
                {
                    if (vertex is RangeVertex)
                    {
                        testResult.TestResultType = TestResultType.Ignore;
                    }
                    else
                    {
                        var cellVertex = (CellVertex)keyValuePair.Value;
                        if (cellVertex.NodeType == NodeType.Constant || cellVertex.IsExternal) continue;

                        testResult.ExpectedValue = cellVertex.Value;
                        var actualValueIsNumeric = IsNumeric(testResult.ActualValue);
                        var expectedValueIsNumeric = IsNumeric(testResult.ExpectedValue);

                        if (object.ReferenceEquals(testResult.ActualValue, null))
                        {
                            testResult.TestResultType = TestResultType.Null;
                            nullCount++;
                        }
                        else if (testResult.ExpectedValue.GetType() != testResult.ActualValue.GetType() &&
                            // two different numeric types are to be treated as equal types
                            !(expectedValueIsNumeric && actualValueIsNumeric))
                        {
                            // test for strings with %, e.g. 0,02 should pass with the expected value was "2%"
                            if (cellVertex.CellType == CellType.Number &&
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
                }
                else
                {
                    // run time or compile error
                    VariableToTestResultDictionary.Add(variableName,
                        new TestResult(vertex.Class.Name, vertex.VariableName, null, null)
                        {
                            ExpectedValue = vertex is CellVertex cellVertex ? cellVertex.Value : null,
                            TestResultType = TestResultType.Error
                        });
                    errorCount++;

                }
            }

            return new TestReport(passCount, nullCount, valueMismatchCount, typeMismatchCount, errorCount);
        }

    }
}
