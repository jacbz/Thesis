﻿// Thesis - An Excel to code converter
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
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using Microsoft.CodeAnalysis.Scripting;
using Microsoft.CSharp.RuntimeBinder;
using Thesis.ViewModels;

namespace Thesis.Models.CodeGeneration.CSharp
{
    public class CSharpTester : Tester
    {
        private readonly ScriptOptions _scriptOptions = ScriptOptions.Default
            // required for dynamic
            .WithReferences(typeof(RuntimeBinderException).GetTypeInfo().Assembly.Location);

        public CSharpTester(List<ClassCode> classesCode) : base(classesCode)
        {
        }

        public override async Task PerformTestAsync()
        {
            //TestResults = new TestResults();

            //var lookup = ClassesCode.ToLookup(c => c.IsStaticClass);
            //var staticClasses = lookup[true];
            //var normalClasses = lookup[false];

            //var logItem = Logger.Log(LogItemType.Info, "Initializing the Roslyn C# scripting engine with testing framework...", true);

            //// framework required for testing, such as the EmptyCell class
            //var framework = Properties.Resources.CSharpTestingFramework;

            //// create a state with all static classes initiated
            //ScriptState testState;
            //try
            //{
            //    testState = await CSharpScript.RunAsync(framework, _scriptOptions);
            //}
            //catch (Exception ex)
            //{
            //    Logger.Log(LogItemType.Error, "Error loading testing framework: " + ex.Message);
            //    return;
            //}
            //logItem.AppendElapsedTime();

            //// test each static class separately
            //foreach (var staticClass in staticClasses)
            //{
            //    try
            //    {
            //        var logItem2 = Logger.Log(LogItemType.Info, "Testing class " + staticClass.ClassName, true);

            //        var testStaticClassState = 
            //            await testState.ContinueWithAsync(staticClass.FieldsCode);
            //        testStaticClassState = 
            //            await testStaticClassState.ContinueWithAsync(staticClass.MethodBodyCode);

            //        foreach (var testResult in VariablesToTestResults(staticClass.ClassName, testStaticClassState))
            //        {
            //            TestResults.Add(testResult.VariableName, testResult);
            //        }
            //        logItem2.AppendElapsedTime();

            //        // initialize the static classes state
            //        testState = await testState.ContinueWithAsync(staticClass.Code);
            //        if (!string.IsNullOrEmpty(staticClass.MethodBodyCode))
            //        {
            //            testState = await testState.ContinueWithAsync(staticClass.ClassName + ".Init();");
            //        }

            //    }
            //    catch (Exception ex)
            //    {
            //        LogException(ex, staticClass.ClassName);
            //    }
            //}

            //// test all normal classes on this state, separately
            //foreach (var normalClass in normalClasses)
            //{
            //    try
            //    {
            //        var logItem2 = Logger.Log(LogItemType.Info, "Testing class " + normalClass.ClassName, true);
            //        var testNormalClassState = await testState.ContinueWithAsync(normalClass.FieldsCode);
            //        testNormalClassState = await testNormalClassState.ContinueWithAsync(normalClass.MethodBodyCode);
            //        foreach (var testResult in VariablesToTestResults(normalClass.ClassName, testNormalClassState))
            //        {
            //            TestResults.Add(testResult.VariableName, testResult);
            //        }
            //        logItem2.AppendElapsedTime();
            //    }
            //    catch (Exception ex)
            //    {
            //        LogException(ex, normalClass.ClassName);
            //    }
            //}
        }

        private static void LogException(Exception ex, string className)
        {
            Logger.Log(LogItemType.Error, ex.GetType().Name + " in " + className + ": " + ex.Message);
        }

        public IEnumerable<TestResult> VariablesToTestResults(string className, ScriptState state)
        {
            return state.Variables.Select(variable =>
                new TestResult(variable.Name, variable.Value, variable.Type));
        }

        //public override TestReport GenerateTestReport(Dictionary<string, Vertex> variableNameToVertexDictionary)
        //{
        //    Logger.Log(LogItemType.Info, "Generating test report...");

        //    int passCount = 0, nullCount = 0, valueMismatchCount = 0, typeMismatchCount = 0, errorCount = 0;

        //    foreach (var keyValuePair in variableNameToVertexDictionary)
        //    {
        //        var name = keyValuePair.Key;
        //        var vertex = keyValuePair.Value;

        //        var testResult = TestResults[name];
        //        if (testResult != null)
        //        {
        //            if (vertex is RangeVertex)
        //            {
        //                testResult.TestResultType = TestResultType.Ignore;
        //            }
        //            else
        //            {
        //                var cellVertex = (CellVertex)keyValuePair.Value;
        //                if (cellVertex.Classification == Classification.Constant || cellVertex.IsExternal) continue;

        //                testResult.SetExpectedValue(cellVertex);
        //                var actualValueIsNumeric = IsNumeric(testResult.ActualValue);
        //                var expectedValueIsNumeric = IsNumeric(testResult.ExpectedValue);

        //                if (object.ReferenceEquals(testResult.ActualValue, null))
        //                {
        //                    testResult.TestResultType = TestResultType.Null;
        //                    nullCount++;
        //                }
        //                else if (testResult.ExpectedValueType == CellType.Error &&
        //                         testResult.ActualValue.GetType().Name.Contains("FormulaError"))
        //                {
        //                    testResult.Annotation = "expected " + testResult.ExpectedValue;
        //                    testResult.TestResultType = TestResultType.Pass;
        //                    passCount++;
        //                }
        //                else if (testResult.ExpectedValue.GetType() != testResult.ActualValue.GetType() &&
        //                    // two different numeric types are to be treated as equal types
        //                    !(expectedValueIsNumeric && actualValueIsNumeric))
        //                {
        //                    // test for strings with %, e.g. 0,02 should pass with the expected value was "2%"
        //                    if (cellVertex.CellType == CellType.Number &&
        //                        testResult.ExpectedValue is string percentNumber &&
        //                        percentNumber.Contains("%") &&
        //                        double.TryParse(percentNumber.Replace("%", ""), out double number))
        //                    {
        //                        if (actualValueIsNumeric && testResult.ActualValue == number * 0.01)
        //                        {
        //                            testResult.Annotation = "expected percentage but value matches that percentage";
        //                            testResult.TestResultType = TestResultType.Pass;
        //                            passCount++;
        //                            continue;
        //                        }
        //                    }

        //                    // handle EmptyCell
        //                    if (testResult.ActualValue.GetType().Name.Contains("EmptyCell"))
        //                    {
        //                        if (expectedValueIsNumeric && Equals(testResult.ActualValue, 0) ||
        //                           testResult.ExpectedValue is string s && s == "" ||
        //                           testResult.ExpectedValue is bool b && b == false)
        //                        {
        //                            testResult.Annotation = "empty cell";
        //                            testResult.TestResultType = TestResultType.Pass;
        //                            passCount++;
        //                            continue;
        //                        }
        //                    }

        //                    testResult.TestResultType = TestResultType.TypeMismatch;
        //                    typeMismatchCount++;
        //                }
        //                else if (testResult.ExpectedValue != testResult.ActualValue)
        //                {
        //                    // allow mismatch in DateTime Hour, Second as functions like TODAY() will yield different times
        //                    if (testResult.ExpectedValue is DateTime dt1 && testResult.ActualValue is DateTime dt2
        //                                                                 && dt1.Year == dt2.Year &&
        //                                                                 dt1.Month == dt2.Month && dt1.Day == dt2.Day)
        //                    {
        //                        testResult.Annotation = "ignored mismatch in hour/second";
        //                        testResult.TestResultType = TestResultType.Pass;
        //                        passCount++;
        //                        continue;
        //                    }

        //                    if (expectedValueIsNumeric && actualValueIsNumeric)
        //                    {
        //                        double epsilon = Math.Abs(Convert.ToDouble(testResult.ActualValue) -
        //                                                  Convert.ToDouble(testResult.ExpectedValue));
        //                        if (epsilon <= 0.0000001)
        //                        {
        //                            testResult.Annotation = $"difference of {epsilon:E}";
        //                            testResult.TestResultType = TestResultType.Pass;
        //                            passCount++;
        //                            continue;
        //                        }
        //                    }

        //                    testResult.TestResultType = TestResultType.ValueMismatch;
        //                    valueMismatchCount++;
        //                }
        //                else
        //                {
        //                    testResult.TestResultType = TestResultType.Pass;
        //                    passCount++;
        //                }
        //            }
        //        }
        //        else
        //        {
        //            var errorTestResult = new TestResult(vertex.Name, null, null)
        //            {
        //                TestResultType = TestResultType.Error
        //            };
        //            errorTestResult.SetExpectedValue(vertex is CellVertex cellVertex ? cellVertex : null);
        //            // run time or compile error
        //            TestResults.Add(name, errorTestResult);
        //            errorCount++;
        //        }
        //    }

        //    return new TestReport(passCount, nullCount, valueMismatchCount, typeMismatchCount, errorCount);
        //}

    }
}
