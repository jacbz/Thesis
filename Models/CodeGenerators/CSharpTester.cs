﻿using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.CodeAnalysis.CSharp.Scripting;
using Microsoft.CodeAnalysis.Scripting;
using Thesis.ViewModels;

namespace Thesis.Models.CodeGenerators
{
    public class CSharpTester : Tester
    {
        private readonly ScriptOptions _scriptOptions = ScriptOptions.Default.WithImports("System");

        public override async Task PerformTestAsync()
        {
            var testResults = new Dictionary<string, TestResult>();

            var lookup = ClassesCode.ToLookup(c => c.IsSharedClass);
            var sharedClasses = lookup[true];
            var normalClasses = lookup[false];

            Logger.DispatcherLog(LogItemType.Info, "Initializing Roslyn CSharp scripting engine...", true);

            // create a state with all shared classes initiated
            ScriptState withSharedClassesInitialized = await CSharpScript.RunAsync("", _scriptOptions);

            // test each shared class separately
            foreach (var sharedClass in sharedClasses)
            {
                var logItem2 = Logger.DispatcherLog(LogItemType.Info, "Testing class " + sharedClass.ClassName, true);
                var testSharedClass = CSharpScript
                    .Create(sharedClass.FieldsCode, _scriptOptions)
                    .ContinueWith(sharedClass.MethodBodyCode);

                var testSharedClassState = await testSharedClass.RunAsync();
                foreach (var testResult in VariablesToTestResults(sharedClass.ClassName, testSharedClassState))
                {
                    testResults.Add(testResult.VariableName, testResult);
                }
                logItem2.DispatcherAppendElapsedTime();

                // initialize the shared classes state
                withSharedClassesInitialized =
                    await withSharedClassesInitialized.ContinueWithAsync(sharedClass.Code);
                withSharedClassesInitialized =
                    await withSharedClassesInitialized.ContinueWithAsync(sharedClass.ClassName + ".Init();");
            }

            // test all normal classes on this state, separately
            foreach (var normalClass in normalClasses)
            {
                var logItem2 = Logger.DispatcherLog(LogItemType.Info, "Testing class " + normalClass.ClassName, true);
                var testNormalClassState =
                    await withSharedClassesInitialized.ContinueWithAsync(normalClass.FieldsCode);
                testNormalClassState =
                    await testNormalClassState.ContinueWithAsync(normalClass.MethodBodyCode);
                foreach (var testResult in VariablesToTestResults(normalClass.ClassName, testNormalClassState))
                {
                    testResults.Add(testResult.VariableName, testResult);
                }
                logItem2.DispatcherAppendElapsedTime();
            }

            VariableToTestResultDictionary = testResults;
        }

        public IEnumerable<TestResult> VariablesToTestResults(string className, ScriptState state)
        {
            return state.Variables.Select(variable => new TestResult(className, variable.Name, variable.Value, variable.Type));
        }
    }
}