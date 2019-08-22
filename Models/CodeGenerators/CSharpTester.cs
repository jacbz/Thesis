using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.CodeAnalysis.CSharp.Scripting;
using Microsoft.CodeAnalysis.Scripting;

namespace Thesis.Models.CodeGenerators
{
    public class CSharpTester : Tester
    {
        private readonly ScriptOptions _scriptOptions = ScriptOptions.Default.WithImports("System");
        public override List<TestResult> TestClasses(List<ClassCode> classesCode)
        {
            var testResults = new List<TestResult>();

            var lookup = classesCode.ToLookup(c => c.IsSharedClass);
            var sharedClasses = lookup[true];
            var normalClasses = lookup[false];

            Task.Run(async () =>
            {
                // test each shared class separately
                foreach(var sharedClass in sharedClasses)
                {
                    var state = await CSharpScript.RunAsync(sharedClass.FieldsCode, _scriptOptions);
                    state = await state.ContinueWithAsync(sharedClass.MethodCode);
                    testResults.AddRange(VariablesToTestResults(sharedClass.ClassName, state));
                }

                // create a state with all shared classes initiated

                // test all normal classes on this state, separately
            });

            return testResults;
        }

        public IEnumerable<TestResult> VariablesToTestResults(string className, ScriptState state)
        {
            foreach(var variable in state.Variables)
            {
                yield return new TestResult(className, variable.Name, variable.Value, variable.Type);
            }
        }
    }
}
