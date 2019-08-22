using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.CSharp;

namespace Thesis.Models.CodeGenerators
{
    public abstract class CodeGenerator
    {
        private protected List<GeneratedClass> GeneratedClasses;
        private protected Dictionary<string, Vertex> AddressToVertexDictionary;
        public Dictionary<string, Vertex> VariableNameToVertexDictionary { get; set; }

        private protected Tester Tester;

        public abstract Task<string> GenerateCodeAsync(Dictionary<string, TestResult> testResults = null);

        public async Task<TestReport> GenerateTestReportAsync()
        {
            await Tester.PerformTestAsync();
            var testReport = Tester.GenerateTestReport(VariableNameToVertexDictionary);
            testReport.Code = await GenerateCodeAsync(Tester.VariableToTestResultDictionary);
            return testReport;
        }

        protected CodeGenerator(List<GeneratedClass> generatedClasses, Dictionary<string, Vertex> addressToVertexDictionary)
        {
            GeneratedClasses = generatedClasses;
            AddressToVertexDictionary = addressToVertexDictionary;
        }

        /// <summary>
        /// var -> var2, var2 -> var3 etc.
        /// </summary>
        protected string GenerateNonDuplicateName(HashSet<string> usedVariableNames, string variableName)
        {
            while (usedVariableNames.Contains(variableName))
            {
                var pattern = @"\d+$";
                var matchNumberAtEnd = Regex.Match(variableName, pattern);
                if (matchNumberAtEnd.Success)
                    variableName = Regex.Replace(variableName, pattern,
                        (Convert.ToInt32(matchNumberAtEnd.Value) + 1).ToString());
                else
                    variableName += "2";
            }
            return variableName;
        }

    }

    public enum Language
    {
        CSharp
    }
}
