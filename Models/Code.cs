using System.Collections.Generic;
using System.Threading.Tasks;
using Thesis.Models.CodeGeneration;
using Thesis.Models.VertexTypes;
using Thesis.ViewModels;

namespace Thesis.Models
{
    public class Code
    {
        public string SourceCode { get; }
        public Dictionary<(string className, string variableName), Vertex> VariableNameToVertexDictionary { get; }

        protected CodeGenerator CodeGenerator;
        protected Tester Tester;

        public Code(
            string sourceCode,
            Dictionary<(string className, string variableName), Vertex> variableNameToVertexDictionary,
            Tester tester)
        {
            SourceCode = sourceCode;
            VariableNameToVertexDictionary = variableNameToVertexDictionary;
            Tester = tester;
        }

        public static async Task<Code> GenerateWith(CodeGenerator codeGenerator)
        {
            Code code = await codeGenerator.GenerateCodeAsync();
            code.CodeGenerator = codeGenerator;

            return code;
        }

        public async Task<TestReport> GenerateTestReportAsync()
        {
            await Tester.PerformTestAsync();
            var testReport = Tester.GenerateTestReport(VariableNameToVertexDictionary);
            Logger.Log(LogItemType.Info, "Generating code with test results as comments...");
            testReport.Code = (await CodeGenerator.GenerateCodeAsync(Tester.TestResults)).SourceCode;
            return testReport;
        }
    }
}