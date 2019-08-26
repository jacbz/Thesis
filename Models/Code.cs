using System.Collections.Generic;
using System.Threading.Tasks;
using Thesis.Models.CodeGenerators;
using Thesis.ViewModels;

namespace Thesis.Models
{
    public class Code
    {
        public string SourceCode { get; }
        public Dictionary<string, Vertex> VariableNameToVertexDictionary { get; }

        protected CodeGenerator CodeGenerator;
        protected Tester Tester;

        public Code(
            string sourceCode,
            Dictionary<string, Vertex> variableNameToVertexDictionary,
            Tester tester)
        {
            SourceCode = sourceCode;
            VariableNameToVertexDictionary = variableNameToVertexDictionary;
            Tester = tester;
        }

        public static async Task<Code> GenerateFrom(CodeGenerator codeGenerator)
        {
            Code code = await codeGenerator.GenerateCodeAsync();
            code.CodeGenerator = codeGenerator;

            return code;
        }

        public async Task<TestReport> GenerateTestReportAsync()
        {
            await Tester.PerformTestAsync();
            var testReport = Tester.GenerateTestReport(VariableNameToVertexDictionary);
            Logger.DispatcherLog(LogItemType.Info, "Generating code with test results as comments...");
            testReport.Code = (await CodeGenerator.GenerateCodeAsync(Tester.VariableToTestResultDictionary)).SourceCode;
            return testReport;
        }
    }
}