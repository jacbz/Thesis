using System.Collections.Generic;
using System.Threading.Tasks;
using Thesis.Models.VertexTypes;
using Thesis.ViewModels;

namespace Thesis.Models.CodeGeneration
{
    public class Code
    {
        public string SourceCode { get; }
        public Dictionary<string, List<Vertex>> VariableNameToVerticesDictionary { get; }

        protected CodeGenerator CodeGenerator;
        protected Tester Tester;

        public Code(
            string sourceCode,
            Dictionary<string, List<Vertex>> variableNameToVerticesDictionary,
            Tester tester)
        {
            SourceCode = sourceCode;
            VariableNameToVerticesDictionary = variableNameToVerticesDictionary;
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
            //var testReport = Tester.GenerateTestReport(VariableNameToVerticesDictionary);
            //Logger.Log(LogItemType.Info, "Generating code with test results as comments...");
            //testReport.Code = (await CodeGenerator.GenerateCodeAsync(Tester.TestResults)).SourceCode;
            //return testReport;
            return null;
        }
    }
}