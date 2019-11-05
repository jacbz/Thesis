// Thesis - An Excel to code converter
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

using System.Collections.Generic;
using System.Threading.Tasks;
using Thesis.Models.VertexTypes;

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