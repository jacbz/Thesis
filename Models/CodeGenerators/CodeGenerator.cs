using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.CSharp;
using Thesis.ViewModels;

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
            Logger.DispatcherLog(LogItemType.Info, "Generating code with test results as comments...");
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


        // e.g. A1, C2 -> [A1,A2,B1,B2,C1,C2]
        protected IEnumerable<string> AddressesInRange(string start, string end)
        {
            int startColumn = ExcelColumnToNumber(new string(start.Where(char.IsLetter).ToArray()));
            int startRow = int.Parse(Regex.Match(start, @"\d+$").Value);

            int endColumn = ExcelColumnToNumber(new string(end.Where(char.IsLetter).ToArray()));
            int endRow = int.Parse(Regex.Match(end, @"\d+$").Value);

            for (int column = startColumn; column <= endColumn; column++)
            for (int row = startRow; row <= endRow; row++)
                yield return NumberToExcelColumn(column) + row;
        }

        // e.g. 7 -> G, 37 -> AK
        // https://stackoverflow.com/questions/48983939/convert-a-number-to-excel-s-base-26
        protected string NumberToExcelColumn(int col)
        {
            var charList = new List<char>();
            while (col > 0)
            {
                (int num, int d) = ExcelDivMod(col);
                col = num;
                charList.Add((char)(d - 1 + 'A'));
            }

            charList.Reverse();
            return new string(charList.ToArray());
        }

        // e.g. N -> 14, XT -> 644
        protected int ExcelColumnToNumber(string column)
        {
            return column.ToCharArray().Select(c => c - 'A').Aggregate(0, (r, x) => r * 26 + x + 1);
        }

        protected (int, int) ExcelDivMod(int n)
        {
            (int a, int b) = (n / 26, n % 26);
            if (b == 0)
                return (a - 1, b + 26);
            return (a, b);
        }

    }

    public enum Language
    {
        CSharp
    }
}
