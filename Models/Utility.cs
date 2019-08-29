using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Irony.Parsing;
using Thesis.Models.VertexTypes;

namespace Thesis.Models
{
    public static class Utility
    {
        public static List<CellVertex> GetCellVertices(this IEnumerable<Vertex> vertices)
        {
            return vertices.OfType<CellVertex>().ToList();
        }

        public static string NodeToString(this ParseTreeNode node, string formula)
        {
            return formula.Substring(node.Span.Location.Column, node.Span.Length);
        }

        // e.g. A1, C2 -> [A1,A2,B1,B2,C1,C2]
        public static IEnumerable<string> AddressesInRange(string start, string end)
        {
            int startColumn = new string(start.Where(char.IsLetter).ToArray()).ToColumnNumber();
            int startRow = int.Parse(Regex.Match(start, @"\d+$").Value);

            int endColumn = new string(end.Where(char.IsLetter).ToArray()).ToColumnNumber();
            int endRow = int.Parse(Regex.Match(end, @"\d+$").Value);

            for (int column = startColumn; column <= endColumn; column++)
            for (int row = startRow; row <= endRow; row++)
                yield return column.ToExcelColumnString() + row;
        }

        // e.g. 7 -> G, 37 -> AK
        // https://stackoverflow.com/questions/48983939/convert-a-number-to-excel-s-base-26
        public static string ToExcelColumnString(this int columnNumber)
        {
            var charList = new List<char>();
            while (columnNumber > 0)
            {
                (int num, int d) = ExcelDivMod(columnNumber);
                columnNumber = num;
                charList.Add((char)(d - 1 + 'A'));
            }

            charList.Reverse();
            return new string(charList.ToArray());
        }

        // e.g. N -> 14, XT -> 644
        public static int ToColumnNumber(this string excelColumnString)
        {
            return excelColumnString.ToCharArray().Select(c => c - 'A').Aggregate(0, (r, x) => r * 26 + x + 1);
        }

        public static (int, int) ExcelDivMod(int n)
        {
            (int a, int b) = (n / 26, n % 26);
            if (b == 0)
                return (a - 1, b + 26);
            return (a, b);
        }

        public static string ToTitleCase(this string s)
        {
            if (string.IsNullOrEmpty(s))
                return string.Empty;

            s = s.ToLower();

            char[] a = s.ToCharArray();
            a[0] = char.ToUpper(a[0]);
            return new string(a);
        }

        public static string LowerFirstCharacter(this string s)
        {
            if (string.IsNullOrEmpty(s))
                return string.Empty;

            char[] a = s.ToCharArray();
            a[0] = char.ToLower(a[0]);
            return new string(a);
        }

        public static string FormatSheetName(this string sheetName)
        {
            if (sheetName.Substring(sheetName.Length - 1, 1) == "!")
                sheetName = sheetName.Substring(0, sheetName.Length - 1);
            if (sheetName.Substring(sheetName.Length - 1, 1) == "'")
                sheetName = sheetName.Substring(0, sheetName.Length - 1);
            return sheetName;
        }

        public static string ToCamelCase(this string inputString)
        {
            var output = ToPascalCase(inputString);
            if (output == "") return "";
            return LowerFirstCharacter(output);
        }

        public static string ToPascalCase(this string inputString)
        {
            if (inputString == "") return "";
            var textInfo = new CultureInfo("en-US", false).TextInfo;
            inputString = textInfo.ToTitleCase(inputString);
            inputString = inputString.MakeNameVariableConform();
            if (inputString == "") return "";
            return inputString;
        }

        public static string MakeNameVariableConform(this string inputString)
        {
            inputString = ProcessDiacritics(inputString);
            inputString = inputString.Replace("%", "Percent");
            inputString = Regex.Replace(inputString, "[^0-9a-zA-Z_]+", "");
            if (inputString.Length > 0 && char.IsDigit(inputString.ToCharArray()[0]))
            {
                if (char.IsDigit(inputString.ToCharArray()[0]))
                    inputString = "_" + inputString;
            }
            return inputString;
        }

        public static string ProcessDiacritics(this string inputString)
        {
            inputString = inputString
                .Replace("ö", "oe")
                .Replace("ä", "ae")
                .Replace("ü", "ue")
                .Replace("ß", "ss");
            var normalizedString = inputString.Normalize(NormalizationForm.FormD);
            var stringBuilder = new StringBuilder();
            foreach (var c in normalizedString)
            {
                if (CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark)
                    stringBuilder.Append(c);
            }

            return stringBuilder.ToString();
        }
    }
}
