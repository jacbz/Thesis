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
using System.Linq;
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

        public static string RaiseFirstCharacter(this string s)
        {
            if (string.IsNullOrEmpty(s))
                return string.Empty;
            char[] a = s.ToCharArray();
            a[0] = char.ToUpper(a[0]);
            return new string(a);
        }

        public static string LowerFirstCharacter(this string s)
        {
            if (string.IsNullOrEmpty(s))
                return string.Empty;

            char[] a = s.ToCharArray();
            int i = 0;
            while (i < a.Length && char.IsUpper(a[i]))
            {
                a[i] = char.ToLower(a[i]);
                i++;
            }
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
    }
}
