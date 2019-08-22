using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Thesis.Models
{
    public static class Utility
    {
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
    }
}
