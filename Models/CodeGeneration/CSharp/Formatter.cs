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

using System;
using System.Linq;
using System.Text.RegularExpressions;
using Syncfusion.Windows.Shared;

namespace Thesis.Models.CodeGeneration.CSharp
{
    public partial class CSharpGenerator
    {
        /// <summary>
        /// Adds line breaks for conditionals (?:) and Row.Of
        /// </summary>
        private string FormatCode(string code)
        {
            var lines = code.Split(new[] {Environment.NewLine}, StringSplitOptions.None).ToList();

            for (var i = 0; i < lines.Count; i++)
            {
                string line = lines[i];

                // ignore comments
                if (line.Contains("//")) continue;

                // match ?, :, Row.Of, but not inside quotes
                var pattern = @"(?=\?|:|Row\.Of)(?=(?:[^""]*""[^""]*"")*[^""]*$)";

                var split = Regex.Split(line, pattern).Where(s => !s.IsNullOrWhiteSpace()).ToArray();
                if (split.Length <= 1) continue;

                int numberOfSpaces = line.TakeWhile(char.IsWhiteSpace).Count();
                lines[i] = split[0];

                var nextLines = split.Skip(1).Select(s => "".PadLeft(numberOfSpaces + 4, ' ') + s).ToArray();

                // for conditionals, indent for every conditional level
                var isQuestionMark = split[1][0] == '?';
                if (isQuestionMark && nextLines.Length % 2 == 0)
                {
                    for (var j = 0; j < nextLines.Length / 2; j++)
                    {
                        var extraPadding = "".PadLeft(j * 4, ' ');
                        nextLines[j] = extraPadding + nextLines[j];
                        var endIndex = nextLines.Length - 1 - j;
                        nextLines[endIndex] = extraPadding + nextLines[endIndex];
                    }
                }

                lines.InsertRange(i + 1, nextLines);
            }

            return string.Join(Environment.NewLine, lines);
        }
    }
}