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
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Thesis.Models.VertexTypes;

namespace Thesis.Models.CodeGeneration
{
    public abstract class CodeGenerator
    {
        protected ClassCollection ClassCollection;
        protected Dictionary<(string worksheet, string address), CellVertex> AddressToVertexDictionary;
        protected Dictionary<string, RangeVertex> RangeDictionary;
        protected Dictionary<string, Vertex> NameDictionary;

        public abstract Task<Code> GenerateCodeAsync(TestResults testResults = null);
        protected abstract string[] BlockedVariableNames { get; }

        protected CodeGenerator(ClassCollection classCollection,
            Dictionary<(string worksheet, string address), CellVertex> addressToVertexDictionary, 
            Dictionary<string, RangeVertex> rangeDictionary, Dictionary<string, Vertex> nameDictionary)
        {
            ClassCollection = classCollection;
            AddressToVertexDictionary = addressToVertexDictionary;
            RangeDictionary = rangeDictionary;
            NameDictionary = nameDictionary;
        }

        protected void GenerateVariableNamesForAll()
        {
            // generate unique class names
            var blockedClassNames = BlockedVariableNames.ToHashSet();
            foreach (var generatedClass in ClassCollection.Classes)
            {
                generatedClass.Name = GenerateUniqueName(blockedClassNames, generatedClass.Name.MakeNameVariableConform());
            }

            // output vertex variable must be unique as they are in the Main class
            var blockedOutputVertexNames = new HashSet<string>();
            // generate unique variable names
            foreach (var generatedClass in ClassCollection.Classes)
            {
                var blockedVariableNames = blockedClassNames.ToHashSet();
                blockedVariableNames.UnionWith(blockedOutputVertexNames);

                foreach (var vertex in generatedClass.Vertices)
                {
                    vertex.Name = GenerateUniqueName(blockedVariableNames, vertex.Name.MakeNameVariableConform());
                    if (vertex == generatedClass.OutputVertex)
                        blockedOutputVertexNames.Add(vertex.Name);
                }
            }
        }

        protected string GenerateUniqueName(HashSet<string> usedVariableNames, string variableName)
        {
            variableName = GenerateNonDuplicateName(usedVariableNames, variableName);
            usedVariableNames.Add(variableName);
            return variableName;
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
}
