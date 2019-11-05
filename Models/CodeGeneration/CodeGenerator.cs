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
    public abstract class CodeGenerator
    {
        protected Graph Graph;
        protected Dictionary<(string worksheet, string address), CellVertex> AddressToVertexDictionary;
        protected Dictionary<string, RangeVertex> RangeDictionary;
        protected Dictionary<string, Vertex> NameDictionary;

        public abstract Task<Code> GenerateCodeAsync(TestResults testResults = null);

        protected CodeGenerator(Graph graph,
            Dictionary<(string worksheet, string address), CellVertex> addressToVertexDictionary,
            Dictionary<string, RangeVertex> rangeDictionary, Dictionary<string, Vertex> nameDictionary)
        {
            Graph = graph;
            AddressToVertexDictionary = addressToVertexDictionary;
            RangeDictionary = rangeDictionary;
            NameDictionary = nameDictionary;
        }
    }
}
