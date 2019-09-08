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
