using System;
using System.Collections.Generic;
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

        protected CodeGenerator(ClassCollection classCollection,
            Dictionary<(string worksheet, string address), CellVertex> addressToVertexDictionary, 
            Dictionary<string, RangeVertex> rangeDictionary, Dictionary<string, Vertex> nameDictionary)
        {
            ClassCollection = classCollection;
            AddressToVertexDictionary = addressToVertexDictionary;
            RangeDictionary = rangeDictionary;
            NameDictionary = nameDictionary;
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
